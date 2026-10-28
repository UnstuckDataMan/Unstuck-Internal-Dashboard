"""Real child-process checks; fixture commands never import the application."""
import importlib.util
import os
from pathlib import Path
import signal
import subprocess
import sys
import time

import pytest


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("dashboard_supervisor", ROOT / "scripts/run_dashboard.py")
supervisor = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(supervisor)


def _wait(predicate, timeout=5):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        result = predicate()
        if result:
            return result
        time.sleep(0.02)
    raise AssertionError("Timed out waiting for fixture process")


def _alive(pid):
    try:
        # A killed orphan may remain a zombie briefly; it is no longer running.
        return Path(f"/proc/{pid}/stat").read_text().split()[2] != "Z"
    except FileNotFoundError:
        return False


def _launch(tmp_path, web_body, worker_body, **options):
    for role, body in (("web", web_body), ("worker", worker_body)):
        (tmp_path / f"{role}.py").write_text("import os,signal,time,subprocess,sys\n" + body)
    commands = [[sys.executable, str(tmp_path / f"{role}.py")] for role in ("web", "worker")]
    opts = {"shutdown_timeout": 0.2, "restart_base": 0.04, "restart_cap": 0.08, "poll_interval": 0.01, **options}
    runner = tmp_path / "runner.py"
    runner.write_text(f"import sys\nsys.path.insert(0,{str(ROOT)!r})\nfrom scripts.run_dashboard import supervise\n"
                      f"raise SystemExit(supervise({commands[0]!r},{commands[1]!r},**{opts!r}))\n")
    return subprocess.Popen([sys.executable, str(runner)], cwd=tmp_path, stdout=subprocess.DEVNULL,
                            stderr=subprocess.PIPE, text=True, start_new_session=True)


def _finish(process):
    if process.poll() is None:
        process.send_signal(signal.SIGTERM)
    try:
        process.communicate(timeout=5)
    except subprocess.TimeoutExpired:
        os.killpg(process.pid, signal.SIGKILL)
        process.communicate(timeout=2)


def test_default_mode_executes_existing_web_command_without_a_worker(monkeypatch):
    calls = []
    monkeypatch.delenv("COPY_STUDIO_EMBEDDED_WORKER", raising=False)
    monkeypatch.setenv("PORT", "10000")
    monkeypatch.setattr(supervisor.os, "execv", lambda executable, command: calls.append((executable, command)))
    monkeypatch.setattr(supervisor, "supervise", lambda *a: pytest.fail("Worker mode must be opt-in"))
    supervisor.main()
    assert calls == [(sys.executable, [sys.executable, "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "10000"])]


def test_explicit_opt_in_uses_separate_durable_worker_command(monkeypatch):
    calls = []
    monkeypatch.setenv("COPY_STUDIO_EMBEDDED_WORKER", "1")
    monkeypatch.setenv("PORT", "10000")
    monkeypatch.setattr(supervisor, "supervise", lambda *a: calls.append(a) or 42)
    assert supervisor.main() == 42
    assert calls[0][1] == [sys.executable, "-m", "app.utils.copy_studio.worker", "--schedule"]


def test_invalid_port_does_not_start_any_process(monkeypatch):
    monkeypatch.setenv("PORT", "10000; echo unsafe")
    monkeypatch.setattr(supervisor.os, "execv", lambda *a: pytest.fail("Invalid port accepted"))
    assert supervisor.main() == 1


def test_web_failure_stops_worker_and_propagates_exit_status(tmp_path):
    process = _launch(tmp_path,
        "while not os.path.exists('worker.pid'):time.sleep(.01)\nraise SystemExit(7)\n",
        "open('worker.pid','w').write(str(os.getpid()))\ntime.sleep(30)\n")
    try:
        worker_pid = int(_wait(lambda: (tmp_path / "worker.pid").read_text() if (tmp_path / "worker.pid").exists() else None))
        _, stderr = process.communicate(timeout=5)
        assert process.returncode == 7
        _wait(lambda: not _alive(worker_pid))
        assert "Web process exited with status 7" in stderr
    finally:
        _finish(process)


def test_worker_retries_back_off_and_cool_down_without_stopping_web_then_recover(tmp_path):
    process = _launch(tmp_path,
        "open('web.pid','w').write(str(os.getpid()))\ntime.sleep(30)\n",
        "with open('starts','a') as f:f.write(str(time.monotonic())+'\\n')\n"
        "if len(open('starts').read().splitlines())<=3:raise SystemExit(3)\n"
        "open('recovered','w').write(str(os.getpid()))\ntime.sleep(30)\n",
        max_worker_restarts=2, restart_cooldown=0.3)
    try:
        _wait(lambda: (tmp_path / "starts").exists() and len((tmp_path / "starts").read_text().splitlines()) == 3)
        web_pid = int((tmp_path / "web.pid").read_text())
        assert process.poll() is None and _alive(web_pid)
        time.sleep(0.1)
        assert len((tmp_path / "starts").read_text().splitlines()) == 3
        assert process.poll() is None and _alive(web_pid)
        recovered_pid = int(_wait(lambda: (tmp_path / "recovered").read_text() if (tmp_path / "recovered").exists() else None))
        times = [float(t) for t in (tmp_path / "starts").read_text().splitlines()]
        assert len(times) == 4
        assert times[1]-times[0] >= 0.04
        assert times[2]-times[1] >= 0.08
        assert times[3]-times[2] >= 0.3
        assert process.poll() is None and _alive(web_pid) and _alive(recovered_pid)
        process.send_signal(signal.SIGTERM)
        _, stderr = process.communicate(timeout=5)
        assert "cooling down while web remains running" in stderr
        assert not _alive(web_pid) and not _alive(recovered_pid)
    finally:
        _finish(process)


def test_signal_interrupts_long_worker_cooldown_without_waiting_for_retry(tmp_path):
    process = _launch(tmp_path,
        "open('web.pid','w').write(str(os.getpid()))\ntime.sleep(30)\n",
        "open('failed','w').write('yes')\nraise SystemExit(2)\n",
        max_worker_restarts=0, restart_cooldown=60)
    try:
        _wait(lambda: (tmp_path / "failed").exists())
        time.sleep(0.05)
        assert process.poll() is None
        process.send_signal(signal.SIGTERM)
        process.communicate(timeout=2)
        assert process.returncode == 128+signal.SIGTERM
        assert not _alive(int((tmp_path / "web.pid").read_text()))
    finally:
        _finish(process)


@pytest.mark.parametrize("signum", [signal.SIGTERM, signal.SIGINT])
def test_shutdown_cleans_both_process_groups_including_uncooperative_descendants(tmp_path, signum):
    body = ("signal.signal(signal.SIGTERM,signal.SIG_IGN)\n"
            "child=subprocess.Popen([sys.executable,'-c','import signal,time;signal.signal(signal.SIGTERM,signal.SIG_IGN);time.sleep(30)'])\n"
            "open('{role}.pids','w').write(str(os.getpid())+' '+str(child.pid))\ntime.sleep(30)\n")
    process = _launch(tmp_path, body.format(role="web"), body.format(role="worker"))
    try:
        _wait(lambda: all((tmp_path / f"{role}.pids").exists() for role in ("web", "worker")))
        pids = [int(p) for role in ("web", "worker") for p in (tmp_path / f"{role}.pids").read_text().split()]
        process.send_signal(signum)
        process.communicate(timeout=5)
        assert process.returncode == 128+signum
        _wait(lambda: not any(_alive(pid) for pid in pids))
    finally:
        _finish(process)


def test_previous_worker_descendants_are_gone_before_replacement(tmp_path):
    process = _launch(tmp_path, "time.sleep(30)\n",
        "if os.path.exists('old.pid'):\n"
        " old=int(open('old.pid').read())\n"
        " try:state=open('/proc/'+str(old)+'/stat').read().split()[2]\n"
        " except FileNotFoundError:state='gone'\n"
        " open('replacement','w').write(state)\n"
        " time.sleep(30)\n"
        "else:\n"
        " child=subprocess.Popen([sys.executable,'-c','import time;time.sleep(30)'])\n"
        " open('old.pid','w').write(str(child.pid))\n"
        " raise SystemExit(4)\n")
    try:
        state = _wait(lambda: (tmp_path / "replacement").read_text() if (tmp_path / "replacement").exists() else None)
        assert state in {"gone", "Z"}
    finally:
        _finish(process)
