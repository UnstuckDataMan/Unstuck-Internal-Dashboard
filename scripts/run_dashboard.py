"""Run the dashboard, optionally supervising its separate durable Studio worker.

COPY_STUDIO_EMBEDDED_WORKER=1 opts into sharing this Render instance. The
default replaces this launcher with the existing web process. No job runs in
the web process, and the database remains the durable source of job ownership.
"""
from __future__ import annotations

import logging
import os
import signal
import subprocess
import sys
import time

logger = logging.getLogger("dashboard_supervisor")


def _group_alive(process):
    process.poll()  # Reap an exited group leader before checking descendants.
    try:
        os.killpg(process.pid, 0)
        return True
    except ProcessLookupError:
        return False


def _signal_group(process, signum):
    try:
        os.killpg(process.pid, signum)
    except ProcessLookupError:
        pass


def _stop(processes, timeout):
    """Terminate and reap whole child groups, including an exited leader's children."""
    processes = [p for p in processes if p is not None]
    for process in processes:
        _signal_group(process, signal.SIGTERM)
    deadline = time.monotonic() + timeout
    while any(_group_alive(p) for p in processes) and time.monotonic() < deadline:
        time.sleep(0.05)
    for process in processes:
        # A grandchild may outlive its group leader. Always clean the group
        # before allowing any replacement worker to start.
        _signal_group(process, signal.SIGKILL)
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            logger.error("Child process did not exit after termination")


def supervise(web_command, worker_command, *, shutdown_timeout=20,
              max_worker_restarts=3, restart_base=1, restart_cap=15,
              restart_cooldown=300, poll_interval=0.1):
    """Supervise two argv lists. Dependency injection is for offline process tests."""
    if not web_command or not worker_command:
        raise ValueError("Both web and worker command lists are required.")
    if (shutdown_timeout < 0 or max_worker_restarts < 0 or restart_base < 0
            or restart_cap < 0 or restart_cooldown <= 0 or poll_interval <= 0):
        raise ValueError("Supervisor timeouts and restart limits must be non-negative.")
    stopping = []
    previous = {}
    web = worker = None
    restart_at, restarts = 0.0, 0

    def stop(signum, _frame):
        if not stopping:
            stopping.append(signum)

    def start(command, role):
        try:
            process = subprocess.Popen(list(command), start_new_session=True)
            logger.info("%s process started", role)
            return process
        except OSError:
            # No argv, environment, provider configuration or exception body.
            logger.error("%s process could not start", role)
            return None

    def retry_worker():
        nonlocal restarts, restart_at
        restarts += 1
        if restarts > max_worker_restarts:
            # A Studio configuration/source failure must not restart unrelated
            # dashboard tools. Pause worker retries, then attempt recovery.
            logger.error("Worker retry limit reached; cooling down while web remains running")
            restarts = 0
            delay = restart_cooldown
        else:
            delay = min(restart_cap, restart_base * 2 ** (restarts-1))
        restart_at = time.monotonic() + delay

    try:
        for signum in (signal.SIGTERM, signal.SIGINT):
            previous[signum] = signal.signal(signum, stop)
        web = start(web_command, "Web")
        if web is None:
            return 1
        while not stopping:
            web_status = web.poll()
            if web_status is not None:
                logger.error("Web process exited with status %s", web_status)
                return web_status if web_status > 0 else 128-web_status if web_status < 0 else 1
            if worker is not None and worker.poll() is not None:
                logger.warning("Worker process exited with status %s", worker.returncode)
                _stop([worker], shutdown_timeout)
                worker = None
                retry_worker()
            if worker is None and time.monotonic() >= restart_at and not stopping:
                worker = start(worker_command, "Worker")
                if worker is None:
                    retry_worker()
            time.sleep(poll_interval)
        logger.info("Supervisor received shutdown signal")
        return 128 + stopping[0]
    finally:
        _stop([worker, web], shutdown_timeout)
        for signum, handler in previous.items():
            signal.signal(signum, handler)


def main():
    logging.basicConfig(level=logging.INFO)
    port = os.environ.get("PORT", "8000")
    if not port.isdigit() or not 1 <= int(port) <= 65535:
        logger.error("PORT must be a valid TCP port")
        return 1
    web = [sys.executable, "-m", "uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", port]
    if os.environ.get("COPY_STUDIO_EMBEDDED_WORKER") != "1":
        os.execv(sys.executable, web)
        return 1  # execv only returns if replaced by a test fixture.
    worker = [sys.executable, "-m", "app.utils.copy_studio.worker", "--schedule"]
    return supervise(web, worker)


if __name__ == "__main__":
    raise SystemExit(main())
