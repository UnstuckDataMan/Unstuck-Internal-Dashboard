from unittest.mock import Mock

import pytest

from app import auth
from app.utils import benchmark_live as live
from app.utils.client_benchmarks import SourceError


PAGE_PATHS = ('/expectations-setting', '/roi-calculator')
BENCHMARK_API = '/api/expectations-setting/benchmarks'


def test_roi_tool_access_is_registered():
    for path in (*PAGE_PATHS, BENCHMARK_API):
        assert auth.tool_for_path(path) == 'roi_calculator'
    assert 'roi_calculator' in auth.effective_tools({'role': 'sdr'})
    assert 'roi_calculator' not in auth.effective_tools({'role': 'viewer'})


@pytest.mark.parametrize('path', PAGE_PATHS)
def test_roi_page_and_reporting_link(client, path):
    page = client.get(path)
    assert page.status_code == 200
    assert 'Expectations Setting' in page.text
    assert '/static/roi-calculator.js' in page.text
    assert '/static/client-benchmarks.js' in page.text
    assert 'Conversion value' in page.text
    assert page.text.index('id="bench-groups"') < page.text.index('id="roi-form"')
    assert '<option value="baseline" selected>Probable' in page.text
    assert 'id="roi-baseline"' in page.text
    home = client.get('/').text
    assert 'href="/expectations-setting"' in home
    assert 'Expectations Setting' in home


def test_roi_requires_login_and_tool(client, monkeypatch):
    monkeypatch.delenv('AUTH_DISABLED', raising=False)
    monkeypatch.setattr(auth, 'get_session_user', lambda request: None)
    for path in PAGE_PATHS:
        assert client.get(path, follow_redirects=False).status_code == 303
    assert client.get(BENCHMARK_API).status_code == 401
    monkeypatch.setattr(auth, 'get_session_user', lambda request: {'role': 'viewer', 'tools': []})
    for path in PAGE_PATHS:
        assert client.get(path).status_code == 403
    assert client.get(BENCHMARK_API).status_code == 403
    assert 'href="/expectations-setting"' not in client.get('/').text


def test_roi_only_access_reads_public_benchmarks_without_granting_audit_access(client, monkeypatch):
    monkeypatch.delenv('AUTH_DISABLED', raising=False)
    monkeypatch.setattr(auth, 'get_session_user', lambda request: {
        'role': 'viewer', 'tools': ['roi_calculator'],
    })
    report = {'groups': [], 'methodology_version': 3, 'stale': False}
    get_report = Mock(return_value=report)
    monkeypatch.setattr(live, 'get_report', get_report)

    for path in PAGE_PATHS:
        assert client.get(path).status_code == 200
    response = client.get(BENCHMARK_API)
    assert response.status_code == 200
    assert response.json() == report
    assert response.headers['cache-control'] == 'no-store'
    get_report.assert_called_once_with()

    # A query parameter must never opt an ROI reader into the private report.
    get_report.reset_mock()
    assert client.get(BENCHMARK_API + '?internal=true').json() == report
    get_report.assert_called_once_with()
    get_report.reset_mock()
    for path in ('/api/client-benchmarks', '/api/client-benchmarks/audit'):
        assert client.get(path).status_code == 403
    assert client.post(BENCHMARK_API).status_code == 405
    get_report.assert_not_called()


def test_benchmark_failure_does_not_prevent_using_the_calculator(client, monkeypatch):
    monkeypatch.setattr(live, 'get_report', Mock(side_effect=SourceError('Sheet unavailable')))
    response = client.get(BENCHMARK_API)
    assert response.status_code == 503
    assert response.json() == {'error': 'Sheet unavailable'}
    assert response.headers['cache-control'] == 'no-store'
    for path in PAGE_PATHS:
        page = client.get(path)
        assert page.status_code == 200
        assert 'id="roi-form"' in page.text
        assert '/static/roi-calculator.js' in page.text
