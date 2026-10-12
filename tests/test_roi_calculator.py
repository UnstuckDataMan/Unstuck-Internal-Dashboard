from app import auth


def test_roi_tool_access_is_registered():
    assert auth.tool_for_path('/roi-calculator') == 'roi_calculator'
    assert 'roi_calculator' in auth.effective_tools({'role': 'sdr'})
    assert 'roi_calculator' not in auth.effective_tools({'role': 'viewer'})


def test_roi_page_and_reporting_link(client):
    page = client.get('/roi-calculator')
    assert page.status_code == 200
    assert '/static/roi-calculator.js' in page.text
    assert 'Conversion value' in page.text
    assert 'href="/roi-calculator"' in client.get('/').text


def test_roi_requires_login_and_tool(client, monkeypatch):
    monkeypatch.delenv('AUTH_DISABLED', raising=False)
    monkeypatch.setattr(auth, 'get_session_user', lambda request: None)
    assert client.get('/roi-calculator', follow_redirects=False).status_code == 303
    monkeypatch.setattr(auth, 'get_session_user', lambda request: {'role': 'viewer', 'tools': []})
    assert client.get('/roi-calculator').status_code == 403
    assert 'href="/roi-calculator"' not in client.get('/').text
