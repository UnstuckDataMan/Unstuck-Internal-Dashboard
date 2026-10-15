from app.utils import client_benchmarks as benchmarks


def test_only_benchmark_tool_is_offline(client, monkeypatch):
    def unexpected_read():
        raise AssertionError('Offline tool must not read Google Sheets')
    monkeypatch.setattr(benchmarks, 'load_source', unexpected_read)
    for path in ['/client-benchmarks', '/api/client-benchmarks', '/api/client-benchmarks/activity-template']:
        response = client.get(path)
        assert response.status_code == 503
        assert response.json()['detail'] == 'Average Client Data is temporarily offline.'
        assert response.headers['cache-control'] == 'no-store'
    response = client.post('/api/client-benchmarks/upload', files={'file': ('test.xlsx', b'not parsed')})
    assert response.status_code == 503
    home = client.get('/')
    assert home.status_code == 200
    assert 'href="/client-benchmarks"' not in home.text
    assert 'href="/roi-calculator"' in home.text
    assert client.get('/roi-calculator').status_code == 200
