import requests

def token():
    r = requests.post('https://api.snov.io/v1/oauth/access_token', data={
        'grant_type': 'client_credentials',
        'client_id': 'e732fbfb08bc827191f4ba1be3300441',
        'client_secret': 'c0e9ed343afaaa7084cb3b7f42fe4bc7'
    })
    return r.json()['access_token']

t = token()
campanhas = requests.get('https://api.snov.io/v1/get-user-campaigns', params={'access_token': t}).json()
ids = [str(c.get('id') or c.get('campaign_id')) for c in campanhas if c.get('id') or c.get('campaign_id')]
print(f"Campanhas: {len(ids)} → {ids}")

for label, de, ate in [
    ('ATUAL',    '2026-03-10', '2026-04-08'),
    ('ANTERIOR', '2026-01-27', '2026-02-10'),
    ('GERAL',    '2025-10-01', '2026-04-08'),
]:
    r = requests.get('https://api.snov.io/v2/statistics/campaign-analytics',
        params={'access_token': t, 'campaign_id': ','.join(ids), 'date_from': de, 'date_to': ate})
    d = r.json()
    print(f"\n{label} ({de} → {ate})")
    for k in ['total_contacted','emails_sent','email_opens','email_opens_rate',
              'email_replies','email_replies_rate','interested','maybe','auto_replied']:
        print(f"  {k}: {d.get(k,0)}")