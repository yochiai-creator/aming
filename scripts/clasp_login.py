# clasp用ログイン：start でURLを出し、finish <貼られたURL> でトークンに引き換えて ~/.clasprc.json を更新
import sys, json, os, secrets, hashlib, base64, urllib.parse, urllib.request
RC = os.path.expanduser('~/.clasprc.json'); ST = os.path.expanduser('~/.clasp_pkce.json')
t = json.load(open(RC))['tokens']['default']
CID, SEC, REDIR = t['client_id'], t['client_secret'], 'http://localhost:8888'
SCOPES = ['https://www.googleapis.com/auth/script.deployments','https://www.googleapis.com/auth/script.projects','https://www.googleapis.com/auth/script.webapp.deploy','https://www.googleapis.com/auth/drive.metadata.readonly','https://www.googleapis.com/auth/drive.file','https://www.googleapis.com/auth/service.management','https://www.googleapis.com/auth/logging.read','https://www.googleapis.com/auth/userinfo.email','https://www.googleapis.com/auth/userinfo.profile','https://www.googleapis.com/auth/cloud-platform']
if sys.argv[1] == 'start':
    v = secrets.token_urlsafe(64); st = secrets.token_urlsafe(24)
    ch = base64.urlsafe_b64encode(hashlib.sha256(v.encode()).digest()).rstrip(b'=').decode()
    json.dump({'v': v, 'state': st}, open(ST, 'w'))
    print('https://accounts.google.com/o/oauth2/v2/auth?' + urllib.parse.urlencode({'redirect_uri': REDIR, 'access_type': 'offline', 'prompt': 'consent', 'scope': ' '.join(SCOPES), 'state': st, 'code_challenge': ch, 'code_challenge_method': 'S256', 'response_type': 'code', 'client_id': CID}))
else:
    q = urllib.parse.parse_qs(urllib.parse.urlparse(sys.argv[2]).query); p = json.load(open(ST))
    assert q['state'][0] == p['state'], 'state違い（古いURL）'
    r = json.loads(urllib.request.urlopen('https://oauth2.googleapis.com/token', urllib.parse.urlencode({'code': q['code'][0], 'client_id': CID, 'client_secret': SEC, 'redirect_uri': REDIR, 'grant_type': 'authorization_code', 'code_verifier': p['v']}).encode()).read())
    d = json.load(open(RC)); tok = d['tokens']['default']
    tok.update({'access_token': r['access_token'], 'refresh_token': r.get('refresh_token', tok['refresh_token']), 'id_token': r.get('id_token', ''), 'expiry_date': int(__import__('time').time()*1000) + r['expires_in']*1000})
    json.dump(d, open(RC, 'w')); os.chmod(RC, 0o600); print('ログインOK')
