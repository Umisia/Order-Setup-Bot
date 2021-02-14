import requests
import json
import config
from logger import Logger

log = Logger(__name__).logger

def refresh_token(refresh_token):
    pars = {
        'refresh_token': refresh_token,
        'client_id': config.client_id,
        'client_secret': config.client_secret,
        'scope': config.inv_scope,
        'redirect_uri': 'http://localhost:8080/',
        'grant_type': 'refresh_token'}
    resp = requests.post("https://accounts.zoho.com/oauth/v2/token", params=pars)
    json_data = json.loads(resp.text)

    access_toks_refreshed = json_data["access_token"]

    log.info(f"status_code: {resp.status_code}")
    log.debug(access_toks_refreshed)

    return access_toks_refreshed
