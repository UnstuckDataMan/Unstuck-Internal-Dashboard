"""Authenticated encryption for explicitly revealed client passwords."""
import base64
import json
import os

from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.hkdf import HKDF
from fastapi import HTTPException

from . import store

_KEY_CONTEXT = b'GmailSet/client-password/v1'


def _cipher():
    secret = os.getenv('GMAIL_SETUP_PASSWORD_KEY') or os.getenv('SESSION_SECRET')
    if not secret:
        raise HTTPException(503, 'Client password encryption is not configured.')
    key = HKDF(algorithm=hashes.SHA256(), length=32, salt=None, info=_KEY_CONTEXT).derive(secret.encode('utf-8'))
    return Fernet(base64.urlsafe_b64encode(key))


def seal(user, client_id, password):
    if password is not None and (not isinstance(password, str) or not 1 <= len(password) <= 1024):
        raise HTTPException(400, 'Enter a password from 1 to 1,024 characters, or remove the saved password.')
    if password is None:
        return None
    cipher = _cipher()
    payload = {'workspace': store.workspace(user), 'client_id': client_id, 'password': password}
    return cipher.encrypt(json.dumps(payload, ensure_ascii=False).encode('utf-8')).decode('ascii')


def reveal(user, client_id, ciphertext):
    cipher = _cipher()
    try:
        if not isinstance(ciphertext, str):
            raise ValueError()
        payload = json.loads(cipher.decrypt(ciphertext.encode('ascii')).decode('utf-8'))
        if not isinstance(payload, dict) or payload.get('workspace') != store.workspace(user) or payload.get('client_id') != client_id:
            raise ValueError()
        password = payload.get('password')
        if not isinstance(password, str) or not 1 <= len(password) <= 1024:
            raise ValueError()
        return password
    except (InvalidToken, ValueError, TypeError, UnicodeError):
        raise HTTPException(503, 'The saved password could not be read. Check the configured encryption key.') from None


def scrub_public_clients(state):
    for client in state.get('clients', []):
        ciphertext = client.get('password_ciphertext')
        for key in list(client):
            if 'password' in key.casefold() and key != 'password_updated_at':
                client.pop(key)
        client['has_password'] = isinstance(ciphertext, str) and bool(ciphertext)
