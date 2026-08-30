"""Client GrabExpress API — tracking status pengiriman.

Autentikasi memakai OAuth 2.0 Client Credentials Flow:
  POST {OAUTH_URL} dengan grant_type=client_credentials + CLIENT_ID/CLIENT_SECRET
  -> dapat access_token yang di-cache sampai mendekati expired.

Endpoint tracking:
  GET {BASE_URL}/grabexpress/v1/deliveries/{deliveryID}

Kredensial & base URL diambil dari config (di-set via .env / environment).
"""

import time
import logging

import requests

import config

logger = logging.getLogger(__name__)


class GrabError(Exception):
    """Error umum saat berkomunikasi dengan API GrabExpress."""
    pass


class GrabNotFound(GrabError):
    """DeliveryID tidak ditemukan / expired (404)."""
    pass


class GrabClient:
    """Client GrabExpress dengan OAuth2 client credentials + cache token."""

    def __init__(self):
        self.client_id = config.GRAB_CLIENT_ID
        self.client_secret = config.GRAB_CLIENT_SECRET
        self.base_url = config.GRAB_BASE_URL.rstrip("/")
        self.oauth_url = config.GRAB_OAUTH_URL
        self.scope = config.GRAB_SCOPE

        self._access_token = None
        self._token_expiry = 0  # epoch detik kapan token kedaluwarsa

    def configured(self) -> bool:
        return bool(self.client_id and self.client_secret and self.base_url and self.oauth_url)

    # ------------------------------------------------------------------
    # OAUTH 2.0 CLIENT CREDENTIALS
    # ------------------------------------------------------------------
    def _get_access_token(self) -> str:
        """Ambil access_token (pakai cache kalau masih valid)."""
        now = time.time()
        # Pakai token cache kalau masih ada dan belum mendekati expired (buffer 60 dtk)
        if self._access_token and now < (self._token_expiry - 60):
            return self._access_token

        if not self.configured():
            raise GrabError("Kredensial GrabExpress belum lengkap (CLIENT_ID/SECRET/URL).")

        payload = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
        }
        if self.scope:
            payload["scope"] = self.scope

        try:
            resp = requests.post(self.oauth_url, json=payload, timeout=15)
            resp.raise_for_status()
            data = resp.json()
        except requests.exceptions.RequestException as e:
            raise GrabError(f"Gagal ambil token OAuth Grab: {e}")
        except ValueError:
            raise GrabError("Respon token OAuth Grab bukan JSON valid.")

        token = data.get("access_token")
        if not token:
            raise GrabError(f"Respon token tidak berisi access_token: {data}")

        expires_in = data.get("expires_in", 3600)
        try:
            expires_in = int(expires_in)
        except (ValueError, TypeError):
            expires_in = 3600

        self._access_token = token
        self._token_expiry = now + expires_in
        logger.info("GrabExpress: access_token baru diperoleh (berlaku %ss).", expires_in)
        return token

    def _auth_headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._get_access_token()}",
            "Accept": "application/json",
        }

    # ------------------------------------------------------------------
    # TRACKING
    # ------------------------------------------------------------------
    def get_delivery(self, delivery_id: str) -> dict:
        """GET /grabexpress/v1/deliveries/{deliveryID}.

        Return dict data delivery. Raise GrabNotFound kalau 404,
        GrabError untuk kendala lain.
        """
        if not delivery_id:
            raise GrabError("deliveryID kosong.")

        url = f"{self.base_url}/grabexpress/v1/deliveries/{delivery_id}"
        try:
            resp = requests.get(url, headers=self._auth_headers(), timeout=15)
        except requests.exceptions.RequestException as e:
            raise GrabError(f"Gagal terhubung ke GrabExpress: {e}")

        if resp.status_code == 404:
            raise GrabNotFound(delivery_id)

        # Token mungkin expired di sisi server -> refresh sekali lalu retry
        if resp.status_code == 401:
            self._access_token = None
            self._token_expiry = 0
            try:
                resp = requests.get(url, headers=self._auth_headers(), timeout=15)
            except requests.exceptions.RequestException as e:
                raise GrabError(f"Gagal terhubung ke GrabExpress: {e}")
            if resp.status_code == 404:
                raise GrabNotFound(delivery_id)

        if not resp.ok:
            try:
                detail = resp.json()
            except ValueError:
                detail = resp.text
            raise GrabError(f"GrabExpress error {resp.status_code}: {detail}")

        try:
            return resp.json()
        except ValueError:
            raise GrabError("Respon tracking GrabExpress bukan JSON valid.")
