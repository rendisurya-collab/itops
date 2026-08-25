import json

import requests

import config


class SDPError(Exception):
    pass


class SDPClient:
    """Client untuk ManageEngine ServiceDesk Plus (on-premise) REST API v3.
    Auth pakai TECHNICIAN_KEY (API key yang di-generate dari profil user)."""

    def __init__(self):
        self.base_url = config.SDP_BASE_URL
        self.api_key = config.SDP_API_KEY
        self.session = requests.Session()
        self.session.headers.update(
            {
                "TECHNICIAN_KEY": self.api_key,
                "Accept": "application/vnd.manageengine.sdp.v3+json",
            }
        )

    def _url(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _check(self, resp: requests.Response):
        if not resp.ok:
            try:
                detail = resp.json()
            except Exception:
                detail = resp.text
            raise SDPError(f"ServiceDesk Plus error {resp.status_code}: {detail}")

    def _request(self, method: str, path: str, input_data: dict = None) -> dict:
        params = {}
        if input_data is not None:
            params["input_data"] = json.dumps(input_data)
        try:
            resp = self.session.request(method, self._url(path), params=params, timeout=20)
        except requests.exceptions.RequestException as e:
            raise SDPError(f"Gagal terhubung ke ServiceDesk Plus ({e.__class__.__name__}): {e}")
        self._check(resp)
        try:
            return resp.json()
        except ValueError:
            raise SDPError(f"Respons tidak sesuai dugaan (bukan JSON): {resp.text[:200]!r}")

    def list_requests(self, row_count: int = 15, status: str = None, group=None) -> list:
        """Ambil daftar tiket/request terbaru. status (opsional) contoh 'Open'.
        group (opsional) bisa string satu nama group, atau list nama group.

        Kalau group berupa list (>1 group), query dilakukan TERPISAH per group
        lalu digabung di sini -- supaya tidak salah logika AND/OR kalau
        digabung jadi satu search_criteria (ManageEngine tidak mendukung
        pengelompokan tanda kurung, jadi 'status=Open AND groupA OR groupB'
        akan salah dibaca jadi '(status=Open AND groupA) OR groupB')."""
        groups = [group] if isinstance(group, str) else (group or [])
        groups = [g for g in groups if g]

        if len(groups) <= 1:
            return self._list_requests_single(row_count, status, groups[0] if groups else None)

        seen_ids = set()
        merged = []
        for g in groups:
            for item in self._list_requests_single(row_count, status, g):
                rid = item.get("id")
                if rid not in seen_ids:
                    seen_ids.add(rid)
                    merged.append(item)

        def _id_key(it):
            try:
                return int(it.get("id", 0))
            except (TypeError, ValueError):
                return 0

        merged.sort(key=_id_key, reverse=True)
        return merged[:row_count]

    def _list_requests_single(self, row_count: int, status: str = None, group: str = None) -> list:
        list_info = {
            "row_count": row_count,
            "start_index": 1,
            "sort_field": "created_time",
            "sort_order": "desc",
            "get_total_count": True,
            "fields_required": ["id", "subject", "status", "requester", "description", "group", "department", "category", "subcategory"],
        }
        criteria = []
        if status:
            criteria.append({"field": "status.name", "condition": "is", "value": status})
        if group:
            item = {"field": "group.name", "condition": "is", "value": group}
            if criteria:
                item["logical_operator"] = "AND"
            criteria.append(item)
        if criteria:
            list_info["search_criteria"] = criteria
        data = self._request("GET", "/api/v3/requests", {"list_info": list_info})
        return data.get("requests", [])

    def get_request(self, request_id: str) -> dict:
        data = self._request("GET", f"/api/v3/requests/{request_id}")
        return data.get("request", {})

    def get_request_notes(self, request_id: str) -> list:
        """Ambil semua notes (conversation) dari sebuah tiket, urut dari terbaru."""
        data = self._request("GET", f"/api/v3/requests/{request_id}/notes")
        return data.get("notes", [])
