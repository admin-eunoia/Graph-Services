import time
import re
from typing import Optional, Dict, Any, Tuple
import requests
from requests import exceptions as requests_exceptions
from urllib.parse import quote
from Services.excel_render import fill_cells_in_memory, EXCEL_MIME
from validators.payload import validate_cell_map_or_raise

# -----------------------------
# Validadores compartidos
# -----------------------------
_CELL_RE = re.compile(r"^(?:([^!]+)!)?([A-Za-z]+[1-9][0-9]*)$")  # [sheet!]ColRow


class GraphAPIError(Exception):
    def __init__(self, status_code: int, message: str, ms_request_id: Optional[str] = None, response_body: Optional[str] = None):
        super().__init__(message)
        self.status_code = status_code
        self.message = message
        self.ms_request_id = ms_request_id
        self.response_body = response_body

    def __str__(self) -> str:
        base = self.message or "Graph API error"
        if self.ms_request_id:
            return f"{base} (status={self.status_code}, ms-request-id={self.ms_request_id})"
        return f"{base} (status={self.status_code})"

# -----------------------------
# Core Graph client with retry
# -----------------------------
class GraphServices:
    def __init__(self, access_token: str, correlation_id: Optional[str] = None, graph_url: str = "https://graph.microsoft.com/v1.0"):
        self.access_token = access_token
        self.graph_url = graph_url
        self.correlation_id = correlation_id  # our own request-id for logs/propagation

    def _headers(self, content_type: str = "application/json") -> Dict[str, str]:
        h = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": content_type,
        }
        # Forward correlation ID to help correlate in your logs (custom header)
        if self.correlation_id:
            h["X-Correlation-ID"] = self.correlation_id
        return h

    # ---------- low-level request with retry/backoff ----------
    def _request_with_retry(self, method: str, url: str, *, expected: Tuple[int, ...] = (200, 201, 204), headers: Optional[Dict[str, str]] = None, **kwargs) -> Tuple[requests.Response, Optional[str]]:
        """
        Centralized HTTP call with:
          - exponential backoff on 423/429/502/503/504
          - honor Retry-After if present
          - returns (response, ms_graph_request_id_header)
        """
        max_attempts = 5
        base_delay = 0.6
        hdrs = headers or self._headers()
        last_ms_req_id = None

        resp: Optional[requests.Response] = None
        last_exception: Optional[requests_exceptions.RequestException] = None

        for attempt in range(1, max_attempts + 1):
            try:
                resp = requests.request(method, url, headers=hdrs, timeout=60, **kwargs)
                last_exception = None
            except requests_exceptions.RequestException as exc:
                last_exception = exc
                delay = base_delay * attempt
                time.sleep(delay)
                continue
            last_ms_req_id = resp.headers.get("request-id") or resp.headers.get("x-ms-request-id")

            if resp.status_code in expected:
                return resp, last_ms_req_id

            # Retryable statuses
            if resp.status_code in (423, 429, 502, 503, 504):
                # Honoring Retry-After if provided
                ra = resp.headers.get("Retry-After")
                if ra:
                    try:
                        delay = float(ra)
                    except ValueError:
                        delay = base_delay * attempt
                else:
                    delay = base_delay * attempt
                time.sleep(delay)
                continue

            # Non-retryable -> raise
            raise GraphAPIError(
                status_code=resp.status_code,
                message=f"{resp.status_code} {resp.reason}",
                ms_request_id=last_ms_req_id,
                response_body=resp.text,
            )

        # If we exit loop, last resp failed repeatedly
        if resp is not None:
            raise GraphAPIError(
                status_code=resp.status_code,
                message=f"Max retries exceeded ({resp.status_code} {resp.reason})",
                ms_request_id=last_ms_req_id,
                response_body=resp.text,
            )
        if last_exception is not None:
            raise GraphAPIError(
                status_code=0,
                message=f"Error de red al contactar Graph: {last_exception}",
                ms_request_id=last_ms_req_id,
            ) from last_exception
        raise GraphAPIError(status_code=500, message="Max retries exceeded", ms_request_id=last_ms_req_id)

    # ---------- high-level helpers ----------
    def download_file_bytes(self, full_path: str, target_user_id: str = None, drive_id: str = None) -> Tuple[bytes, Optional[str]]:
        full_path_enc = "/".join(quote(p) for p in full_path.split("/"))
        if drive_id:
            url = f"{self.graph_url}/drives/{drive_id}/root:/{full_path_enc}:/content"
        elif target_user_id:
            url = f"{self.graph_url}/users/{target_user_id}/drive/root:/{full_path_enc}:/content"
        else:
            raise ValueError("Debes pasar target_user_id o drive_id")

        resp, ms_id = self._request_with_retry("GET", url, expected=(200,), headers=self._headers())
        return resp.content, ms_id

    def upload_file_bytes(self, file_bytes: bytes, dest_path: str, conflict_behavior: str = "fail", target_user_id: str = None, drive_id: str = None) -> Tuple[dict, Optional[str]]:
        dest_path_enc = "/".join(quote(p) for p in dest_path.split("/"))
        if drive_id:
            url = f"{self.graph_url}/drives/{drive_id}/root:/{dest_path_enc}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"
        elif target_user_id:
            url = f"{self.graph_url}/users/{target_user_id}/drive/root:/{dest_path_enc}:/content?@microsoft.graph.conflictBehavior={conflict_behavior}"
        else:
            raise ValueError("Debes pasar target_user_id o drive_id")

        resp, ms_id = self._request_with_retry(
            "PUT", url,
            expected=(200, 201),
            headers=self._headers(EXCEL_MIME),
            data=file_bytes
        )
        return resp.json(), ms_id

    def _resolve_item_id(self, full_path: str, *, target_user_id: str = None, drive_id: str = None) -> Tuple[str, Optional[str]]:
        full_path_enc = "/".join(quote(p) for p in full_path.split("/"))
        if drive_id:
            url = f"{self.graph_url}/drives/{drive_id}/root:/{full_path_enc}"
        elif target_user_id:
            url = f"{self.graph_url}/users/{target_user_id}/drive/root:/{full_path_enc}"
        else:
            raise ValueError("Debes pasar target_user_id o drive_id")
        resp, ms_id = self._request_with_retry("GET", url, expected=(200,), headers=self._headers())
        return resp.json()["id"], ms_id

    def _resolve_worksheets(self, *, item_id: str, target_user_id: str = None, drive_id: str = None) -> Tuple[list[dict], Optional[str]]:
        if drive_id:
            url = f"{self.graph_url}/drives/{drive_id}/items/{item_id}/workbook/worksheets?$select=id,name"
        else:
            url = f"{self.graph_url}/users/{target_user_id}/drive/items/{item_id}/workbook/worksheets?$select=id,name"
        resp, ms_id = self._request_with_retry("GET", url, expected=(200,), headers=self._headers())
        return resp.json().get("value", []), ms_id

    def write_cells_graph(self, *, full_dest_path: str, data: dict, target_user_id: str = None, drive_id: str = None) -> Tuple[dict, Dict[str, str]]:
        # We assume data already validated by routes
        item_id, ms_resolve_id = self._resolve_item_id(full_dest_path, target_user_id=target_user_id, drive_id=drive_id)
        sheets, ms_ws_id = self._resolve_worksheets(item_id=item_id, target_user_id=target_user_id, drive_id=drive_id)
        if not sheets:
            raise Exception("El workbook no tiene hojas.")

        # Build sheet map for fast lookup
        by_name = {s.get("name"): s.get("id") for s in sheets}
        default_ws_id = sheets[0]["id"]
        ms_ids_accum = {"resolve_item": ms_resolve_id, "list_sheets": ms_ws_id}
        results = {}

        for key, value in data.items():
            m = _CELL_RE.match(key)
            if not m:
                results[key] = {
                    "status": "error",
                    "message": "Dirección inválida",
                    "http_status": None,
                }
                continue

            sheet_name, addr = m.group(1), m.group(2)
            ws_id = by_name.get(sheet_name) if sheet_name else default_ws_id
            if not ws_id:
                results[key] = {
                    "status": "error",
                    "message": f"Hoja '{sheet_name}' no encontrada",
                    "http_status": 404,
                }
                continue

            if drive_id:
                base = f"{self.graph_url}/drives/{drive_id}/items/{item_id}"
            else:
                base = f"{self.graph_url}/users/{target_user_id}/drive/items/{item_id}"

            url = f"{base}/workbook/worksheets/{ws_id}/range(address='{addr}')"
            try:
                resp, ms_patch_id = self._request_with_retry(
                    "PATCH",
                    url,
                    expected=(200,),
                    headers=self._headers(),
                    json={"values": [[value]]},
                )
                ms_ids_accum[f"patch_{key}"] = ms_patch_id
                results[key] = {"status": "ok"}
            except GraphAPIError as ge:
                ms_ids_accum[f"patch_{key}"] = ge.ms_request_id
                results[key] = {
                    "status": "error",
                    "message": ge.message,
                    "http_status": ge.status_code,
                    "ms_request_id": ge.ms_request_id,
                }
            except Exception as err:
                results[key] = {
                    "status": "error",
                    "message": str(err),
                    "http_status": None,
                }

        return {"written": results}, ms_ids_accum

    # ---------- in-memory Excel render ----------
    def render_in_memory(self, template_bytes: bytes, data: dict) -> bytes:
        out = fill_cells_in_memory(template_bytes, data)
        return out.getvalue()
