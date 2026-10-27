"""Microsoft identity device-code flow for a public Entra client.

The token file is the one ``vdirsyncer-graph`` reads (refresh token plus
``expires_at``). CaCs runs the flow itself so the UI can show the code. The
sync subprocess never prompts.
"""
from __future__ import annotations

import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

import confgen
from graphsetup import GRAPH_SCOPES

_TENANT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9.-]{0,200}$")
DEVICE_VERIFY_URL = "https://microsoft.com/devicelogin"


class GraphAuthError(Exception):
    def __init__(self, code: str, message: str):
        self.code = code
        super().__init__(message)


def normalize_tenant(value: str | None) -> str:
    tenant = (value or "").strip() or "common"
    if not _TENANT_RE.match(tenant):
        raise ValueError(
            "Tenant must be common, organizations, consumers, a domain, or a tenant id."
        )
    return tenant


def token_path(account_id: str, token_dir: str | None = None) -> str:
    """Same path confgen writes as ``token_file`` for the contacts storage."""
    base = token_dir if token_dir is not None else confgen.TOKEN_DIR
    return os.path.join(base, f"acc_{account_id}_card.token")


def _authority(tenant: str) -> str:
    return f"https://login.microsoftonline.com/{normalize_tenant(tenant)}/oauth2/v2.0"


def _post_form(url: str, form: dict[str, str]) -> tuple[int, dict[str, Any]]:
    data = urllib.parse.urlencode(form).encode()
    req = urllib.request.Request(
        url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"}
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read()
            status = getattr(resp, "status", 200)
    except urllib.error.HTTPError as exc:
        raw = exc.read()
        status = exc.code
    if not raw:
        return status, {}
    try:
        body = json.loads(raw.decode())
    except json.JSONDecodeError as exc:
        raise GraphAuthError("bad_response", "Microsoft returned a non-JSON token response") from exc
    if not isinstance(body, dict):
        raise GraphAuthError("bad_response", "Microsoft returned an unexpected token response")
    return status, body


def request_device_code(tenant: str, client_id: str) -> dict[str, Any]:
    """Start device-code sign-in. ``device_code`` stays on the server."""
    if not (client_id or "").strip():
        raise GraphAuthError("no_client_id", "Set the Application (client) ID first.")
    status, body = _post_form(
        f"{_authority(tenant)}/devicecode",
        {"client_id": client_id.strip(), "scope": " ".join(GRAPH_SCOPES)},
    )
    if status != 200 or not body.get("device_code") or not body.get("user_code"):
        raise GraphAuthError(
            body.get("error") or "device_code_failed",
            body.get("error_description") or "Microsoft did not return a device code.",
        )
    return {
        "device_code": body["device_code"],
        "user_code": body["user_code"],
        "verification_uri": body.get("verification_uri") or DEVICE_VERIFY_URL,
        "message": body.get("message") or "",
        "interval": max(5, int(body.get("interval") or 5)),
        "expires_in": int(body.get("expires_in") or 900),
    }


def poll_device_code(tenant: str, client_id: str, device_code: str) -> dict[str, Any]:
    """One token poll. ``{"pending": True}`` or ``{"pending": False, "token": ...}``."""
    status, body = _post_form(
        f"{_authority(tenant)}/token",
        {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id.strip(),
            "device_code": device_code,
        },
    )
    if status == 200 and body.get("access_token"):
        if not body.get("refresh_token"):
            raise GraphAuthError(
                "no_refresh_token",
                "Microsoft did not return a refresh token. In the app registration, "
                "allow public client flows and include the offline_access permission, then connect again.",
            )
        return {"pending": False, "token": body}
    err = body.get("error") or ""
    if err in ("authorization_pending", "slow_down"):
        return {"pending": True, "error": err}
    raise GraphAuthError(
        err or "token_failed",
        body.get("error_description") or "Microsoft sign-in failed. Start the device code again.",
    )


def write_token(account_id: str, token: dict[str, Any], token_dir: str | None = None) -> str:
    """Write a vdirsyncer-graph token file (mode 0600) for this account."""
    path = token_path(account_id, token_dir)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    blob = dict(token)
    if blob.get("expires_in") is not None and "expires_at" not in blob:
        blob["expires_at"] = time.time() + float(blob["expires_in"])
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(blob, fh)
    try:
        os.chmod(tmp, 0o600)
    except OSError:
        pass
    os.replace(tmp, path)
    return path
