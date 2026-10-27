"""Microsoft 365 contacts: config generation, scope patch, device-code flow.

No network and no real mailbox. Graph HTTP is mocked. vdirsyncer is only asked
to parse the generated config.
"""
from __future__ import annotations

import asyncio
import io
import json
import os
import urllib.error
import urllib.parse
from pathlib import Path

import pytest

APP = Path(__file__).resolve().parents[1] / "app"


def _pair_cfg(direction="a_to_b", service="contacts"):
    return {
        "accounts": {
            "ic": {
                "kind": "icloud",
                "name": "iCloud",
                "username": "a@b.c",
                "password": "secret-apple",
            },
            "ms": {
                "kind": "graph",
                "name": "Work",
                "client_id": "client-secret-looking",
                "tenant": "common",
            },
        },
        "pairs": {
            "p1": {
                "name": "contacts",
                "service": service,
                "a": "ic",
                "b": "ms",
                "direction": direction,
                "conflict_resolution": "a wins",
                "collections": [["people", "card", "contacts", direction, "a wins"]],
            }
        },
    }


def test_one_way_graph_config_omits_secrets_metadata_and_sets_partial_sync():
    import confgen
    from store import vpair_name

    text, env = confgen.generate(_pair_cfg(), only_pair="p1", only_collection="people")
    assert "client-secret-looking" not in text
    assert "secret-apple" not in text
    assert "graph_contacts" in text
    assert "graph_calendar" not in text
    assert "graph_tasks" not in text
    assert 'partial_sync = "ignore"' in text
    assert "metadata" not in text
    assert f'cache_dir = "{confgen.GRAPH_CACHE_DIR}"' in text
    assert 'tenant = "common"' in text
    assert "client_id.fetch" in text
    assert env["CACS_SEC_MS_CLIENT_ID"] == "client-secret-looking"
    assert env["CACS_SEC_IC_PASSWORD"] == "secret-apple"
    # read_only belongs to the iCloud storage, the source.
    icloud, graph = text.split("[storage acc_ms_card]")
    assert "read_only = true" in icloud
    assert "read_only" not in graph.split("[pair ")[0]
    assert f"[pair {vpair_name('p1', 'people')}]" in text


def test_two_way_graph_config_has_no_partial_sync_and_no_metadata():
    import confgen

    text, _env = confgen.generate(
        _pair_cfg(direction="both"), only_pair="p1", only_collection="people")
    assert "partial_sync" not in text
    assert "metadata" not in text
    assert "read_only" not in text


def test_discover_config_does_not_freeze_either_side():
    import confgen

    text, _env = confgen.generate(_pair_cfg(), only_pair="p1", list_all=True)
    assert "partial_sync" not in text
    assert "read_only" not in text
    assert "metadata" not in text


def test_google_one_way_is_unchanged():
    import confgen

    cfg = {
        "accounts": {
            "ic": {"kind": "icloud", "username": "a@b.c", "password": "pw"},
            "g": {"kind": "google", "client_id": "gid", "client_secret": "gsec"},
        },
        "pairs": {
            "p1": {
                "service": "contacts",
                "a": "ic",
                "b": "g",
                "direction": "a_to_b",
                "conflict_resolution": "a wins",
                "collections": [["people", "card", "contacts", "a_to_b", "a wins"]],
            }
        },
    }
    text, _env = confgen.generate(cfg, only_pair="p1", only_collection="people")
    assert "partial_sync" not in text
    assert 'metadata = ["displayname"]' in text
    assert "google_contacts" in text
    assert "graph_contacts" not in text


def test_graph_calendar_is_rejected():
    import confgen

    with pytest.raises(ValueError, match="calendar"):
        confgen.generate(_pair_cfg(service="calendar"), only_pair="p1", only_collection="people")


def test_bad_tenant_is_rejected():
    import confgen
    from graphauth import normalize_tenant

    with pytest.raises(ValueError):
        normalize_tenant("not a tenant")
    with pytest.raises(ValueError):
        normalize_tenant("../etc")
    assert normalize_tenant("  ") == "common"
    assert normalize_tenant("contoso.onmicrosoft.com") == "contoso.onmicrosoft.com"
    cfg = _pair_cfg()
    cfg["accounts"]["ms"]["tenant"] = "bad tenant"
    with pytest.raises(ValueError):
        confgen.generate(cfg, only_pair="p1", only_collection="people")


def test_collection_config_for_clear_and_enrich():
    import confgen
    import enrich

    built = enrich.build_collection_config(
        "ms",
        {"kind": "graph", "client_id": "cid", "tenant": "organizations"},
        "contacts",
        {"collection": "CaCs", "collection_id": "folder-1"},
    )
    assert built["type"] == "graph_contacts"
    assert built["collection"] == "CaCs"
    assert built["collection_id"] == "folder-1"
    assert built["tenant"] == "organizations"
    assert built["cache_dir"] == confgen.GRAPH_CACHE_DIR
    assert built["token_file"].endswith("acc_ms_card.token")
    assert enrich.build_collection_config(
        "ms", {"kind": "graph", "client_id": "cid"}, "calendar", {"collection": "x"}
    ) is None


def test_reconnect_message_and_token_gate(tmp_path, monkeypatch):
    import confgen
    import graphauth
    from runner import Runner

    monkeypatch.setattr(confgen, "TOKEN_DIR", str(tmp_path))
    runner = Runner(db=None, store=None)
    cfg = _pair_cfg()
    msg = runner._blocked_oauth(cfg, "p1")
    assert msg is not None
    assert "Work" in msg
    assert "reconnect" in msg
    assert "Connect Microsoft 365" in msg
    token = graphauth.token_path("ms")
    token_path = Path(token)
    token_path.write_text("{}", encoding="utf-8")
    assert runner._blocked_oauth(cfg, "p1") is None

    google = {
        "accounts": {
            "g": {"kind": "google", "name": "Gmail"},
            "ic": {"kind": "icloud", "name": "iCloud"},
        },
        "pairs": {"p1": {"a": "ic", "b": "g", "service": "contacts"}},
    }
    gmsg = runner._blocked_oauth(google, "p1")
    assert "Connect Google" in gmsg
    assert "Microsoft" not in gmsg


def test_vdirsyncer_accepts_the_generated_one_way_config(tmp_path, monkeypatch):
    import confgen
    import graphsetup
    from store import vpair_name

    graphsetup.install()
    text, env = confgen.generate(_pair_cfg(), only_pair="p1", only_collection="people")
    for key, value in env.items():
        monkeypatch.setenv(key, value)
    path = tmp_path / "vdirsyncer.conf"
    path.write_text(text, encoding="utf-8")

    from vdirsyncer.cli.config import Config
    from vdirsyncer.cli.utils import storage_names

    assert storage_names["graph_contacts"].storage_name == "graph_contacts"
    config = Config.from_filename_or_environment(str(path))
    pair = config.get_pair(vpair_name("p1", "people"))
    assert pair.partial_sync == "ignore"
    assert pair.metadata == ()
    graph = config.get_storage_args("acc_ms_card")
    assert graph["type"] == "graph_contacts"
    assert graph["client_id"] == "client-secret-looking"
    assert graph["tenant"] == "common"
    assert graph["cache_dir"] == confgen.GRAPH_CACHE_DIR
    # read_only is emitted only when true; the Outlook side of A→B stays writable.
    assert graph.get("read_only", False) is False
    icloud = config.get_storage_args("acc_ic_card")
    assert icloud["type"] == "carddav"
    assert icloud["read_only"] is True
    assert icloud["password"] == "secret-apple"


def test_sitecustomize_registers_contacts_scopes_and_refuses_interactive_login():
    import importlib.util

    import graphsetup

    spec = importlib.util.spec_from_file_location(
        "cacs_sitecustomize_under_test", APP / "sitecustomize.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    import vdirsyncer_graph.client as graph_client
    from vdirsyncer_graph.client import GraphClient, GraphError

    joined = " ".join(graph_client.SCOPES)
    assert joined == " ".join(graphsetup.GRAPH_SCOPES)
    assert "Contacts.ReadWrite" in joined
    assert "Calendars" not in joined
    assert "Tasks" not in joined
    client = GraphClient(client_id="cid", token_file="/tmp/cacs-no-such-token.json")
    with pytest.raises(GraphError, match="reconnect"):
        asyncio.run(client._interactive())


class _Response:
    def __init__(self, status, payload):
        self.status = status
        self._raw = json.dumps(payload).encode()

    def read(self):
        return self._raw

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


def test_device_code_requests_contacts_scopes_only(monkeypatch, tmp_path):
    import confgen
    import graphauth

    monkeypatch.setattr(confgen, "TOKEN_DIR", str(tmp_path))
    seen = []

    def fake_urlopen(req, timeout=30):
        seen.append(req)
        url = req.full_url
        form = urllib.parse.parse_qs(req.data.decode())
        if url.endswith("/devicecode"):
            assert "Calendars" not in form["scope"][0]
            assert "Tasks" not in form["scope"][0]
            assert "Contacts.ReadWrite" in form["scope"][0]
            assert "offline_access" in form["scope"][0]
            return _Response(200, {
                "device_code": "device-secret",
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://microsoft.com/devicelogin",
                "interval": 5,
                "expires_in": 900,
                "message": "Enter the code",
            })
        if len(seen) == 2:
            raise urllib.error.HTTPError(
                url, 400, "Bad Request", hdrs=None,
                fp=io.BytesIO(json.dumps({"error": "authorization_pending"}).encode()),
            )
        assert form["device_code"] == ["device-secret"]
        return _Response(200, {
            "access_token": "at",
            "refresh_token": "rt",
            "expires_in": 3600,
            "token_type": "Bearer",
            "scope": form.get("scope", ["contacts"])[0] if False else "User.Read",
        })

    monkeypatch.setattr(graphauth.urllib.request, "urlopen", fake_urlopen)
    started = graphauth.request_device_code("common", "cid")
    assert started["user_code"] == "ABCD-EFGH"
    assert "device_code" in started
    pending = graphauth.poll_device_code("common", "cid", started["device_code"])
    assert pending == {"pending": True, "error": "authorization_pending"}
    done = graphauth.poll_device_code("common", "cid", started["device_code"])
    path = graphauth.write_token("ms", done["token"])
    blob = json.loads(Path(path).read_text(encoding="utf-8"))
    assert blob["refresh_token"] == "rt"
    assert blob["expires_at"] > blob.get("expires_in", 0)
    assert os.stat(path).st_mode & 0o777 == 0o600


def test_api_device_flow_and_calendar_pair_rejected(monkeypatch, tmp_path):
    import confgen
    import graphauth

    monkeypatch.setattr(confgen, "TOKEN_DIR", str(tmp_path))
    calls = {"n": 0}

    def fake_urlopen(req, timeout=30):
        calls["n"] += 1
        url = req.full_url
        if url.endswith("/devicecode"):
            return _Response(200, {
                "device_code": "device-secret",
                "user_code": "WXYZ-1234",
                "verification_uri": "https://microsoft.com/devicelogin",
                "interval": 5,
                "expires_in": 900,
            })
        if calls["n"] == 2:
            raise urllib.error.HTTPError(
                url, 400, "Bad Request", hdrs=None,
                fp=io.BytesIO(json.dumps({"error": "authorization_pending"}).encode()),
            )
        return _Response(200, {
            "access_token": "at",
            "refresh_token": "rt",
            "expires_in": 3600,
            "token_type": "Bearer",
        })

    monkeypatch.setattr(graphauth.urllib.request, "urlopen", fake_urlopen)

    import main
    from fastapi.testclient import TestClient

    with TestClient(main.app) as client:
        setup = client.post("/api/setup", json={"username": "admin", "password": "secret1"})
        assert setup.status_code == 200, setup.text
        icloud = client.post("/api/accounts", json={
            "kind": "icloud", "name": "iCloud", "username": "a@b.c", "password": "pw"})
        assert icloud.status_code == 200, icloud.text
        graph = client.post("/api/accounts", json={
            "kind": "graph", "name": "Work", "client_id": "app-id", "tenant": "common"})
        assert graph.status_code == 200, graph.text
        gid = graph.json()["id"]
        iid = icloud.json()["id"]
        assert graph.json()["config"]["accounts"][gid]["client_id"] == "__SET__"
        assert graph.json()["config"]["accounts"][gid]["graph_connected"] is False

        rejected = client.post("/api/pairs", json={
            "name": "cals", "service": "calendar", "a": iid, "b": gid,
            "direction": "a_to_b", "collections": [],
        })
        assert rejected.status_code == 400
        assert "contacts" in rejected.json()["detail"]

        kept = client.post("/api/pairs", json={
            "name": "people", "service": "contacts", "a": iid, "b": gid,
            "direction": "a_to_b", "collections": [], "register": False,
        })
        assert kept.status_code == 200, kept.text

        bad = client.post("/api/accounts", json={
            "id": gid, "kind": "graph", "name": "Work",
            "client_id": "__SET__", "tenant": "bad tenant",
        })
        assert bad.status_code == 400

        started = client.post("/api/graph/device/start", json={"account": gid})
        assert started.status_code == 200, started.text
        body = started.json()
        assert body["user_code"] == "WXYZ-1234"
        assert "device_code" not in body

        pending = client.post("/api/graph/device/poll", json={"account": gid})
        assert pending.status_code == 200
        assert pending.json()["pending"] is True

        connected = client.post("/api/graph/device/poll", json={"account": gid})
        assert connected.status_code == 200, connected.text
        assert connected.json()["connected"] is True
        token = json.loads(Path(graphauth.token_path(gid)).read_text(encoding="utf-8"))
        assert token["refresh_token"] == "rt"
        view = client.get("/api/config").json()
        assert view["accounts"][gid]["graph_connected"] is True

        changed = client.post("/api/accounts", json={
            "id": gid, "kind": "graph", "name": "Work",
            "client_id": "other-app", "tenant": "common",
        })
        assert changed.status_code == 200, changed.text
        assert changed.json()["config"]["accounts"][gid]["graph_connected"] is False
        assert not Path(graphauth.token_path(gid)).exists()
