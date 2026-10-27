"""Register vdirsyncer-graph and keep it contacts-only.

vdirsyncer has no storage plugin entry point. ``vdirsyncer_graph.register``
writes the private storage index. ``sitecustomize`` calls :func:`install` for
the ``vdirsyncer`` subprocess; the web process calls it too.

The pinned plugin (vdirsyncer-graph 0.1.0) always requests calendar and tasks
scopes, and a dead refresh token starts an interactive login. Both are wrong
inside this container: CaCs syncs contacts only, and stdin is closed. Patching
the module-level ``SCOPES`` list is safe on this pin because token requests
read that global at call time. Replacing ``_interactive`` turns an expired
sign-in into an error the UI can show, instead of a device-code poll that
would be killed after 300s of silence.
"""
from __future__ import annotations

import logging

log = logging.getLogger("cacs.graph")

# Must stay aligned with graphauth. The plugin's refresh sends this list.
GRAPH_SCOPES = [
    "offline_access",
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Contacts.ReadWrite",
]

_INSTALLED = False


def install() -> None:
    """Register Graph storages and apply the contacts-only runtime patches.

    Idempotent. Raises if ``vdirsyncer-graph`` is not installed.
    """
    global _INSTALLED
    import vdirsyncer_graph
    from vdirsyncer_graph.client import GraphClient, GraphError
    import vdirsyncer_graph.client as graph_client

    vdirsyncer_graph.register()
    # Mutate the list object. ``SCOPES`` is looked up by name on each request.
    graph_client.SCOPES[:] = list(GRAPH_SCOPES)

    async def _refuse_interactive(self) -> None:
        raise GraphError(
            0,
            "interaction_required",
            "Microsoft 365 sign-in expired or was revoked. "
            "Open the account in CaCs and reconnect with a device code.",
        )

    GraphClient._interactive = _refuse_interactive  # type: ignore[method-assign]
    _INSTALLED = True
    log.info("Registered graph_contacts (contacts-only scopes, no interactive login)")


def try_install() -> bool:
    """Like :func:`install`, but log and return False instead of raising."""
    try:
        install()
        return True
    except Exception:
        log.warning("Microsoft Graph storages were not registered", exc_info=True)
        return False
