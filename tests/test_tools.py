"""Tool behavior through the MCP client: errors, paging, mentions, message links.

Tools run in-process via `mcp.Client(server.mcp)`. Graph is a real GraphClient whose
HTTP transport is an httpx.MockTransport fake, and sign-in is either a stub or a real
AuthManager over a fake MSAL app, so no test touches the network or the token cache.
"""

import json
import os

import httpx
import pytest
from mcp import Client

os.environ.setdefault("TEAMS_MCP_TENANT_ID", "test-tenant")
os.environ.setdefault("TEAMS_MCP_CLIENT_ID", "test-client")

from teams_mcp import server  # noqa: E402
from teams_mcp.auth import AuthManager  # noqa: E402
from teams_mcp.graph import GRAPH_BASE, GraphClient  # noqa: E402

CHAT_ID = (
    "19:5f4e2a10-aaaa-4bbb-8ccc-000000000001_9c8b7a60-dddd-4eee-8fff-000000000002"
    "@unq.gbl.spaces"
)


class StubAuth:
    """Stands in for AuthManager when a test only needs signed in or signed out."""

    def __init__(self, signed_in: bool = True):
        self.signed_in = signed_in

    def is_authenticated(self) -> bool:
        return self.signed_in


class RecordingGraph:
    """httpx handler: records every request and answers with `respond(request)`."""

    def __init__(self, respond=None):
        self.requests: list[httpx.Request] = []
        self._respond = respond or (lambda request: httpx.Response(500))

    def __call__(self, request: httpx.Request) -> httpx.Response:
        self.requests.append(request)
        return self._respond(request)


def graph_error(status: int, code: str, message: str) -> httpx.Response:
    return httpx.Response(status, json={"error": {
        "code": code,
        "message": message,
        "innerError": {
            "date": "2026-10-08T09:00:00",
            "request-id": "4b1c4f4e-0000-4000-8000-000000000001",
            "client-request-id": "4b1c4f4e-0000-4000-8000-000000000001",
        },
    }})


@pytest.fixture
def install(monkeypatch):
    """Install sign-in and a GraphClient backed by an httpx handler for one test."""

    def _install(handler, *, auth=None, token: str | None = "test-token") -> GraphClient:
        graph = GraphClient(token_provider=lambda: token)
        graph._http = httpx.AsyncClient(base_url=GRAPH_BASE, transport=httpx.MockTransport(handler))
        monkeypatch.setattr(server, "auth", auth if auth is not None else StubAuth())
        monkeypatch.setattr(server, "graph", graph)
        monkeypatch.setattr(server, "_pending_flow", None)
        return graph

    return _install


async def call(name: str, arguments: dict | None = None):
    async with Client(server.mcp) as client:
        return await client.call_tool(name, arguments or {})


def text(result) -> str:
    return result.content[0].text


# --- errors reach the client --------------------------------------------------


async def test_signed_out_call_reports_not_authenticated(install):
    install(RecordingGraph(), auth=StubAuth(signed_in=False))

    result = await call("list_teams")

    assert result.is_error
    assert "Not authenticated" in text(result)


async def test_missing_token_at_request_time_reports_not_authenticated(install):
    graph = RecordingGraph()
    install(graph, token=None)

    result = await call("list_teams")

    assert result.is_error
    assert "Not authenticated" in text(result)
    assert graph.requests == []


async def test_graph_error_reaches_client(install):
    install(RecordingGraph(lambda request: graph_error(
        403, "Forbidden", "Insufficient privileges to complete the operation.",
    )))

    result = await call("list_teams")

    assert result.is_error
    assert "Forbidden" in text(result)
    assert "Insufficient privileges to complete the operation." in text(result)


@pytest.mark.parametrize("tool, arguments", [
    ("delete_message", {"message_id": "1759838400000"}),
    ("update_message", {"message_id": "1759838400000", "content": "fixed"}),
    ("set_reaction", {"message_id": "1759838400000", "reaction": "like"}),
    ("unset_reaction", {"message_id": "1759838400000", "reaction": "like"}),
    ("download_attachment", {"message_id": "1759838400000", "hosted_content_id": "aWQ9"}),
])
async def test_dual_context_tool_without_target_is_error(install, tool, arguments):
    graph = RecordingGraph()
    install(graph)

    result = await call(tool, arguments)

    assert result.is_error
    assert "chat_id" in text(result)
    assert graph.requests == []


async def test_group_chat_with_one_other_member_is_error(install):
    install(RecordingGraph())

    result = await call("create_group_chat", {"member_emails": "ann@example.com"})

    assert result.is_error
    assert "at least 2" in text(result)


async def test_complete_login_without_pending_login_is_error(install):
    install(RecordingGraph())

    result = await call("complete_login")

    assert result.is_error
    assert "login" in text(result)


class FakeMsalApp:
    """msal.PublicClientApplication without the network: no cached account, scripted flows."""

    device_flow: dict = {}
    device_flow_result: dict = {}

    def __init__(self, client_id, authority=None, token_cache=None, **kwargs):
        pass

    def get_accounts(self, username=None):
        return []

    def acquire_token_silent(self, scopes, account, **kwargs):
        return None

    def initiate_device_flow(self, scopes=None, **kwargs):
        return dict(self.device_flow)

    def acquire_token_by_device_flow(self, flow, **kwargs):
        return dict(self.device_flow_result)


DEVICE_FLOW = {
    "device_code": "DAQABAAEAAAD--DLA3VO7QrddgJg7WevrUCV",
    "user_code": "F7KQ2XRTN",
    "verification_uri": "https://microsoft.com/devicelogin",
    "expires_in": 900,
    "interval": 5,
    "message": (
        "To sign in, use a web browser to open the page https://microsoft.com/devicelogin "
        "and enter the code F7KQ2XRTN to authenticate."
    ),
    "expires_at": 1759840000.0,
    "_correlation_id": "7d1a3c2e-0000-4000-8000-000000000003",
}


def oauth_error(error: str, description: str, code: int) -> dict:
    return {
        "error": error,
        "error_description": description,
        "error_codes": [code],
        "timestamp": "2026-10-08 09:00:00Z",
        "trace_id": "0f1e2d3c-0000-4000-8000-000000000004",
        "correlation_id": "7d1a3c2e-0000-4000-8000-000000000003",
        "error_uri": "https://login.microsoftonline.com/error?code=" + str(code),
    }


@pytest.fixture
def msal_app(monkeypatch, tmp_path):
    """A real AuthManager over FakeMsalApp, with its token cache in tmp_path."""
    monkeypatch.setattr("teams_mcp.auth.msal.PublicClientApplication", FakeMsalApp)

    def _make(device_flow: dict, device_flow_result: dict | None = None) -> AuthManager:
        monkeypatch.setattr(FakeMsalApp, "device_flow", device_flow)
        monkeypatch.setattr(FakeMsalApp, "device_flow_result", device_flow_result or {})
        return AuthManager(
            tenant_id="test-tenant", client_id="test-client", cache_dir=str(tmp_path),
        )

    return _make


async def test_failed_device_flow_reason_reaches_client(install, msal_app):
    install(RecordingGraph(), auth=msal_app(oauth_error(
        "invalid_client",
        "AADSTS7000218: The request body must contain the following parameter: "
        "'client_assertion' or 'client_secret'.",
        7000218,
    )))

    result = await call("login")

    assert result.is_error
    assert "AADSTS7000218" in text(result)


async def test_failed_sign_in_reason_reaches_client(install, msal_app):
    install(RecordingGraph(), auth=msal_app(DEVICE_FLOW, oauth_error(
        "expired_token",
        "AADSTS70020: The provided value for the input parameter 'device_code' is not valid. "
        "This device code has expired.",
        70020,
    )))

    started = await call("login")
    assert not started.is_error
    assert json.loads(text(started))["user_code"] == "F7KQ2XRTN"

    result = await call("complete_login")

    assert result.is_error
    assert "AADSTS70020" in text(result)


async def test_tools_keep_their_parameter_schemas():
    tools = {tool.name: tool for tool in await server.mcp.list_tools()}

    send = tools["send_chat_message"]
    assert set(send.input_schema["properties"]) == {"chat_id", "content", "mentions", "reply_to"}
    assert send.input_schema["required"] == ["chat_id", "content"]
    assert send.description.startswith("Send a message to a Teams chat.")
    assert tools["login"].input_schema["properties"] == {}
