"""Tool behavior through the MCP client: errors, paging, mentions, message links.

Tools run in-process via `mcp.Client(server.mcp)`. Graph is a real GraphClient whose
HTTP transport is an httpx.MockTransport fake, and sign-in is either a stub or a real
AuthManager over a fake MSAL app, so no test touches the network or the token cache.
"""

import json
import os
import re
import time
from datetime import datetime, timedelta

import httpx
import pytest
from mcp import Client

os.environ.setdefault("TEAMS_MCP_TENANT_ID", "test-tenant")
os.environ.setdefault("TEAMS_MCP_CLIENT_ID", "test-client")

from teams_mcp import server  # noqa: E402
from teams_mcp.auth import AuthManager  # noqa: E402
from teams_mcp.graph import GRAPH_BASE, MAX_PAGES, GraphClient  # noqa: E402

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


# --- chat history paging ------------------------------------------------------

USER_ANN = {
    "@odata.type": "#microsoft.graph.teamworkUserIdentity",
    "id": "8b081ef6-4792-4def-b2c9-c363a1bf41d5",
    "displayName": "Ann Lee",
    "userIdentityType": "aadUser",
    "tenantId": "2432b57b-0abd-43db-aa7b-16eadd115d34",
}


def chat_message(created: str, *, modified: str | None = None, system: bool = False) -> dict:
    """A chatMessage as Graph v1.0 returns it; a chat message id is its creation time in ms."""
    message_id = str(int(datetime.fromisoformat(created).timestamp() * 1000))
    return {
        "id": message_id,
        "replyToId": None,
        "etag": message_id,
        "messageType": "systemEventMessage" if system else "message",
        "createdDateTime": created,
        "lastModifiedDateTime": modified or created,
        "lastEditedDateTime": None,
        "deletedDateTime": None,
        "subject": None,
        "summary": None,
        "chatId": CHAT_ID,
        "importance": "normal",
        "locale": "en-us",
        "webUrl": None,
        "channelIdentity": None,
        "policyViolation": None,
        "eventDetail": {
            "@odata.type": "#microsoft.graph.membersAddedEventMessageDetail",
            "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
            "members": [{"id": USER_ANN["id"], "displayName": None, "userIdentityType": "aadUser"}],
            "initiator": {"application": None, "device": None, "user": USER_ANN},
        } if system else None,
        "from": None if system else {"application": None, "device": None, "user": USER_ANN},
        "body": {
            "contentType": "html",
            "content": "<systemEventMessage/>" if system else f"<p>sent at {created}</p>",
        },
        "attachments": [],
        "mentions": [],
        "reactions": [],
    }


def every_minute(count: int, start: str = "2026-10-07T08:00:00.000Z") -> list[dict]:
    first = datetime.fromisoformat(start)
    return [
        chat_message((first + timedelta(minutes=i)).strftime("%Y-%m-%dT%H:%M:%S.000Z"))
        for i in range(count)
    ]


class FakeChatMessages(RecordingGraph):
    """GET /chats/{id}/messages as Graph v1.0 documents it: $top at most 50; $orderby
    lastModifiedDateTime desc (the default) or createdDateTime desc; $filter
    "createdDateTime lt <time>", ignored unless $orderby sorts by createdDateTime; further
    pages through @odata.nextLink. `page_size` lets a page hold fewer than $top items;
    `next_link_drops_orderby` models a nextLink that loses $orderby, which the docs do
    not rule out.
    """

    def __init__(
        self, messages: list[dict], page_size: int = 50, next_link_drops_orderby: bool = False,
    ):
        super().__init__(self._page)
        self.messages = messages
        self.page_size = page_size
        self.next_link_drops_orderby = next_link_drops_orderby

    def _page(self, request: httpx.Request) -> httpx.Response:
        params = request.url.params
        top = int(params.get("$top", "20"))
        if top > 50:
            return graph_error(400, "BadRequest", "Invalid page size requested.")
        key = params.get("$orderby", "lastModifiedDateTime desc").split()[0]
        items = sorted(self.messages, key=lambda m: m[key], reverse=True)
        if "$filter" in params and key == "createdDateTime":
            match = re.fullmatch(r"createdDateTime lt (\S+)", params["$filter"])
            if not match:
                return graph_error(400, "BadRequest", "Invalid filter clause.")
            bound = datetime.fromisoformat(match.group(1))
            items = [m for m in items if datetime.fromisoformat(m["createdDateTime"]) < bound]
        offset = int(params.get("$skiptoken", "0"))
        size = min(top, self.page_size)
        body = {
            "@odata.context": f"https://graph.microsoft.com/v1.0/$metadata#chats('{CHAT_ID}')/messages",
            "@odata.count": len(items[offset:offset + size]),
            "value": items[offset:offset + size],
        }
        if offset + size < len(items):
            next_url = request.url.copy_merge_params({"$skiptoken": str(offset + size)})
            if self.next_link_drops_orderby:
                next_url = next_url.copy_remove_param("$orderby")
            body["@odata.nextLink"] = str(next_url)
        return httpx.Response(200, json=body)


async def chat_page(**arguments) -> dict:
    result = await call("list_chat_messages", {"chat_id": CHAT_ID, **arguments})
    assert not result.is_error, text(result)
    return json.loads(text(result))


async def test_chat_page_is_newest_first_by_creation_time(install):
    reacted_to_late = chat_message("2026-10-07T09:00:00.000Z", modified="2026-10-07T12:30:00.000Z")
    install(FakeChatMessages([
        reacted_to_late,
        chat_message("2026-10-07T10:00:00.000Z"),
        chat_message("2026-10-07T11:00:00.000Z"),
    ]))

    page = await chat_page(limit=2)

    assert [m["createdDateTime"] for m in page["messages"]] == [
        "2026-10-07T11:00:00.000Z",
        "2026-10-07T10:00:00.000Z",
    ]


async def test_next_before_walks_the_whole_chat_once(install):
    history = every_minute(7)
    history.insert(3, chat_message("2026-10-07T08:02:30.000Z", system=True))
    install(FakeChatMessages(history))

    seen, before = [], None
    for _ in range(10):  # a cursor that never runs out must not hang the test
        page = await chat_page(limit=3, **({"before": before} if before else {}))
        seen += [m["createdDateTime"] for m in page["messages"]]
        before = page["next_before"]
        if before is None:
            break

    assert seen == [
        "2026-10-07T08:06:00.000Z",
        "2026-10-07T08:05:00.000Z",
        "2026-10-07T08:04:00.000Z",
        "2026-10-07T08:03:00.000Z",
        "2026-10-07T08:02:00.000Z",
        "2026-10-07T08:01:00.000Z",
        "2026-10-07T08:00:00.000Z",
    ]


async def test_after_stops_paging_at_the_bound(install):
    graph = FakeChatMessages(every_minute(6), page_size=2)
    install(graph)

    page = await chat_page(limit=10, after="2026-10-07T08:02:00Z")

    assert [m["createdDateTime"] for m in page["messages"]] == [
        "2026-10-07T08:05:00.000Z",
        "2026-10-07T08:04:00.000Z",
        "2026-10-07T08:03:00.000Z",
    ]
    assert page["next_before"] is None
    assert len(graph.requests) == 2


async def test_limit_above_one_graph_page_follows_next_link(install):
    graph = FakeChatMessages(every_minute(120))
    install(graph)

    page = await chat_page(limit=110)

    assert len(page["messages"]) == 110
    assert page["messages"][0]["createdDateTime"] == "2026-10-07T09:59:00.000Z"
    assert page["next_before"] == "2026-10-07T08:10:00.000Z"
    assert len(graph.requests) == 3


@pytest.fixture
def local_time_not_utc():
    """Run with a local zone ahead of UTC, so reading a naive time as local time shows up.

    Windows has no time.tzset; there the machine's own zone applies.
    """
    if not hasattr(time, "tzset"):
        yield
        return
    previous = os.environ.get("TZ")
    os.environ["TZ"] = "Asia/Tashkent"
    time.tzset()
    try:
        yield
    finally:
        if previous is None:
            del os.environ["TZ"]
        else:
            os.environ["TZ"] = previous
        time.tzset()


@pytest.mark.parametrize("before", [
    "2026-10-07T12:00:00Z",
    "2026-10-07T12:00:00",
    "2026-10-07T15:00:00+03:00",
])
async def test_before_is_exclusive_and_read_as_utc_without_offset(
    install, local_time_not_utc, before,
):
    install(FakeChatMessages([
        chat_message("2026-10-07T11:59:00.000Z"),
        chat_message("2026-10-07T12:00:00.000Z"),
        chat_message("2026-10-07T12:01:00.000Z"),
    ]))

    page = await chat_page(before=before)

    assert [m["createdDateTime"] for m in page["messages"]] == ["2026-10-07T11:59:00.000Z"]


@pytest.mark.parametrize("arguments, named", [
    ({"limit": 0}, "limit"),
    ({"limit": server.MAX_LIST_LIMIT + 1}, "limit"),
    ({"before": "yesterday"}, "before"),
    ({"after": "10/07/2026"}, "after"),
])
async def test_invalid_paging_arguments_are_errors(install, arguments, named):
    graph = FakeChatMessages(every_minute(3))
    install(graph)

    result = await call("list_chat_messages", {"chat_id": CHAT_ID, **arguments})

    assert result.is_error
    assert named in text(result)
    assert graph.requests == []


# --- channel messages and thread replies beyond one page -----------------------

TEAM_ID = "fbe2bf47-16c8-47cf-b4a5-4b9b187c508b"
CHANNEL_ID = "19:4a95f7d8db4c4e7fae857bcebe0623e6@thread.tacv2"
THREAD_ID = "1759820400000"


def channel_message(created: str, *, reply_to: str | None = None) -> dict:
    """A channel chatMessage as Graph v1.0 returns it (replies carry replyToId)."""
    message = chat_message(created)
    message.update({
        "replyToId": reply_to,
        "chatId": None,
        "channelIdentity": {"teamId": TEAM_ID, "channelId": CHANNEL_ID},
    })
    return message


def channel_history(count: int, *, reply_to: str | None = None) -> list[dict]:
    """Messages one minute apart, newest first as Graph lists them."""
    return [
        channel_message(m["createdDateTime"], reply_to=reply_to)
        for m in reversed(every_minute(count))
    ]


class FakeChannel(RecordingGraph):
    """Channel endpoints as Graph v1.0 documents them: the message and reply lists take
    only $top (at most 50) and page through @odata.nextLink; one message by id."""

    def __init__(self, messages: list[dict], replies: list[dict] | None = None):
        super().__init__(self._route)
        self.messages = messages
        self.replies = replies or []

    def _route(self, request: httpx.Request) -> httpx.Response:
        base = f"/v1.0/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages"
        path = request.url.path
        if path == base:
            return self._page(request, self.messages)
        if path == f"{base}/{THREAD_ID}/replies":
            return self._page(request, self.replies)
        if path == f"{base}/{THREAD_ID}":
            return httpx.Response(200, json=channel_message("2026-10-07T07:00:00.000Z"))
        return graph_error(404, "NotFound", "Resource not found.")

    @staticmethod
    def _page(request: httpx.Request, items: list[dict]) -> httpx.Response:
        params = request.url.params
        top = int(params.get("$top", "20"))
        if top > 50:
            return graph_error(400, "BadRequest", "Invalid page size requested.")
        offset = int(params.get("$skiptoken", "0"))
        body = {"value": items[offset:offset + top]}
        if offset + top < len(items):
            body["@odata.nextLink"] = str(
                request.url.copy_merge_params({"$skiptoken": str(offset + top)})
            )
        return httpx.Response(200, json=body)


async def test_channel_messages_beyond_one_graph_page(install):
    graph = FakeChannel(channel_history(80))
    install(graph)

    result = await call("list_channel_messages", {
        "team_id": TEAM_ID, "channel_id": CHANNEL_ID, "limit": 70,
    })

    assert not result.is_error, text(result)
    assert len(json.loads(text(result))) == 70
    assert len(graph.requests) == 2


async def test_thread_replies_beyond_one_graph_page(install):
    install(FakeChannel([], channel_history(65, reply_to=THREAD_ID)))

    result = await call("list_thread_replies", {
        "team_id": TEAM_ID, "channel_id": CHANNEL_ID, "message_id": THREAD_ID, "limit": 60,
    })

    assert not result.is_error, text(result)
    thread = json.loads(text(result))
    assert thread[0]["createdDateTime"] == "2026-10-07T07:00:00.000Z"
    assert len(thread) == 1 + 60


@pytest.mark.parametrize("tool, arguments", [
    ("list_channel_messages", {"team_id": TEAM_ID, "channel_id": CHANNEL_ID}),
    (
        "list_thread_replies",
        {"team_id": TEAM_ID, "channel_id": CHANNEL_ID, "message_id": THREAD_ID},
    ),
])
@pytest.mark.parametrize("limit", [0, server.MAX_LIST_LIMIT + 1])
async def test_channel_list_limit_out_of_range_is_error(install, tool, arguments, limit):
    graph = FakeChannel(channel_history(3), channel_history(3, reply_to=THREAD_ID))
    install(graph)

    result = await call(tool, {**arguments, "limit": limit})

    assert result.is_error
    assert "limit" in text(result)
    assert graph.requests == []


# --- mentions ------------------------------------------------------------------


@pytest.mark.parametrize("mentions", [
    "@Ann Lee",
    '{"user_id": "8b081ef6-4792-4def-b2c9-c363a1bf41d5", "name": "Ann Lee"}',
    [{"user_id": "8b081ef6-4792-4def-b2c9-c363a1bf41d5"}],
    [{"name": "Ann Lee"}],
    [{"user_id": "8b081ef6-4792-4def-b2c9-c363a1bf41d5", "tag_id": "TAG1", "name": "Ann Lee"}],
])
async def test_malformed_mentions_are_rejected_before_sending(install, mentions):
    graph = RecordingGraph()
    install(graph)

    result = await call("send_chat_message", {
        "chat_id": CHAT_ID, "content": "ping @Ann Lee", "mentions": mentions,
    })

    assert result.is_error
    assert "mentions" in text(result)
    assert graph.requests == []


async def test_mentions_as_json_string_are_sent(install):
    graph = RecordingGraph(lambda request: httpx.Response(
        201, json=chat_message("2026-10-08T09:00:00.000Z"),
    ))
    install(graph)

    result = await call("send_chat_message", {
        "chat_id": CHAT_ID,
        "content": "ping @Ann Lee",
        "mentions": '[{"user_id": "8b081ef6-4792-4def-b2c9-c363a1bf41d5", "name": "Ann Lee"}]',
    })

    assert not result.is_error, text(result)
    sent = json.loads(graph.requests[0].content)
    assert sent["mentions"][0]["mentioned"]["user"]["id"] == "8b081ef6-4792-4def-b2c9-c363a1bf41d5"


# --- get_message: open a message by Teams link or ids ---------------------------

CHAT_LINK = (
    "https://teams.microsoft.com/l/message/"
    "19%3A5f4e2a10-aaaa-4bbb-8ccc-000000000001_9c8b7a60-dddd-4eee-8fff-000000000002"
    "%40unq.gbl.spaces/1759838400000?context=%7B%22contextType%22%3A%22chat%22%7D"
)
CHANNEL_REPLY_LINK = (
    "https://teams.microsoft.com/l/message/19%3A4a95f7d8db4c4e7fae857bcebe0623e6%40thread.tacv2"
    "/1759838500000?tenantId=2432b57b-0abd-43db-aa7b-16eadd115d34"
    "&groupId=fbe2bf47-16c8-47cf-b4a5-4b9b187c508b&parentMessageId=1759820400000"
    "&teamName=Data%20Platform&channelName=General&createdTime=1759838500000"
)
CHANNEL_ROOT_LINK = (
    "https://teams.microsoft.com/l/message/19%3A4a95f7d8db4c4e7fae857bcebe0623e6%40thread.tacv2"
    "/1759820400000?tenantId=2432b57b-0abd-43db-aa7b-16eadd115d34"
    "&groupId=fbe2bf47-16c8-47cf-b4a5-4b9b187c508b&parentMessageId=1759820400000"
    "&teamName=Data%20Platform&channelName=General&createdTime=1759820400000"
)
CHANNEL_BASE = f"/v1.0/teams/{TEAM_ID}/channels/{CHANNEL_ID}/messages"


@pytest.mark.parametrize("arguments, graph_path, location", [
    (
        {"link": CHAT_LINK},
        f"/v1.0/chats/{CHAT_ID}/messages/1759838400000",
        {"chat_id": CHAT_ID},
    ),
    (
        {"link": CHANNEL_REPLY_LINK},
        f"{CHANNEL_BASE}/1759820400000/replies/1759838500000",
        {"team_id": TEAM_ID, "channel_id": CHANNEL_ID, "parent_message_id": "1759820400000"},
    ),
    (
        {"link": CHANNEL_ROOT_LINK},
        f"{CHANNEL_BASE}/1759820400000",
        {"team_id": TEAM_ID, "channel_id": CHANNEL_ID},
    ),
    (
        {"chat_id": CHAT_ID, "message_id": "1759838400000"},
        f"/v1.0/chats/{CHAT_ID}/messages/1759838400000",
        {"chat_id": CHAT_ID},
    ),
])
async def test_get_message_finds_it_where_the_link_points(install, arguments, graph_path, location):
    graph = RecordingGraph(lambda request: httpx.Response(
        200, json=chat_message("2026-10-07T12:00:00.000Z"),
    ))
    install(graph)

    result = await call("get_message", arguments)

    assert not result.is_error, text(result)
    assert [r.url.path for r in graph.requests] == [graph_path]
    found = json.loads(text(result))
    assert {k: v for k, v in found.items() if k != "message"} == location
    assert found["message"]["createdDateTime"] == "2026-10-07T12:00:00.000Z"


@pytest.mark.parametrize("arguments", [
    {"link": "https://teams.microsoft.com/l/chat/19%3Aabc%40thread.v2/0"},
    {"link": "1759838400000"},
    {"message_id": "1759838400000"},
    {},
])
async def test_get_message_without_a_usable_location_is_error(install, arguments):
    graph = RecordingGraph()
    install(graph)

    result = await call("get_message", arguments)

    assert result.is_error
    assert "link" in text(result)
    assert graph.requests == []


# --- review fixes ------------------------------------------------------------


async def test_unexpected_failure_stays_a_crash(install):
    """Only anticipated failures carry their text; a bug keeps the SDK's bare message
    (and its traceback in the server log)."""
    install(RecordingGraph(lambda request: httpx.Response(
        200, text="<html>Service Unavailable</html>", headers={"content-type": "text/html"},
    )))

    result = await call("list_teams")

    assert result.is_error
    assert text(result) == "Error executing tool list_teams"


@pytest.mark.parametrize("arguments", [
    # a conversation id that climbs out of /chats into another Graph resource
    {"link": "https://teams.microsoft.com/l/message/..%2F..%2Fbeta%2Fme%2Fmessages%2FAAMkAD%3F"
             "/1759838400000"},
    # a team id that does the same through groupId
    {"link": "https://teams.microsoft.com/l/message/19%3A4a95f7d8db4c4e7fae857bcebe0623e6"
             "%40thread.tacv2/1759838500000?groupId=..%2F..%2Fme%3F"
             "&parentMessageId=1759820400000"},
    {"chat_id": "../../me/messages?", "message_id": "1759838400000"},
])
async def test_ids_cannot_steer_requests_to_other_resources(install, arguments):
    graph = RecordingGraph()
    install(graph)

    result = await call("get_message", arguments)

    assert result.is_error
    assert graph.requests == []


async def test_list_tool_ids_cannot_steer_requests(install):
    graph = FakeChatMessages(every_minute(3))
    install(graph)

    result = await call("list_chat_messages", {"chat_id": "../../me/messages?"})

    assert result.is_error
    assert graph.requests == []


async def test_link_copied_from_message_html_still_resolves(install):
    graph = RecordingGraph(lambda request: httpx.Response(
        200, json=channel_message("2026-10-07T12:00:00.000Z", reply_to=THREAD_ID),
    ))
    install(graph)

    result = await call("get_message", {"link": CHANNEL_REPLY_LINK.replace("&", "&amp;")})

    assert not result.is_error, text(result)
    assert [r.url.path for r in graph.requests] == [
        f"{CHANNEL_BASE}/1759820400000/replies/1759838500000",
    ]


async def test_channel_link_without_team_is_error(install):
    graph = RecordingGraph()
    install(graph)

    result = await call("get_message", {"link": CHANNEL_ROOT_LINK.split("?")[0]})

    assert result.is_error
    assert "team_id" in text(result)
    assert graph.requests == []


async def test_out_of_range_timestamp_is_error(install):
    graph = FakeChatMessages(every_minute(3))
    install(graph)

    result = await call("list_chat_messages", {
        "chat_id": CHAT_ID, "before": "0001-01-01T00:00:00+05:00",
    })

    assert result.is_error
    assert "before" in text(result)
    assert graph.requests == []


async def test_before_finer_than_a_millisecond_keeps_earlier_messages(install):
    install(FakeChatMessages([
        chat_message("2026-10-07T12:00:00.000Z"),
        chat_message("2026-10-07T12:00:00.001Z"),
    ]))

    page = await chat_page(before="2026-10-07T12:00:00.000500Z")

    assert [m["createdDateTime"] for m in page["messages"]] == ["2026-10-07T12:00:00.000Z"]


@pytest.mark.parametrize("blank", ["", "  "])
async def test_blank_time_bounds_mean_no_bound(install, blank):
    install(FakeChatMessages(every_minute(2)))

    page = await chat_page(before=blank, after=blank)

    assert len(page["messages"]) == 2


@pytest.mark.parametrize("mentions", ["", "null", "[]"])
async def test_empty_mentions_send_a_plain_message(install, mentions):
    graph = RecordingGraph(lambda request: httpx.Response(
        201, json=chat_message("2026-10-08T09:00:00.000Z"),
    ))
    install(graph)

    result = await call("send_chat_message", {
        "chat_id": CHAT_ID, "content": "hello", "mentions": mentions,
    })

    assert not result.is_error, text(result)
    assert "mentions" not in json.loads(graph.requests[0].content)


async def test_next_link_off_graph_is_not_followed(install):
    def respond(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json={
            "value": [chat_message("2026-10-07T12:00:00.000Z")],
            "@odata.nextLink": "https://graph.example.net/v1.0/chats/x/messages?$skiptoken=1",
        })

    graph = RecordingGraph(respond)
    install(graph)

    result = await call("list_chat_messages", {"chat_id": CHAT_ID, "limit": 5})

    assert result.is_error
    assert {r.url.host for r in graph.requests} == {"graph.microsoft.com"}


async def test_endless_empty_pages_stop_with_an_error(install):
    def respond(request: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json={
            "value": [],
            "@odata.nextLink": f"{GRAPH_BASE}/chats/{CHAT_ID}/messages?$skiptoken=again",
        })

    graph = RecordingGraph(respond)
    install(graph)

    result = await call("list_chat_messages", {"chat_id": CHAT_ID})

    assert result.is_error
    assert len(graph.requests) == MAX_PAGES


async def test_pages_out_of_creation_order_are_an_error(install):
    history = every_minute(4)
    history[0] = chat_message("2026-10-07T08:00:00.000Z", modified="2026-10-07T09:00:00.000Z")
    install(FakeChatMessages(history, page_size=2, next_link_drops_orderby=True))

    result = await call("list_chat_messages", {"chat_id": CHAT_ID, "limit": 4})

    assert result.is_error
    assert "order" in text(result)
