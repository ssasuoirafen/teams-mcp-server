import json

import httpx
import pytest

from teams_mcp.graph import GRAPH_BASE, GraphClient


def make_client(token: str | None = "test-token", transport: httpx.MockTransport | None = None) -> GraphClient:
    client = GraphClient(token_provider=lambda: token)
    if transport is not None:
        client._http = httpx.AsyncClient(base_url=GRAPH_BASE, timeout=30.0, transport=transport)
    return client


def mock_transport(responses: dict[tuple[str, str], tuple[int, dict]]) -> httpx.MockTransport:
    """
    responses: mapping of (method, path_prefix) -> (status_code, body_dict)
    """
    def handler(request: httpx.Request) -> httpx.Response:
        for (method, path_prefix), (status, body) in responses.items():
            if request.method == method and request.url.path.startswith(path_prefix):
                return httpx.Response(status, json=body, request=request)
        return httpx.Response(404, json={"error": "not found"}, request=request)

    return httpx.MockTransport(handler)


@pytest.mark.asyncio
async def test_list_teams():
    transport = mock_transport({
        ("GET", "/v1.0/me/joinedTeams"): (200, {"value": [{"id": "t1", "displayName": "Team Alpha"}]}),
    })
    client = make_client(transport=transport)
    result = await client.list_teams()
    assert result == [{"id": "t1", "displayName": "Team Alpha"}]
    await client.close()


@pytest.mark.asyncio
async def test_list_channels():
    team_id = "team-123"
    transport = mock_transport({
        ("GET", f"/v1.0/teams/{team_id}/channels"): (
            200,
            {"value": [{"id": "c1", "displayName": "General", "membershipType": "standard"}]},
        ),
    })
    client = make_client(transport=transport)
    result = await client.list_channels(team_id)
    assert len(result) == 1
    assert result[0]["id"] == "c1"
    await client.close()


@pytest.mark.asyncio
async def test_list_chats():
    transport = mock_transport({
        ("GET", "/v1.0/me/chats"): (
            200,
            {"value": [{"id": "chat-1", "chatType": "oneOnOne", "topic": None}]},
        ),
    })
    client = make_client(transport=transport)
    result = await client.list_chats(limit=10)
    assert result == [{"id": "chat-1", "chatType": "oneOnOne", "topic": None}]
    await client.close()


@pytest.mark.asyncio
async def test_send_channel_message():
    team_id = "team-abc"
    channel_id = "chan-xyz"
    captured: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(201, json={"id": "msg-1", "body": {"content": "Hello!"}}, request=request)

    client = make_client(transport=httpx.MockTransport(handler))
    result = await client.send_channel_message(team_id, channel_id, "Hello!")
    assert result["id"] == "msg-1"

    assert len(captured) == 1
    sent = json.loads(captured[0].content)
    assert sent == {"body": {"content": "Hello!", "contentType": "html"}}
    await client.close()


@pytest.mark.asyncio
async def test_no_token_raises():
    client = make_client(token=None)
    with pytest.raises(RuntimeError, match="Not authenticated"):
        await client.list_teams()
    await client.close()


@pytest.mark.asyncio
async def test_http_error_raises():
    transport = mock_transport({
        ("GET", "/v1.0/me/joinedTeams"): (401, {"error": {"code": "Unauthorized"}}),
    })
    client = make_client(transport=transport)
    with pytest.raises(httpx.HTTPStatusError):
        await client.list_teams()
    await client.close()


@pytest.mark.asyncio
async def test_post_no_content():
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(204, request=request)

    client = make_client(transport=httpx.MockTransport(handler))
    await client._post_no_content("/test/action", {"key": "value"})
    await client.close()


@pytest.mark.asyncio
async def test_patch():
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(204, request=request)

    client = make_client(transport=httpx.MockTransport(handler))
    await client._patch("/test/resource", {"body": {"content": "updated"}})
    await client.close()


@pytest.mark.asyncio
async def test_delete():
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(204, request=request)

    client = make_client(transport=httpx.MockTransport(handler))
    await client._delete("/test/resource/123")
    await client.close()


MEMBER_RESPONSE = {
    "value": [
        {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "id": "mem-1",
            "displayName": "Alice Smith",
            "email": "alice@example.com",
            "roles": ["owner"],
        }
    ]
}


@pytest.mark.asyncio
async def test_list_team_members():
    transport = mock_transport({
        ("GET", "/v1.0/teams/team-1/members"): (200, MEMBER_RESPONSE),
    })
    client = make_client(transport=transport)
    result = await client.list_team_members("team-1")
    assert len(result) == 1
    assert result[0]["displayName"] == "Alice Smith"
    await client.close()


@pytest.mark.asyncio
async def test_list_channel_members():
    transport = mock_transport({
        ("GET", "/v1.0/teams/team-1/channels/chan-1/members"): (200, MEMBER_RESPONSE),
    })
    client = make_client(transport=transport)
    result = await client.list_channel_members("team-1", "chan-1")
    assert len(result) == 1
    await client.close()


@pytest.mark.asyncio
async def test_list_chat_members():
    transport = mock_transport({
        ("GET", "/v1.0/chats/chat-1/members"): (200, MEMBER_RESPONSE),
    })
    client = make_client(transport=transport)
    result = await client.list_chat_members("chat-1")
    assert len(result) == 1
    await client.close()


@pytest.mark.asyncio
async def test_soft_delete_channel_message():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.soft_delete_channel_message("team-1", "chan-1", "msg-1")
    assert "/softDelete" in captured[0].url.path
    await client.close()


@pytest.mark.asyncio
async def test_soft_delete_chat_message():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.soft_delete_chat_message("chat-1", "msg-1")
    assert "/chats/chat-1/messages/msg-1/softDelete" in captured[0].url.path
    await client.close()


@pytest.mark.asyncio
async def test_update_channel_message():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.update_channel_message("team-1", "chan-1", "msg-1", "updated text")
    assert captured[0].method == "PATCH"
    body = json.loads(captured[0].content)
    assert body["body"]["contentType"] == "html"
    await client.close()


@pytest.mark.asyncio
async def test_update_chat_message():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.update_chat_message("chat-1", "msg-1", "new content")
    assert captured[0].method == "PATCH"
    assert "/chats/chat-1/messages/msg-1" in captured[0].url.path
    await client.close()


@pytest.mark.asyncio
async def test_set_reaction_channel():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.set_reaction_channel("team-1", "chan-1", "msg-1", "like")
    assert len(captured) == 1
    assert "/messages/msg-1/setReaction" in captured[0].url.path
    body = json.loads(captured[0].content)
    assert body == {"reactionType": "like"}
    await client.close()


@pytest.mark.asyncio
async def test_set_reaction_chat():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.set_reaction_chat("chat-1", "msg-1", "heart")
    assert "/chats/chat-1/messages/msg-1/setReaction" in captured[0].url.path
    await client.close()


@pytest.mark.asyncio
async def test_unset_reaction_channel():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.unset_reaction_channel("team-1", "chan-1", "msg-1", "like")
    assert "/unsetReaction" in captured[0].url.path
    await client.close()


@pytest.mark.asyncio
async def test_create_group_chat():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(201, json={"id": "chat-new", "chatType": "group"}, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    result = await client.create_group_chat(
        my_id="me-id",
        member_emails=["a@example.com", "b@example.com"],
        topic="Project X",
    )
    assert result["chatType"] == "group"
    body = json.loads(captured[0].content)
    assert body["chatType"] == "group"
    assert body["topic"] == "Project X"
    assert len(body["members"]) == 3
    await client.close()


@pytest.mark.asyncio
async def test_pin_message():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(201, json={"id": "pin-1"}, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    result = await client.pin_message("chat-1", "msg-1")
    body = json.loads(captured[0].content)
    assert "message@odata.bind" in body
    await client.close()


@pytest.mark.asyncio
async def test_unpin_message():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.unpin_message("chat-1", "pin-1")
    assert captured[0].method == "DELETE"
    assert "/pinnedMessages/pin-1" in captured[0].url.path
    await client.close()


@pytest.mark.asyncio
async def test_list_pinned_messages():
    transport = mock_transport({
        ("GET", "/v1.0/chats/chat-1/pinnedMessages"): (
            200,
            {"value": [{"id": "pin-1", "message": {"id": "msg-1", "body": {"content": "Important"}}}]},
        ),
    })
    client = make_client(transport=transport)
    result = await client.list_pinned_messages("chat-1")
    assert len(result) == 1
    await client.close()


@pytest.mark.asyncio
async def test_mark_chat_read():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.mark_chat_read("chat-1", "user-1")
    assert "/markChatReadForUser" in captured[0].url.path
    body = json.loads(captured[0].content)
    assert body["user"]["id"] == "user-1"
    await client.close()


@pytest.mark.asyncio
async def test_mark_chat_unread():
    captured: list[httpx.Request] = []
    def handler(request: httpx.Request) -> httpx.Response:
        captured.append(request)
        return httpx.Response(204, request=request)
    client = make_client(transport=httpx.MockTransport(handler))
    await client.mark_chat_unread("chat-1", "user-1", "2026-03-26T10:00:00Z")
    assert "/markChatUnreadForUser" in captured[0].url.path
    body = json.loads(captured[0].content)
    assert body["lastMessageReadDateTime"] == "2026-03-26T10:00:00Z"
    await client.close()


@pytest.mark.asyncio
async def test_get_user_presence():
    transport = mock_transport({
        ("GET", "/v1.0/users/user-1/presence"): (
            200,
            {"availability": "Available", "activity": "Available"},
        ),
    })
    client = make_client(transport=transport)
    result = await client.get_user_presence("user-1")
    assert result["availability"] == "Available"
    await client.close()


@pytest.mark.asyncio
async def test_search_users():
    transport = mock_transport({
        ("GET", "/v1.0/users"): (
            200,
            {
                "value": [
                    {"id": "u1", "displayName": "Alice", "mail": "alice@example.com", "userPrincipalName": "alice@example.com"}
                ]
            },
        ),
    })
    client = make_client(transport=transport)
    result = await client.search_users("Alice")
    assert len(result) == 1
    assert result[0]["displayName"] == "Alice"
    await client.close()
