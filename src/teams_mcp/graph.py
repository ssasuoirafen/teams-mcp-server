from typing import Any, Callable

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"


class GraphApiError(Exception):
    def __init__(self, status_code: int, code: str, message: str):
        self.status_code = status_code
        self.code = code
        super().__init__(message)


class GraphClient:
    def __init__(self, token_provider: Callable[[], str | None]):
        self._token_provider = token_provider
        self._http = httpx.AsyncClient(
            base_url=GRAPH_BASE,
            timeout=30.0,
        )

    def _headers(self) -> dict[str, str]:
        token = self._token_provider()
        if not token:
            raise RuntimeError("Not authenticated. Call the login tool first.")
        return {"Authorization": f"Bearer {token}"}

    @staticmethod
    def _raise_for_status(resp: httpx.Response) -> None:
        if resp.is_success:
            return
        try:
            error = resp.json().get("error", {})
            code = error.get("code", resp.reason_phrase)
            message = error.get("message", resp.text)
        except Exception:
            code = resp.reason_phrase or str(resp.status_code)
            message = resp.text
        raise GraphApiError(resp.status_code, code, f"Graph API error {resp.status_code} ({code}): {message}")

    async def _get(self, path: str, params: dict | None = None) -> dict[str, Any]:
        resp = await self._http.get(path, headers=self._headers(), params=params)
        self._raise_for_status(resp)
        return resp.json()

    async def _post(self, path: str, json_body: dict) -> dict[str, Any]:
        resp = await self._http.post(path, headers=self._headers(), json=json_body)
        self._raise_for_status(resp)
        return resp.json()

    async def _post_no_content(self, path: str, json_body: dict | None = None) -> None:
        resp = await self._http.post(path, headers=self._headers(), json=json_body)
        self._raise_for_status(resp)

    async def _patch(self, path: str, json_body: dict) -> None:
        resp = await self._http.patch(path, headers=self._headers(), json=json_body)
        self._raise_for_status(resp)

    async def _delete(self, path: str) -> None:
        resp = await self._http.delete(path, headers=self._headers())
        self._raise_for_status(resp)

    async def list_teams(self) -> list[dict]:
        data = await self._get("/me/joinedTeams", params={"$select": "id,displayName,description"})
        return data.get("value", [])

    async def list_channels(self, team_id: str) -> list[dict]:
        data = await self._get(
            f"/teams/{team_id}/channels",
            params={"$select": "id,displayName,description,membershipType"},
        )
        return data.get("value", [])

    async def list_chats(self, limit: int = 20) -> list[dict]:
        data = await self._get(
            "/me/chats",
            params={
                "$top": limit,
                "$expand": "members",
            },
        )
        return data.get("value", [])

    async def list_channel_messages(self, team_id: str, channel_id: str, limit: int = 20) -> list[dict]:
        data = await self._get(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            params={"$top": limit},
        )
        return data.get("value", [])

    async def list_thread_replies(
        self, team_id: str, channel_id: str, message_id: str, limit: int = 20
    ) -> list[dict]:
        data = await self._get(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            params={"$top": limit},
        )
        return data.get("value", [])

    async def get_channel_message(
        self, team_id: str, channel_id: str, message_id: str
    ) -> dict:
        return await self._get(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}",
        )

    async def list_chat_messages(self, chat_id: str, limit: int = 20) -> list[dict]:
        data = await self._get(
            f"/chats/{chat_id}/messages",
            params={"$top": limit},
        )
        return data.get("value", [])

    @staticmethod
    def _to_html(text: str) -> str:
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")

    @staticmethod
    def _build_message_body(text: str, mentions: list[dict] | None = None) -> dict:
        html = GraphClient._to_html(text)
        payload: dict[str, Any] = {
            "body": {"content": html, "contentType": "html"},
        }
        if mentions:
            mention_objects = []
            idx = 0
            for m in sorted(mentions, key=lambda x: len(x["name"]), reverse=True):
                name = m["name"]
                escaped = (
                    name.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                )
                at_tag = f'<at id="{idx}">{escaped}</at>'
                new_html = html.replace(f"@{escaped}", at_tag)
                if new_html != html:
                    html = new_html
                    mention_objects.append({
                        "id": idx,
                        "mentionText": name,
                        "mentioned": {
                            "user": {
                                "id": m["user_id"],
                                "displayName": name,
                                "userIdentityType": "aadUser",
                            }
                        },
                    })
                    idx += 1
            payload["body"]["content"] = html
            if mention_objects:
                payload["mentions"] = mention_objects
        return payload

    async def send_channel_message(
        self, team_id: str, channel_id: str, content: str, mentions: list[dict] | None = None,
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            self._build_message_body(content, mentions),
        )

    async def send_chat_message(
        self, chat_id: str, content: str, mentions: list[dict] | None = None,
    ) -> dict:
        return await self._post(
            f"/chats/{chat_id}/messages",
            self._build_message_body(content, mentions),
        )

    async def reply_to_channel_message(
        self, team_id: str, channel_id: str, message_id: str, content: str,
        mentions: list[dict] | None = None,
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            self._build_message_body(content, mentions),
        )

    async def get_me(self) -> dict:
        return await self._get("/me", params={"$select": "id"})

    async def create_chat(self, my_id: str, user_email: str) -> dict:
        return await self._post(
            "/chats",
            {
                "chatType": "oneOnOne",
                "members": [
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
                    },
                    {
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "roles": ["owner"],
                        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_email}')",
                    },
                ],
            },
        )

    async def list_team_members(self, team_id: str) -> list[dict]:
        data = await self._get(f"/teams/{team_id}/members")
        return data.get("value", [])

    async def list_channel_members(self, team_id: str, channel_id: str) -> list[dict]:
        data = await self._get(f"/teams/{team_id}/channels/{channel_id}/members")
        return data.get("value", [])

    async def list_chat_members(self, chat_id: str) -> list[dict]:
        data = await self._get(f"/chats/{chat_id}/members")
        return data.get("value", [])

    async def soft_delete_channel_message(
        self, team_id: str, channel_id: str, message_id: str,
    ) -> None:
        await self._post_no_content(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/softDelete",
        )

    async def soft_delete_chat_message(self, chat_id: str, message_id: str) -> None:
        await self._post_no_content(
            f"/chats/{chat_id}/messages/{message_id}/softDelete",
        )

    async def update_channel_message(
        self, team_id: str, channel_id: str, message_id: str, content: str,
    ) -> None:
        await self._patch(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}",
            {"body": {"content": self._to_html(content), "contentType": "html"}},
        )

    async def update_chat_message(self, chat_id: str, message_id: str, content: str) -> None:
        await self._patch(
            f"/chats/{chat_id}/messages/{message_id}",
            {"body": {"content": self._to_html(content), "contentType": "html"}},
        )

    async def set_reaction_channel(
        self, team_id: str, channel_id: str, message_id: str, reaction: str,
    ) -> None:
        await self._post_no_content(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/setReaction",
            {"reactionType": reaction},
        )

    async def set_reaction_chat(self, chat_id: str, message_id: str, reaction: str) -> None:
        await self._post_no_content(
            f"/chats/{chat_id}/messages/{message_id}/setReaction",
            {"reactionType": reaction},
        )

    async def unset_reaction_channel(
        self, team_id: str, channel_id: str, message_id: str, reaction: str,
    ) -> None:
        await self._post_no_content(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/unsetReaction",
            {"reactionType": reaction},
        )

    async def unset_reaction_chat(self, chat_id: str, message_id: str, reaction: str) -> None:
        await self._post_no_content(
            f"/chats/{chat_id}/messages/{message_id}/unsetReaction",
            {"reactionType": reaction},
        )

    async def create_group_chat(
        self, my_id: str, member_emails: list[str], topic: str | None = None,
    ) -> dict:
        members = [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{my_id}')",
            }
        ]
        for email in member_emails:
            members.append({
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{email}')",
            })
        body: dict[str, Any] = {"chatType": "group", "members": members}
        if topic:
            body["topic"] = topic
        return await self._post("/chats", body)

    async def pin_message(self, chat_id: str, message_id: str) -> dict:
        return await self._post(
            f"/chats/{chat_id}/pinnedMessages",
            {
                "message@odata.bind": (
                    f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages/{message_id}"
                ),
            },
        )

    async def unpin_message(self, chat_id: str, pinned_message_id: str) -> None:
        await self._delete(f"/chats/{chat_id}/pinnedMessages/{pinned_message_id}")

    async def list_pinned_messages(self, chat_id: str) -> list[dict]:
        data = await self._get(
            f"/chats/{chat_id}/pinnedMessages",
            params={"$expand": "message"},
        )
        return data.get("value", [])

    async def mark_chat_read(self, chat_id: str, user_id: str) -> None:
        await self._post_no_content(
            f"/chats/{chat_id}/markChatReadForUser",
            {"user": {"id": user_id, "@odata.type": "#microsoft.graph.teamworkUserIdentity"}},
        )

    async def mark_chat_unread(
        self, chat_id: str, user_id: str, last_message_read_date_time: str,
    ) -> None:
        await self._post_no_content(
            f"/chats/{chat_id}/markChatUnreadForUser",
            {
                "user": {"id": user_id, "@odata.type": "#microsoft.graph.teamworkUserIdentity"},
                "lastMessageReadDateTime": last_message_read_date_time,
            },
        )

    async def get_user_presence(self, user_id: str) -> dict:
        return await self._get(f"/users/{user_id}/presence")

    async def search_messages(self, query: str, size: int = 25) -> list[dict]:
        resp = await self._http.post(
            f"{GRAPH_BETA}/search/query",
            headers=self._headers(),
            json={
                "requests": [
                    {
                        "entityTypes": ["chatMessage"],
                        "query": {"queryString": query},
                        "from": 0,
                        "size": size,
                    }
                ],
            },
        )
        self._raise_for_status(resp)
        data = resp.json()
        values = data.get("value") or [{}]
        containers = values[0].get("hitsContainers", [])
        if not containers:
            return []
        return containers[0].get("hits", [])

    async def search_users(self, query: str, limit: int = 10) -> list[dict]:
        safe = query.replace("'", "''")
        data = await self._get(
            "/users",
            params={
                "$filter": f"startsWith(displayName,'{safe}') or startsWith(mail,'{safe}')",
                "$select": "id,displayName,mail,userPrincipalName,jobTitle",
                "$top": limit,
            },
        )
        return data.get("value", [])

    async def close(self):
        await self._http.aclose()
