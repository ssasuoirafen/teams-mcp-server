from typing import Any, Callable

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"


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

    async def _get(self, path: str, params: dict | None = None) -> dict[str, Any]:
        resp = await self._http.get(path, headers=self._headers(), params=params)
        resp.raise_for_status()
        return resp.json()

    async def _post(self, path: str, json_body: dict) -> dict[str, Any]:
        resp = await self._http.post(path, headers=self._headers(), json=json_body)
        resp.raise_for_status()
        return resp.json()

    async def _post_no_content(self, path: str, json_body: dict | None = None) -> None:
        resp = await self._http.post(path, headers=self._headers(), json=json_body)
        resp.raise_for_status()

    async def _patch(self, path: str, json_body: dict) -> None:
        resp = await self._http.patch(path, headers=self._headers(), json=json_body)
        resp.raise_for_status()

    async def _delete(self, path: str) -> None:
        resp = await self._http.delete(path, headers=self._headers())
        resp.raise_for_status()

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

    async def send_channel_message(self, team_id: str, channel_id: str, content: str) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            {"body": {"content": self._to_html(content), "contentType": "html"}},
        )

    async def send_chat_message(self, chat_id: str, content: str) -> dict:
        return await self._post(
            f"/chats/{chat_id}/messages",
            {"body": {"content": self._to_html(content), "contentType": "html"}},
        )

    async def reply_to_channel_message(
        self, team_id: str, channel_id: str, message_id: str, content: str
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            {"body": {"content": self._to_html(content), "contentType": "html"}},
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

    async def close(self):
        await self._http.aclose()
