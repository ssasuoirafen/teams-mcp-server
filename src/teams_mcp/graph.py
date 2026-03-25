from typing import Any, Callable

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


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
                "$select": "id,topic,chatType,lastUpdatedDateTime",
            },
        )
        return data.get("value", [])

    async def list_channel_messages(self, team_id: str, channel_id: str, limit: int = 20) -> list[dict]:
        data = await self._get(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            params={"$top": limit},
        )
        return data.get("value", [])

    async def list_chat_messages(self, chat_id: str, limit: int = 20) -> list[dict]:
        data = await self._get(
            f"/chats/{chat_id}/messages",
            params={"$top": limit},
        )
        return data.get("value", [])

    async def send_channel_message(
        self, team_id: str, channel_id: str, content: str, content_type: str = "text"
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            {"body": {"content": content, "contentType": content_type}},
        )

    async def send_chat_message(self, chat_id: str, content: str, content_type: str = "text") -> dict:
        return await self._post(
            f"/chats/{chat_id}/messages",
            {"body": {"content": content, "contentType": content_type}},
        )

    async def reply_to_channel_message(
        self, team_id: str, channel_id: str, message_id: str, content: str, content_type: str = "text"
    ) -> dict:
        return await self._post(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            {"body": {"content": content, "contentType": content_type}},
        )

    async def close(self):
        await self._http.aclose()
