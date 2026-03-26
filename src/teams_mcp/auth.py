import os
from pathlib import Path

import msal


DEFAULT_SCOPES = ["https://graph.microsoft.com/.default"]


class AuthManager:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        scopes: list[str] | None = None,
        cache_dir: str | None = None,
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.scopes = scopes or DEFAULT_SCOPES
        self._cache_dir = Path(cache_dir or os.path.expanduser("~/.teams-mcp"))
        self._cache_dir.mkdir(parents=True, exist_ok=True)
        self._cache_path = self._cache_dir / "token_cache.json"
        self._cache = msal.SerializableTokenCache()
        self._load_cache()
        self._app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            token_cache=self._cache,
        )

    def _load_cache(self):
        if self._cache_path.exists():
            self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))

    def _save_cache(self):
        if self._cache.has_state_changed:
            self._cache_path.write_text(
                self._cache.serialize(), encoding="utf-8"
            )

    def get_token(self) -> str | None:
        accounts = self._app.get_accounts()
        if not accounts:
            return None
        result = self._app.acquire_token_silent(
            scopes=self.scopes, account=accounts[0]
        )
        self._save_cache()
        if result and "access_token" in result:
            return result["access_token"]
        return None

    def login(self) -> dict:
        flow = self._app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Device flow failed: {flow.get('error_description', 'unknown error')}"
            )
        return flow

    def complete_login(self, flow: dict) -> dict:
        result = self._app.acquire_token_by_device_flow(flow)
        self._save_cache()
        if "access_token" in result:
            return {
                "status": "ok",
                "account": result.get("id_token_claims", {}).get(
                    "preferred_username", "unknown"
                ),
            }
        raise RuntimeError(
            result.get("error_description", "Authentication failed")
        )

    def is_authenticated(self) -> bool:
        return self.get_token() is not None
