import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import msal
import requests

DEFAULT_SCOPES = ["https://graph.microsoft.com/.default"]


class AuthError(Exception):
    """Sign-in is missing or failed; the message says what the user should do."""


@contextmanager
def _network_errors_as_auth_errors(action: str) -> Iterator[None]:
    """msal reaches Entra through requests and lets its exceptions through.

    A network failure is a sign-in failure the user can act on, not a crash. The message
    must not say "Not authenticated": that sends the agent to log in again.
    """
    try:
        yield
    except requests.RequestException as exc:
        raise AuthError(f"{action} failed: {type(exc).__name__}: {exc}") from exc


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
        with _network_errors_as_auth_errors("Refreshing the sign-in"):
            result = self._app.acquire_token_silent(
                scopes=self.scopes, account=accounts[0]
            )
        self._save_cache()
        if result and "access_token" in result:
            return result["access_token"]
        return None

    def login(self) -> dict:
        with _network_errors_as_auth_errors("Starting the device code sign-in"):
            flow = self._app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            raise AuthError(
                f"Device flow failed: {flow.get('error_description', 'unknown error')}"
            )
        return flow

    def complete_login(self, flow: dict) -> dict:
        with _network_errors_as_auth_errors("Completing the device code sign-in"):
            result = self._app.acquire_token_by_device_flow(flow)
        self._save_cache()
        if "access_token" in result:
            return {
                "status": "ok",
                "account": result.get("id_token_claims", {}).get(
                    "preferred_username", "unknown"
                ),
            }
        raise AuthError(
            result.get("error_description", "Authentication failed")
        )

    def is_authenticated(self) -> bool:
        return self.get_token() is not None
