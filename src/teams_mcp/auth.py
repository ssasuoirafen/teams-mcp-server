import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import msal
import requests

DEFAULT_SCOPES = ["https://graph.microsoft.com/.default"]


class AuthError(Exception):
    """Sign-in is missing or failed; the message says what the user should do."""


# The server instructions key off "Not authenticated", so keep that prefix.
NOT_AUTHENTICATED = (
    "Not authenticated. Ask the user to run `teams-mcp login` in a terminal (the command "
    "that starts this server, with `login` appended), then retry."
)


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
        self._cache_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
        self._cache_path = self._cache_dir / "token_cache.json"
        self._cache = msal.SerializableTokenCache()
        self._cache_text: str | None = None  # the file content this process last read or wrote
        self._load_cache()
        self._msal_app: msal.PublicClientApplication | None = None

    @property
    def _app(self) -> msal.PublicClientApplication:
        """The MSAL app, created on first use.

        msal contacts Entra while constructing it (tenant discovery), so a wrong tenant or a
        missing network fails here, with a reason, on a tool call or `teams-mcp login`
        rather than as a crash of the server at start-up.
        """
        if self._msal_app is None:
            with _network_errors_as_auth_errors("Reaching Entra"):
                try:
                    self._msal_app = msal.PublicClientApplication(
                        client_id=self.client_id,
                        authority=f"https://login.microsoftonline.com/{self.tenant_id}",
                        token_cache=self._cache,
                    )
                except ValueError as exc:
                    # msal's own message is generic; Entra's reason (e.g. AADSTS90002 Tenant
                    # not found) is on the discovery error it was handling
                    reason = exc.__context__ or exc
                    raise AuthError(
                        f"Entra rejected TEAMS_MCP_TENANT_ID {self.tenant_id!r}, check the "
                        f"tenant ID: {reason}"
                    ) from exc
        return self._msal_app

    def _load_cache(self) -> bool:
        """Load the cache file if it changed since this process last read or wrote it.

        Returns True when it did: another process wrote it, e.g. `teams-mcp login`.
        """
        if not self._cache_path.exists():
            return False
        text = self._cache_path.read_text(encoding="utf-8")
        if text == self._cache_text:
            return False
        self._cache.deserialize(text)
        self._cache_text = text
        return True

    def _save_cache(self):
        if not self._cache.has_state_changed:
            return
        text = self._cache.serialize()
        # The server and `teams-mcp login` share this file, so write a temp file and swap
        # it in: a reader never sees it half written. It holds refresh tokens, so only
        # the owner may read it.
        tmp_path = self._cache_path.with_name(self._cache_path.name + ".tmp")
        fd = os.open(tmp_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as tmp:
            tmp.write(text)
        os.chmod(tmp_path, 0o600)  # in case a leftover temp file had a wider mode
        os.replace(tmp_path, self._cache_path)
        self._cache_text = text

    def get_token(self) -> str | None:
        token = self._silent_token()
        # The server keeps the cache in memory from startup. If the user signed in from a
        # terminal since then (`teams-mcp login`), the token is only in the file.
        if token is None and self._load_cache():
            token = self._silent_token()
        return token

    def _silent_token(self) -> str | None:
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

    def username(self) -> str | None:
        """The signed-in account's user name, if the cache holds one."""
        accounts = self._app.get_accounts()
        return accounts[0].get("username") if accounts else None

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
