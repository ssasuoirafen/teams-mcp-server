from unittest.mock import MagicMock, patch

import pytest

from teams_mcp.auth import DEFAULT_SCOPES, AuthManager


@pytest.fixture
def auth_manager(tmp_path):
    with patch("msal.PublicClientApplication"), \
         patch("msal.SerializableTokenCache") as mock_cache_cls:
        mock_cache = MagicMock()
        mock_cache.has_state_changed = True
        mock_cache.serialize.return_value = "{}"
        mock_cache_cls.return_value = mock_cache
        manager = AuthManager(
            tenant_id="test-tenant",
            client_id="test-client",
            cache_dir=str(tmp_path),
        )
        return manager


def test_auth_manager_init(tmp_path):
    with patch("msal.PublicClientApplication") as mock_app_cls, \
         patch("msal.SerializableTokenCache"):
        manager = AuthManager(
            tenant_id="test-tenant",
            client_id="test-client",
            cache_dir=str(tmp_path),
        )
        assert manager.tenant_id == "test-tenant"
        assert manager.client_id == "test-client"
        assert manager.scopes == DEFAULT_SCOPES
        mock_app_cls.assert_called_once_with(
            client_id="test-client",
            authority="https://login.microsoftonline.com/test-tenant",
            token_cache=manager._cache,
        )


def test_not_authenticated_initially(auth_manager):
    auth_manager._app.get_accounts.return_value = []
    assert auth_manager.is_authenticated() is False


def test_get_token_with_cached_account(auth_manager):
    mock_account = {"username": "user@example.com"}
    auth_manager._app.get_accounts.return_value = [mock_account]
    auth_manager._app.acquire_token_silent.return_value = {
        "access_token": "fake-token-abc"
    }
    token = auth_manager.get_token()
    assert token == "fake-token-abc"
    auth_manager._app.acquire_token_silent.assert_called_once_with(
        scopes=auth_manager.scopes, account=mock_account
    )


def test_login_returns_flow(auth_manager):
    flow = {
        "user_code": "ABCD1234",
        "verification_uri": "https://microsoft.com/devicelogin",
        "message": "Visit https://microsoft.com/devicelogin and enter ABCD1234",
    }
    auth_manager._app.initiate_device_flow.return_value = flow
    result = auth_manager.login()
    assert result == flow
    auth_manager._app.initiate_device_flow.assert_called_once_with(
        scopes=auth_manager.scopes
    )


def test_login_failure(auth_manager):
    auth_manager._app.initiate_device_flow.return_value = {
        "error": "authorization_pending",
        "error_description": "Device flow initiation failed",
    }
    with pytest.raises(RuntimeError, match="Device flow failed: Device flow initiation failed"):
        auth_manager.login()


def test_complete_login_success(auth_manager):
    flow = {"device_code": "some-device-code"}
    auth_manager._app.acquire_token_by_device_flow.return_value = {
        "access_token": "real-token",
        "id_token_claims": {"preferred_username": "user@example.com"},
    }
    result = auth_manager.complete_login(flow)
    assert result == {"status": "ok", "account": "user@example.com"}
    auth_manager._app.acquire_token_by_device_flow.assert_called_once_with(flow)


def test_complete_login_failure(auth_manager):
    flow = {"device_code": "some-device-code"}
    auth_manager._app.acquire_token_by_device_flow.return_value = {
        "error": "authorization_declined",
        "error_description": "User declined the request",
    }
    with pytest.raises(RuntimeError, match="User declined the request"):
        auth_manager.complete_login(flow)
