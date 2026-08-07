from unittest.mock import MagicMock, patch

import pytest

import outlook_cli.mcp_server as mcp_server


class TestGetClient:
    def test_returns_outlook_mock_when_env_set(self, monkeypatch):
        monkeypatch.setenv("OUTLOOK_MOCK", "1")
        from outlook_cli.mock import OutlookMock

        client = mcp_server.get_client()

        assert isinstance(client, OutlookMock)
