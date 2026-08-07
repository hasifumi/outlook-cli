from unittest.mock import MagicMock, patch

import pytest

import outlook_cli.mcp_server as mcp_server


class TestGetClient:
    def test_returns_outlook_mock_when_env_set(self, monkeypatch):
        monkeypatch.setenv("OUTLOOK_MOCK", "1")
        from outlook_cli.mock import OutlookMock

        client = mcp_server.get_client()

        assert isinstance(client, OutlookMock)


class TestListTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_list_mails_with_folder_and_limit(self, mock_get_client):
        client = MagicMock()
        client.list_mails.return_value = [{"id": "1", "subject": "test"}]
        mock_get_client.return_value = client

        result = mcp_server.list_mails_tool(folder="inbox", limit=5)

        client.list_mails.assert_called_once_with(folder="inbox", limit=5)
        assert result == [{"id": "1", "subject": "test"}]

    @patch("outlook_cli.mcp_server.get_client")
    def test_uses_default_folder_and_limit(self, mock_get_client):
        client = MagicMock()
        client.list_mails.return_value = []
        mock_get_client.return_value = client

        mcp_server.list_mails_tool()

        client.list_mails.assert_called_once_with(folder="inbox", limit=20)


class TestSearchTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_search_with_all_params(self, mock_get_client):
        client = MagicMock()
        client.search.return_value = [{"id": "1", "subject": "hit"}]
        mock_get_client.return_value = client

        result = mcp_server.search(keyword="foo", days=3, sender="a@b.com")

        client.search.assert_called_once_with(keyword="foo", days=3, sender="a@b.com")
        assert result == [{"id": "1", "subject": "hit"}]

    @patch("outlook_cli.mcp_server.get_client")
    def test_uses_default_days_and_sender(self, mock_get_client):
        client = MagicMock()
        client.search.return_value = []
        mock_get_client.return_value = client

        mcp_server.search(keyword="foo")

        client.search.assert_called_once_with(keyword="foo", days=7, sender=None)
