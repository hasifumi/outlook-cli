from datetime import datetime
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


class TestReadTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_read_with_mail_id(self, mock_get_client):
        client = MagicMock()
        client.read.return_value = {"id": "1", "subject": "test", "body": "hello"}
        mock_get_client.return_value = client

        result = mcp_server.read(mail_id="1")

        client.read.assert_called_once_with("1")
        assert result == {"id": "1", "subject": "test", "body": "hello"}


class TestSendTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_send_and_returns_confirmation(self, mock_get_client):
        client = MagicMock()
        mock_get_client.return_value = client

        result = mcp_server.send(to="a@b.com", subject="s", body="b")

        client.send.assert_called_once_with(to="a@b.com", subject="s", body="b")
        assert result == {"status": "sent"}


class TestReplyTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_reply_and_returns_confirmation(self, mock_get_client):
        client = MagicMock()
        mock_get_client.return_value = client

        result = mcp_server.reply(mail_id="1", body="b")

        client.reply.assert_called_once_with(mail_id="1", body="b")
        assert result == {"status": "replied"}


class TestUnreadCountTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_returns_dict_with_folder_keys_and_total(self, mock_get_client):
        client = MagicMock()
        client.unread_count.return_value = {"inbox": 3, "total": 3}
        mock_get_client.return_value = client

        result = mcp_server.unread_count(folder=None)

        client.unread_count.assert_called_once_with(folder=None)
        assert result == {"inbox": 3, "total": 3}


class TestUnreadSummaryTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_with_limit_and_folder(self, mock_get_client):
        client = MagicMock()
        client.unread_summary.return_value = [{"subject": "s", "preview": "p"}]
        mock_get_client.return_value = client

        result = mcp_server.unread_summary(limit=5, folder="inbox")

        client.unread_summary.assert_called_once_with(limit=5, folder="inbox")
        assert result == [{"subject": "s", "preview": "p"}]

    @patch("outlook_cli.mcp_server.get_client")
    def test_uses_defaults(self, mock_get_client):
        client = MagicMock()
        client.unread_summary.return_value = []
        mock_get_client.return_value = client

        mcp_server.unread_summary()

        client.unread_summary.assert_called_once_with(limit=10, folder="inbox")


class TestSentTodayTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_client_with_date(self, mock_get_client):
        client = MagicMock()
        client.sent_today.return_value = [{"subject": "s", "to": "a@b.com"}]
        mock_get_client.return_value = client

        result = mcp_server.sent_today(date="2026-08-07")

        client.sent_today.assert_called_once_with(date="2026-08-07")
        assert result == [{"subject": "s", "to": "a@b.com"}]

    @patch("outlook_cli.mcp_server.get_client")
    def test_uses_default_date(self, mock_get_client):
        client = MagicMock()
        client.sent_today.return_value = []
        mock_get_client.return_value = client

        mcp_server.sent_today()

        client.sent_today.assert_called_once_with(date=None)


class TestFlaggedTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_delegates_to_flagged_or_due(self, mock_get_client):
        client = MagicMock()
        client.flagged_or_due.return_value = [{"id": "1", "subject": "important"}]
        mock_get_client.return_value = client

        result = mcp_server.flagged(folder="inbox", days=3)

        client.flagged_or_due.assert_called_once_with(days=3, folder="inbox")
        assert result == [{"id": "1", "subject": "important"}]

    @patch("outlook_cli.mcp_server.get_client")
    def test_uses_defaults(self, mock_get_client):
        client = MagicMock()
        client.flagged_or_due.return_value = []
        mock_get_client.return_value = client

        mcp_server.flagged()

        client.flagged_or_due.assert_called_once_with(days=7, folder="inbox")


class TestCalTodayTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_get_calendar_with_single_day_range(self, mock_get_client):
        client = MagicMock()
        client.get_calendar.return_value = [{"subject": "meeting"}]
        mock_get_client.return_value = client

        result = mcp_server.cal_today(date="2026-08-07")

        client.get_calendar.assert_called_once_with(
            datetime(2026, 8, 7), datetime(2026, 8, 8)
        )
        assert result == [{"subject": "meeting"}]


class TestCalWeekTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_calls_get_calendar_with_monday_to_friday_range(self, mock_get_client):
        client = MagicMock()
        client.get_calendar.return_value = [{"subject": "weekly"}]
        mock_get_client.return_value = client

        # 2026-08-07 is a Friday; the week's Monday is 2026-08-03
        result = mcp_server.cal_week(date="2026-08-07")

        client.get_calendar.assert_called_once_with(
            datetime(2026, 8, 3), datetime(2026, 8, 8)
        )
        assert result == [{"subject": "weekly"}]


class TestFindSlotTool:
    @patch("outlook_cli.mcp_server.get_client")
    def test_builds_freebusy_map_and_delegates_to_compute_free_slots(self, mock_get_client):
        client = MagicMock()
        client.get_freebusy.side_effect = ["0" * 336, "0" * 336]
        mock_get_client.return_value = client

        result = mcp_server.find_slot(
            attendees="tanaka@company.com;sato@company.com",
            duration=60, days=5, work_start=9, work_end=18,
        )

        assert client.get_freebusy.call_count == 2
        assert isinstance(result, list)
        if result:
            assert set(result[0].keys()) == {"start", "end", "tentative"}

    @patch("outlook_cli.mcp_server.get_client")
    def test_raises_when_no_freebusy_available(self, mock_get_client):
        client = MagicMock()
        client.get_freebusy.return_value = ""
        mock_get_client.return_value = client

        with pytest.raises(ValueError):
            mcp_server.find_slot(attendees="tanaka@company.com")
