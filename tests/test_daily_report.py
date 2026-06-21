from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from daily_report.collect import collect_calendar, collect_sent_mail, collect_activitywatch
from daily_report.report import save_report
from daily_report.llm import call_llm


class TestCollectCalendar:
    def _make_client(self, today_events=None, tomorrow_events=None):
        client = MagicMock()
        client.get_calendar.side_effect = [today_events or [], tomorrow_events or []]
        return client

    def test_includes_today_and_tomorrow_sections(self):
        client = self._make_client(
            today_events=[{
                "subject": "朝会", "start": "2026-06-21T09:00:00",
                "end": "2026-06-21T09:30:00", "location": "", "all_day": False,
            }],
            tomorrow_events=[{
                "subject": "週次MTG", "start": "2026-06-22T10:00:00",
                "end": "2026-06-22T11:00:00", "location": "", "all_day": False,
            }],
        )
        result = collect_calendar(client, date(2026, 6, 21))
        assert "朝会" in result
        assert "週次MTG" in result
        assert "今日" in result
        assert "明日" in result

    def test_empty_calendar_shows_placeholder(self):
        client = self._make_client()
        result = collect_calendar(client, date(2026, 6, 21))
        assert "予定なし" in result

    def test_all_day_event_shows_label(self):
        client = self._make_client(
            today_events=[{
                "subject": "創立記念日", "start": "2026-06-21T00:00:00",
                "end": "2026-06-22T00:00:00", "location": "", "all_day": True,
            }]
        )
        result = collect_calendar(client, date(2026, 6, 21))
        assert "終日" in result


class TestCollectSentMail:
    def test_formats_recipient_and_subject(self):
        client = MagicMock()
        client.sent_today.return_value = [{
            "subject": "プロジェクト進捗", "to": "boss@company.com",
            "date": "2026-06-21T14:30:00",
        }]
        result = collect_sent_mail(client, date(2026, 6, 21))
        assert "プロジェクト進捗" in result
        assert "boss@company.com" in result

    def test_empty_returns_placeholder(self):
        client = MagicMock()
        client.sent_today.return_value = []
        result = collect_sent_mail(client, date(2026, 6, 21))
        assert "送信なし" in result

    def test_passes_correct_date_to_client(self):
        client = MagicMock()
        client.sent_today.return_value = []
        collect_sent_mail(client, date(2026, 6, 21))
        client.sent_today.assert_called_once_with(date="2026-06-21")


class TestCollectActivitywatch:
    def _make_rest_mocks(self, events):
        buckets_resp = MagicMock()
        buckets_resp.raise_for_status = MagicMock()
        buckets_resp.json.return_value = [{"id": "aw-watcher-window_TESTPC"}]

        events_resp = MagicMock()
        events_resp.raise_for_status = MagicMock()
        events_resp.json.return_value = events
        return [buckets_resp, events_resp]

    def test_aggregates_by_app(self):
        events = [
            {"data": {"app": "Code", "title": "test.py"}, "duration": 3600},
            {"data": {"app": "Code", "title": "main.py"}, "duration": 1800},
            {"data": {"app": "Firefox", "title": "Google"}, "duration": 900},
        ]
        with patch("daily_report.collect.requests.get", side_effect=self._make_rest_mocks(events)):
            result = collect_activitywatch(date(2026, 6, 21))
        assert "Code" in result
        assert "Firefox" in result

    def test_returns_empty_when_aw_unavailable(self):
        with patch("daily_report.collect.requests.get", side_effect=Exception("connection refused")):
            result = collect_activitywatch(date(2026, 6, 21))
        assert isinstance(result, str)


class TestSaveReport:
    def test_creates_file_with_date(self, tmp_path):
        target = date(2026, 6, 21)
        path = save_report(str(tmp_path), "# 日報テスト", target)
        assert path.exists()
        assert path.name == "2026-06-21.md"
        assert "日報テスト" in path.read_text(encoding="utf-8")

    def test_custom_output_dir(self, tmp_path):
        custom_dir = tmp_path / "custom"
        custom_dir.mkdir()
        path = save_report(str(custom_dir), "テスト", date(2026, 6, 21))
        assert path.parent == custom_dir

    def test_creates_output_dir_if_missing(self, tmp_path):
        new_dir = tmp_path / "new_subdir"
        path = save_report(str(new_dir), "テスト", date(2026, 6, 21))
        assert path.exists()


class TestCallLlm:
    def test_sends_3_text_blocks(self):
        mock_resp = MagicMock()
        mock_resp.raise_for_status = MagicMock()
        mock_resp.json.return_value = {
            "choices": [{"message": {"content": "# 日報\n## やったこと\n- 作業"}}]
        }
        with patch("daily_report.llm.requests.post", return_value=mock_resp) as mock_post:
            result = call_llm(
                "http://localhost:8089/v1/chat/completions",
                "カレンダー内容",
                "送信メール内容",
                "AW内容",
            )
        payload = mock_post.call_args[1]["json"]
        user_msg = next(m for m in payload["messages"] if m["role"] == "user")
        assert isinstance(user_msg["content"], list)
        assert len(user_msg["content"]) == 3
        assert result == "# 日報\n## やったこと\n- 作業"

    def test_returns_llm_response_text(self):
        mock_resp = MagicMock()
        mock_resp.raise_for_status = MagicMock()
        mock_resp.json.return_value = {
            "choices": [{"message": {"content": "生成された日報テキスト"}}]
        }
        with patch("daily_report.llm.requests.post", return_value=mock_resp):
            result = call_llm("http://x", "a", "b", "c")
        assert result == "生成された日報テキスト"
