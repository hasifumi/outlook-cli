"""予定の作成・更新・削除、カテゴリ、他人の予定表（mailbox）のテスト"""
import json
import os
from datetime import datetime

import pytest
from click.testing import CliRunner

from outlook_cli import cli as cli_module
from outlook_cli.cli import cli
from outlook_cli.mock import OutlookMock

ENV = {"OUTLOOK_MOCK": "1"}
MONDAY = datetime(2026, 6, 22)          # 2026-06-22 は月曜日
WEEK_END = datetime(2026, 6, 27)


def _day(client, y, m, d):
    return client.get_calendar(datetime(y, m, d), datetime(y, m, d, 23, 59, 59))


class TestMockEventFields:
    def test_events_have_normalized_extra_fields(self):
        events = _day(OutlookMock(), 2026, 6, 21)
        assert events
        for e in events:
            assert "body" in e and "attachments" in e
            assert "category" in e and "reminder_minutes" in e

    def test_get_event_returns_event(self):
        assert OutlookMock().get_event("cal-001")["subject"] == "週次定例MTG"

    def test_get_event_unknown_raises(self):
        with pytest.raises(ValueError):
            OutlookMock().get_event("no-such")


class TestMockCreate:
    def test_create_event_appears_in_calendar(self):
        client = OutlookMock()
        event_id = client.create_event(
            "打ち合わせ", datetime(2026, 7, 1, 10, 0), datetime(2026, 7, 1, 11, 0),
            attendees=["a@x.com", "b@x.com"], location="会議室", body="議題", category="青カテゴリ",
            attachments=["/tmp/資料.pdf"],
        )
        e = client.get_event(event_id)
        assert e["subject"] == "打ち合わせ"
        assert e["start"] == "2026-07-01T10:00:00" and e["end"] == "2026-07-01T11:00:00"
        assert e["attendees"] == "a@x.com;b@x.com"
        assert e["location"] == "会議室" and e["body"] == "議題"
        assert e["category"] == "青カテゴリ"
        assert e["attachments"] == ["資料.pdf"]
        assert e["busy_status"] == 2 and e["reminder_minutes"] == 15
        assert any(x["id"] == event_id for x in _day(client, 2026, 7, 1))

    def test_negative_reminder_means_none(self):
        client = OutlookMock()
        event_id = client.create_event(
            "通知なし", datetime(2026, 7, 1, 10, 0), datetime(2026, 7, 1, 11, 0), reminder_minutes=-1)
        assert client.get_event(event_id)["reminder_minutes"] is None

    def test_ids_stay_unique_after_delete(self):
        client = OutlookMock()
        a = client.create_event("a", datetime(2026, 7, 1, 9), datetime(2026, 7, 1, 10))
        b = client.create_event("b", datetime(2026, 7, 1, 9), datetime(2026, 7, 1, 10))
        client.delete_event(a)
        c = client.create_event("c", datetime(2026, 7, 1, 9), datetime(2026, 7, 1, 10))
        assert len({a, b, c}) == 3


class TestMockUpdateDelete:
    def test_update_changes_fields(self):
        client = OutlookMock()
        client.update_event(
            "cal-002", "1on1 変更", datetime(2026, 6, 21, 16, 0), datetime(2026, 6, 21, 16, 30),
            location="A室", body="本文")
        e = client.get_event("cal-002")
        assert e["subject"] == "1on1 変更" and e["location"] == "A室"
        assert e["start"] == "2026-06-21T16:00:00"

    def test_update_appends_attachments_like_outlook(self):
        client = OutlookMock()
        client.update_event("cal-002", "x", datetime(2026, 6, 21, 15), datetime(2026, 6, 21, 16),
                            attachments=["/tmp/a.pdf"])
        client.update_event("cal-002", "x", datetime(2026, 6, 21, 15), datetime(2026, 6, 21, 16),
                            attachments=["/tmp/b.pdf"])
        assert client.get_event("cal-002")["attachments"] == ["a.pdf", "b.pdf"]

    def test_update_without_attendees_keeps_existing_ones(self):
        client = OutlookMock()
        before = client.get_event("cal-001")["attendees"]
        assert before
        client.update_event("cal-001", "週次定例MTG", datetime(2026, 6, 21, 10), datetime(2026, 6, 21, 11),
                            location="別室")
        assert client.get_event("cal-001")["attendees"] == before

    def test_update_with_attendees_replaces_them(self):
        client = OutlookMock()
        client.update_event("cal-001", "週次定例MTG", datetime(2026, 6, 21, 10), datetime(2026, 6, 21, 11),
                            attendees=["z@x.com"])
        assert client.get_event("cal-001")["attendees"] == "z@x.com"

    def test_update_unknown_raises(self):
        with pytest.raises(ValueError):
            OutlookMock().update_event("no-such", "x", datetime(2026, 6, 21, 15), datetime(2026, 6, 21, 16))

    def test_delete_removes_event(self):
        client = OutlookMock()
        client.delete_event("cal-002")
        assert all(e["id"] != "cal-002" for e in _day(client, 2026, 6, 21))

    def test_delete_unknown_raises(self):
        with pytest.raises(ValueError):
            OutlookMock().delete_event("no-such")


class TestMockMisc:
    def test_categories(self):
        cats = OutlookMock().get_categories()
        assert cats and all(isinstance(c, str) for c in cats)

    def test_mailbox_calendar_uses_busy_schedule(self):
        events = OutlookMock().get_calendar(MONDAY, WEEK_END, mailbox="tanaka@company.com")
        assert len(events) == 4
        assert all(e["subject"] == "tanaka@company.comの予定" for e in events)
        assert events == sorted(events, key=lambda e: e["start"])

    def test_unknown_mailbox_is_empty(self):
        assert OutlookMock().get_calendar(MONDAY, WEEK_END, mailbox="nobody@x.com") == []

    def test_open_in_outlook_and_attachment_are_recorded(self):
        client = OutlookMock()
        client.open_in_outlook("cal-001")
        client.open_in_outlook()
        client.open_attachment("cal-001", "a.pdf")
        assert client.opened_in_outlook == ["cal-001", None]
        assert client.opened_attachments == [("cal-001", "a.pdf")]


class CalSpy:
    def __init__(self, existing=None):
        self.calls = []
        self.existing = existing

    def get_calendar(self, start, end, mailbox=None):
        self.calls.append(("get_calendar", {"start": start, "end": end, "mailbox": mailbox}))
        return []

    def get_event(self, event_id):
        if self.existing is None:
            raise ValueError(f"予定が見つかりません: {event_id}")
        return dict(self.existing)

    def create_event(self, **kw):
        self.calls.append(("create_event", kw))
        return "ev-new"

    def update_event(self, **kw):
        self.calls.append(("update_event", kw))

    def delete_event(self, event_id):
        self.calls.append(("delete_event", {"event_id": event_id}))


EXISTING = {
    "id": "e1", "subject": "定例", "start": "2026-06-23T10:00:00", "end": "2026-06-23T11:00:00",
    "location": "A室", "organizer": "me", "attendees": "a@x.com;b@x.com", "all_day": False,
    "busy_status": 2, "body": "議題", "reminder_minutes": 10, "category": "青", "attachments": ["x.pdf"],
}


@pytest.fixture
def spy(monkeypatch):
    s = CalSpy(existing=EXISTING)
    monkeypatch.setattr(cli_module, "get_client", lambda: s)
    return s


class TestCliCreate:
    def test_create_parses_options(self, spy, tmp_path):
        f = tmp_path / "資料.pdf"
        f.write_text("x")
        r = CliRunner().invoke(cli, [
            "cal", "create", "--subject", "打ち合わせ",
            "--start", "2026-07-01T10:00", "--end", "2026-07-01T11:30",
            "--attendee", "a@x.com", "--attendee", "b@x.com",
            "--location", "A室", "--body", "議題", "--category", "青", "--attach", str(f),
        ])
        assert r.exit_code == 0, r.output
        name, kw = spy.calls[0]
        assert name == "create_event"
        assert kw["subject"] == "打ち合わせ"
        assert kw["start"] == datetime(2026, 7, 1, 10, 0) and kw["end"] == datetime(2026, 7, 1, 11, 30)
        assert kw["attendees"] == ["a@x.com", "b@x.com"]
        assert kw["busy_status"] == 2 and kw["reminder_minutes"] == 15
        assert os.path.samefile(kw["attachments"][0], f)

    def test_create_without_attendees_passes_none(self, spy):
        r = CliRunner().invoke(cli, [
            "cal", "create", "--subject", "個人", "--start", "2026-07-01 10:00", "--end", "2026-07-01 11:00"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["attendees"] is None

    def test_create_json_returns_id(self, spy):
        r = CliRunner().invoke(cli, [
            "cal", "create", "--subject", "s", "--start", "2026-07-01T10:00",
            "--end", "2026-07-01T11:00", "--json-output"])
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == {"id": "ev-new"}

    def test_create_rejects_end_before_start(self, spy):
        r = CliRunner().invoke(cli, [
            "cal", "create", "--subject", "s", "--start", "2026-07-01T11:00", "--end", "2026-07-01T10:00"])
        assert r.exit_code != 0
        assert spy.calls == []

    def test_create_rejects_bad_datetime(self, spy):
        r = CliRunner().invoke(cli, [
            "cal", "create", "--subject", "s", "--start", "明日の10時", "--end", "2026-07-01T11:00"])
        assert r.exit_code != 0
        assert spy.calls == []

    def test_create_on_mock_prints_id(self):
        r = CliRunner().invoke(cli, [
            "cal", "create", "--subject", "s", "--start", "2026-07-01T10:00",
            "--end", "2026-07-01T11:00", "--json-output"], env=ENV)
        assert r.exit_code == 0, r.output
        assert json.loads(r.output)["id"].startswith("cal-mock-")


class TestCliUpdate:
    def test_update_changes_only_given_fields(self, spy):
        r = CliRunner().invoke(cli, ["cal", "update", "e1", "--location", "B室"])
        assert r.exit_code == 0, r.output
        name, kw = spy.calls[0]
        assert name == "update_event"
        assert kw["event_id"] == "e1" and kw["location"] == "B室"
        assert kw["subject"] == "定例" and kw["body"] == "議題" and kw["category"] == "青"
        assert kw["start"] == datetime(2026, 6, 23, 10, 0) and kw["end"] == datetime(2026, 6, 23, 11, 0)
        # 出席者は触らない（None なら通知も送らない。実機の RequiredAttendees は表示名で、再設定は危険）
        assert kw["attendees"] is None
        assert kw["reminder_minutes"] == 10 and kw["busy_status"] == 2
        assert kw["attachments"] is None

    def test_update_start_only_keeps_duration(self, spy):
        r = CliRunner().invoke(cli, ["cal", "update", "e1", "--start", "2026-06-24T13:00"])
        assert r.exit_code == 0, r.output
        kw = spy.calls[0][1]
        assert kw["start"] == datetime(2026, 6, 24, 13, 0)
        assert kw["end"] == datetime(2026, 6, 24, 14, 0)        # 元の 1 時間を保つ

    def test_update_attendees_replace(self, spy):
        r = CliRunner().invoke(cli, ["cal", "update", "e1", "--attendee", "c@x.com"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["attendees"] == ["c@x.com"]

    def test_update_reminder_can_be_turned_off(self, spy):
        r = CliRunner().invoke(cli, ["cal", "update", "e1", "--reminder", "-1"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["reminder_minutes"] == -1

    def test_update_unknown_event_fails_cleanly(self, monkeypatch):
        monkeypatch.setattr(cli_module, "get_client", lambda: CalSpy(existing=None))
        r = CliRunner().invoke(cli, ["cal", "update", "no-such", "--location", "x"])
        assert r.exit_code != 0
        assert "no-such" in r.output

    def test_update_rejects_end_before_start(self, spy):
        r = CliRunner().invoke(cli, ["cal", "update", "e1", "--end", "2026-06-23T09:00"])
        assert r.exit_code != 0
        assert spy.calls == []


class TestCliDelete:
    def test_delete_requires_yes(self, spy):
        r = CliRunner().invoke(cli, ["cal", "delete", "e1"])
        assert r.exit_code != 0
        assert "--yes" in r.output
        assert spy.calls == []

    def test_delete_with_yes(self, spy):
        r = CliRunner().invoke(cli, ["cal", "delete", "e1", "--yes"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0] == ("delete_event", {"event_id": "e1"})

    def test_delete_unknown_on_mock_fails_cleanly(self):
        r = CliRunner().invoke(cli, ["cal", "delete", "no-such", "--yes"], env=ENV)
        assert r.exit_code != 0
        assert "no-such" in r.output


class TestCliMailboxAndCategories:
    def test_cal_today_passes_mailbox(self, spy):
        r = CliRunner().invoke(cli, ["cal", "today", "--date", "2026-06-22", "--mailbox", "tanaka@company.com"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["mailbox"] == "tanaka@company.com"

    def test_cal_week_passes_mailbox(self, spy):
        r = CliRunner().invoke(cli, ["cal", "week", "--date", "2026-06-22", "--mailbox", "tanaka@company.com"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["mailbox"] == "tanaka@company.com"

    def test_cal_week_for_mailbox_on_mock(self):
        r = CliRunner().invoke(
            cli, ["cal", "week", "--date", "2026-06-22", "--mailbox", "tanaka@company.com", "--json-output"],
            env=ENV)
        assert r.exit_code == 0, r.output
        assert len(json.loads(r.output)) == 4

    def test_cal_today_without_mailbox_is_unchanged(self, spy):
        r = CliRunner().invoke(cli, ["cal", "today", "--date", "2026-06-22"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["mailbox"] is None

    def test_categories_json(self):
        r = CliRunner().invoke(cli, ["categories", "--json-output"], env=ENV)
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == OutlookMock().get_categories()
