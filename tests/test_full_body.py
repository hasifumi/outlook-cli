"""search / list の本文全文出力と、スレッド整理用フィールドのテスト"""
import json

from click.testing import CliRunner

from outlook_cli.cli import cli
from outlook_cli.mock import OutlookMock

# モックデータの日付は 2026-04 前後なので、十分大きい日数で検索する
ALL_DAYS = 36500


class TestMockSearchBody:
    def test_search_excludes_body_by_default(self):
        mails = OutlookMock().search("定例", days=ALL_DAYS)
        assert mails
        assert all("body" not in m for m in mails)

    def test_search_includes_full_body_when_requested(self):
        mails = OutlookMock().search("定例", days=ALL_DAYS, include_body=True)
        target = next(m for m in mails if m["id"] == "mock-001")
        assert target["body"] == (
            "お疲れ様です。来週月曜の定例会議ですが、場所が第2会議室に変更になりました。よろしくお願いします。"
        )


class TestMockListBody:
    def test_list_excludes_body_by_default(self):
        mails = OutlookMock().list_mails(folder="inbox", limit=5)
        assert all("body" not in m for m in mails)

    def test_list_includes_full_body_when_requested(self):
        mails = OutlookMock().list_mails(folder="inbox", limit=5, include_body=True)
        assert mails
        assert all(m["body"] for m in mails)


class TestConversationFields:
    def test_reply_shares_conversation_topic_with_original(self):
        client = OutlookMock()
        original = client.read("mock-001")      # 来週の定例会議について
        reply = client.read("mock-s01")         # Re: 来週の定例会議について
        assert original["conversation_topic"] == "来週の定例会議について"
        assert reply["conversation_topic"] == "来週の定例会議について"
        assert original["conversation_id"] == reply["conversation_id"]

    def test_fullwidth_colon_and_forward_prefix_are_stripped(self):
        client = OutlookMock()
        client._data["inbox"].append({
            "id": "mock-x01", "subject": "FW：RE: 承認依頼：経費精算",
            "from": "a@company.com", "to": "you@company.com",
            "date": "2026-04-25T10:00:00", "unread": False, "body": "転送です",
        })
        fwd = client.read("mock-x01")
        orig = client.read("mock-003")          # 承認依頼：経費精算
        assert fwd["conversation_topic"] == "承認依頼：経費精算"
        assert fwd["conversation_id"] == orig["conversation_id"]

    def test_different_threads_have_different_ids(self):
        client = OutlookMock()
        assert client.read("mock-001")["conversation_id"] != client.read("mock-002")["conversation_id"]

    def test_search_and_list_results_have_conversation_fields(self):
        client = OutlookMock()
        for m in client.search("定例", days=ALL_DAYS) + client.list_mails(limit=5):
            assert m["conversation_id"]
            assert m["conversation_topic"]


class TestCliWithBody:
    def test_search_with_body_json_contains_body(self):
        result = CliRunner().invoke(
            cli, ["search", "定例", "--days", str(ALL_DAYS), "--json-output", "--with-body"],
            env={"OUTLOOK_MOCK": "1"},
        )
        assert result.exit_code == 0, result.output
        mails = json.loads(result.output)
        assert any("第2会議室に変更" in m["body"] for m in mails)

    def test_list_with_body_json_contains_body(self):
        result = CliRunner().invoke(
            cli, ["list", "--limit", "3", "--json-output", "--with-body"],
            env={"OUTLOOK_MOCK": "1"},
        )
        assert result.exit_code == 0, result.output
        mails = json.loads(result.output)
        assert len(mails) == 3
        assert all("body" in m for m in mails)

    def test_search_with_body_text_mode_prints_body(self):
        result = CliRunner().invoke(
            cli, ["search", "定例", "--days", str(ALL_DAYS), "--with-body"],
            env={"OUTLOOK_MOCK": "1"},
        )
        assert result.exit_code == 0, result.output
        assert "第2会議室に変更" in result.output
