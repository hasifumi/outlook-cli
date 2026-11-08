"""send の cc・添付、reply_all、forward、save_draft、Outlook 本体で開く操作のテスト"""
import json
import os

import pytest
from click.testing import CliRunner

from outlook_cli import cli as cli_module
from outlook_cli.cli import cli
from outlook_cli.mock import OutlookMock

ENV = {"OUTLOOK_MOCK": "1"}


class TestMockSend:
    def test_send_with_cc_and_attachments(self):
        client = OutlookMock()
        client.send("a@x.com", "件名", "本文", cc="c@x.com",
                    attachments=[os.path.join("C:\\", "tmp", "報告.xlsx")])
        sent = client._data["sent"][-1]
        assert sent["cc"] == "c@x.com"
        assert sent["attachments"] == ["報告.xlsx"]      # パスではなくファイル名で持つ
        assert sent["has_attachments"] is True

    def test_send_defaults_keep_old_behaviour(self):
        client = OutlookMock()
        client.send("a@x.com", "件名", "本文")
        sent = client._data["sent"][-1]
        assert sent["cc"] == ""
        assert sent["attachments"] == []
        assert sent["has_attachments"] is False


class TestMockReply:
    def test_reply_default_has_no_cc(self):
        client = OutlookMock()
        client._data["inbox"].append({
            "id": "x1", "subject": "相談", "from": "a@x.com", "to": "you@company.com",
            "cc": "boss@x.com", "date": "2026-04-25T10:00:00", "unread": False, "body": "b",
        })
        client.reply("x1", "返信")
        sent = client._data["sent"][-1]
        assert sent["to"] == "a@x.com"
        assert sent["subject"] == "Re: 相談"
        assert sent["cc"] == ""

    def test_reply_all_keeps_cc(self):
        client = OutlookMock()
        client._data["inbox"].append({
            "id": "x1", "subject": "相談", "from": "a@x.com", "to": "you@company.com",
            "cc": "boss@x.com", "date": "2026-04-25T10:00:00", "unread": False, "body": "b",
        })
        client.reply("x1", "返信", reply_all=True)
        assert client._data["sent"][-1]["cc"] == "boss@x.com"

    def test_reply_with_attachment(self):
        client = OutlookMock()
        client.reply("mock-001", "返信", attachments=["/tmp/資料.pdf"])
        assert client._data["sent"][-1]["attachments"] == ["資料.pdf"]


class TestMockForward:
    def test_forward_carries_original_attachments(self):
        client = OutlookMock()
        client.forward("mock-003", to="boss@x.com", body="ご確認ください")
        sent = client._data["sent"][-1]
        assert sent["to"] == "boss@x.com"
        assert sent["subject"] == "Fwd: 承認依頼：経費精算"
        assert sent["attachments"] == ["経費精算書.xlsx", "領収書.pdf"]

    def test_forward_adds_new_attachment(self):
        client = OutlookMock()
        client.forward("mock-002", to="boss@x.com", attachments=["/tmp/追加.docx"])
        assert client._data["sent"][-1]["attachments"] == ["Q1レポート様式.xlsx", "追加.docx"]

    def test_forward_unknown_mail_raises(self):
        with pytest.raises(ValueError):
            OutlookMock().forward("no-such", to="a@x.com")


class TestMockDraft:
    def test_save_draft_returns_id_and_is_readable(self):
        client = OutlookMock()
        draft_id = client.save_draft("a@x.com", "下書き件名", "下書き本文", cc="c@x.com",
                                     attachments=["/tmp/案.docx"])
        draft = client.read(draft_id)
        assert draft["subject"] == "下書き件名"
        assert draft["cc"] == "c@x.com"
        assert draft["attachments"] == ["案.docx"]
        assert any(m["id"] == draft_id for m in client.list_mails("drafts"))

    def test_draft_does_not_go_to_sent(self):
        client = OutlookMock()
        before = len(client._data["sent"])
        client.save_draft("a@x.com", "s", "b")
        assert len(client._data["sent"]) == before

    def test_draft_ids_are_unique(self):
        client = OutlookMock()
        assert client.save_draft("a@x.com", "s", "b") != client.save_draft("a@x.com", "s", "b")


class TestMockOpen:
    def test_open_mail_is_recorded(self):
        client = OutlookMock()
        client.open_mail("mock-001")
        client.open_mail()
        assert client.opened_mail == ["mock-001", None]

    def test_open_mail_attachment_is_recorded(self):
        client = OutlookMock()
        client.open_mail_attachment("mock-002", "Q1レポート様式.xlsx")
        assert client.opened_mail_attachments == [("mock-002", "Q1レポート様式.xlsx")]


class Spy:
    """CLI が client に何を渡したかを記録する偽クライアント"""

    def __init__(self):
        self.calls = []

    def send(self, **kw):
        self.calls.append(("send", kw))

    def reply(self, **kw):
        self.calls.append(("reply", kw))

    def forward(self, **kw):
        self.calls.append(("forward", kw))

    def save_draft(self, **kw):
        self.calls.append(("save_draft", kw))
        return "draft-id-1"

    def open_mail(self, mail_id=None):
        self.calls.append(("open_mail", {"mail_id": mail_id}))


@pytest.fixture
def spy(monkeypatch):
    s = Spy()
    monkeypatch.setattr(cli_module, "get_client", lambda: s)
    return s


class TestCliSend:
    def test_send_passes_cc_and_absolute_attachment_paths(self, spy, tmp_path):
        f = tmp_path / "報告.xlsx"
        f.write_text("x")
        r = CliRunner().invoke(cli, [
            "send", "--to", "a@x.com", "--subject", "s", "--body", "b",
            "--cc", "c@x.com", "--attach", str(f),
        ])
        assert r.exit_code == 0, r.output
        name, kw = spy.calls[0]
        assert name == "send"
        assert kw["cc"] == "c@x.com"
        assert len(kw["attachments"]) == 1
        assert os.path.isabs(kw["attachments"][0])
        assert os.path.samefile(kw["attachments"][0], f)    # 短縮名の展開差を許す

    def test_send_without_new_options_still_works(self, spy):
        r = CliRunner().invoke(cli, ["send", "--to", "a@x.com", "--subject", "s", "--body", "b"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["cc"] == ""
        assert spy.calls[0][1]["attachments"] is None

    def test_send_missing_attachment_file_is_rejected_before_sending(self, spy, tmp_path):
        r = CliRunner().invoke(cli, [
            "send", "--to", "a@x.com", "--subject", "s", "--body", "b",
            "--attach", str(tmp_path / "ない.xlsx"),
        ])
        assert r.exit_code != 0
        assert spy.calls == []                       # 送信されていない

    def test_send_multiple_attachments(self, spy, tmp_path):
        a, b = tmp_path / "a.txt", tmp_path / "b.txt"
        a.write_text("1")
        b.write_text("2")
        r = CliRunner().invoke(cli, [
            "send", "--to", "a@x.com", "--subject", "s", "--body", "b",
            "--attach", str(a), "--attach", str(b),
        ])
        assert r.exit_code == 0, r.output
        got = spy.calls[0][1]["attachments"]
        assert len(got) == 2
        assert os.path.samefile(got[0], a) and os.path.samefile(got[1], b)   # 指定順を保つ


class TestCliReplyForwardDraft:
    def test_reply_all_flag(self, spy):
        r = CliRunner().invoke(cli, ["reply", "m1", "--body", "b", "--all"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["reply_all"] is True

    def test_reply_default_is_not_reply_all(self, spy):
        r = CliRunner().invoke(cli, ["reply", "m1", "--body", "b"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0][1]["reply_all"] is False

    def test_forward(self, spy):
        r = CliRunner().invoke(cli, ["forward", "m1", "--to", "a@x.com", "--body", "ご確認"])
        assert r.exit_code == 0, r.output
        name, kw = spy.calls[0]
        assert name == "forward"
        assert kw["mail_id"] == "m1" and kw["to"] == "a@x.com" and kw["body"] == "ご確認"
        assert "転送しました" in r.output

    def test_draft_json_returns_id(self, spy):
        r = CliRunner().invoke(cli, [
            "draft", "--to", "a@x.com", "--subject", "s", "--body", "b", "--json-output",
        ])
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == {"id": "draft-id-1"}

    def test_draft_text_mode(self, spy):
        r = CliRunner().invoke(cli, ["draft", "--to", "a@x.com", "--subject", "s", "--body", "b"])
        assert r.exit_code == 0, r.output
        assert "下書きに保存しました" in r.output

    def test_open_with_id(self, spy):
        r = CliRunner().invoke(cli, ["open", "m1"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0] == ("open_mail", {"mail_id": "m1"})

    def test_open_without_id_opens_inbox(self, spy):
        r = CliRunner().invoke(cli, ["open"])
        assert r.exit_code == 0, r.output
        assert spy.calls[0] == ("open_mail", {"mail_id": None})


class TestCliOnMock:
    def test_forward_on_mock_succeeds(self):
        r = CliRunner().invoke(cli, ["forward", "mock-003", "--to", "boss@x.com"], env=ENV)
        assert r.exit_code == 0, r.output

    def test_forward_unknown_mail_fails_cleanly(self):
        r = CliRunner().invoke(cli, ["forward", "no-such", "--to", "boss@x.com"], env=ENV)
        assert r.exit_code != 0
        assert "no-such" in r.output

    def test_draft_on_mock_prints_id(self):
        r = CliRunner().invoke(
            cli, ["draft", "--to", "a@x.com", "--subject", "s", "--body", "b", "--json-output"],
            env=ENV,
        )
        assert r.exit_code == 0, r.output
        assert json.loads(r.output)["id"].startswith("mock-d")
