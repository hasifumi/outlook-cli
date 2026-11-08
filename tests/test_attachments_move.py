"""添付ファイル（一覧・保存）とメール移動のテスト"""
import json

import pytest
from click.testing import CliRunner

from outlook_cli._files import safe_filename, unique_path
from outlook_cli.cli import cli
from outlook_cli.mock import OutlookMock

ENV = {"OUTLOOK_MOCK": "1"}


class TestFileHelpers:
    def test_unique_path_returns_plain_name_when_free(self, tmp_path):
        assert unique_path(tmp_path, "a.xlsx") == tmp_path / "a.xlsx"

    def test_unique_path_numbers_on_collision(self, tmp_path):
        (tmp_path / "a.xlsx").write_text("x")
        assert unique_path(tmp_path, "a.xlsx") == tmp_path / "a_1.xlsx"
        (tmp_path / "a_1.xlsx").write_text("x")
        assert unique_path(tmp_path, "a.xlsx") == tmp_path / "a_2.xlsx"

    def test_safe_filename_replaces_illegal_characters(self):
        assert safe_filename('a/b\\c:d*e?f"g<h>i|j.txt') == "a_b_c_d_e_f_g_h_i_j.txt"

    def test_safe_filename_never_empty_or_dotted(self):
        assert safe_filename("") == "attachment"
        assert safe_filename("..") == "attachment"


class TestMockAttachments:
    def test_read_lists_attachment_names(self):
        client = OutlookMock()
        assert client.read("mock-002")["attachments"] == ["Q1レポート様式.xlsx"]
        assert client.read("mock-003")["attachments"] == ["経費精算書.xlsx", "領収書.pdf"]

    def test_read_without_attachments_is_empty_list(self):
        assert OutlookMock().read("mock-001")["attachments"] == []

    def test_list_with_body_includes_attachments(self):
        mails = OutlookMock().list_mails(limit=5, include_body=True)
        by_id = {m["id"]: m for m in mails}
        assert by_id["mock-003"]["attachments"] == ["経費精算書.xlsx", "領収書.pdf"]

    def test_list_without_body_has_no_attachment_names(self):
        mails = OutlookMock().list_mails(limit=5)
        assert all("attachments" not in m for m in mails)

    def test_save_attachment_writes_file(self, tmp_path):
        path = OutlookMock().save_mail_attachment("mock-002", "Q1レポート様式.xlsx", str(tmp_path))
        assert path == str(tmp_path / "Q1レポート様式.xlsx")
        assert (tmp_path / "Q1レポート様式.xlsx").stat().st_size > 0

    def test_save_attachment_creates_dest_dir(self, tmp_path):
        dest = tmp_path / "new" / "sub"
        path = OutlookMock().save_mail_attachment("mock-003", "領収書.pdf", str(dest))
        assert (dest / "領収書.pdf").exists()
        assert path == str(dest / "領収書.pdf")

    def test_save_attachment_never_overwrites(self, tmp_path):
        client = OutlookMock()
        first = client.save_mail_attachment("mock-002", "Q1レポート様式.xlsx", str(tmp_path))
        second = client.save_mail_attachment("mock-002", "Q1レポート様式.xlsx", str(tmp_path))
        assert first != second
        assert second.endswith("Q1レポート様式_1.xlsx")
        assert len(list(tmp_path.iterdir())) == 2

    def test_save_unknown_attachment_raises(self, tmp_path):
        with pytest.raises(ValueError):
            OutlookMock().save_mail_attachment("mock-002", "nothing.zip", str(tmp_path))

    def test_save_attachment_of_unknown_mail_raises(self, tmp_path):
        with pytest.raises(ValueError):
            OutlookMock().save_mail_attachment("no-such", "a.txt", str(tmp_path))

    def test_save_to_temp_returns_existing_file(self):
        path = OutlookMock().save_mail_attachment_to_temp("mock-002", "Q1レポート様式.xlsx")
        import os
        assert os.path.isfile(path)


class TestMockMove:
    def test_move_to_subfolder(self):
        client = OutlookMock()
        client.move_mail("mock-002", "プロジェクトA")
        assert all(m["id"] != "mock-002" for m in client.list_mails("inbox", limit=50))
        assert any(m["id"] == "mock-002" for m in client.list_mails("プロジェクトA", limit=50))

    def test_move_to_standard_folder(self):
        client = OutlookMock()
        client.move_mail("mock-001", "trash")
        assert any(m["id"] == "mock-001" for m in client.list_mails("trash", limit=50))

    def test_move_to_new_folder_creates_it(self):
        client = OutlookMock()
        client.move_mail("mock-004", "保管済み")
        assert "保管済み" in client.list_subfolders()

    def test_move_unknown_mail_raises(self):
        with pytest.raises(ValueError):
            OutlookMock().move_mail("no-such", "trash")

    def test_move_keeps_attachments(self):
        client = OutlookMock()
        client.move_mail("mock-003", "社内連絡")
        moved = client.read("mock-003")
        assert moved["attachments"] == ["経費精算書.xlsx", "領収書.pdf"]


class TestFoldersList:
    def test_list_subfolders_default_is_names(self):
        assert OutlookMock().list_subfolders() == ["プロジェクトA", "社内連絡"]

    def test_list_subfolders_recursive_flag_accepted(self):
        assert OutlookMock().list_subfolders(recursive=True) == ["プロジェクトA", "社内連絡"]


class TestCliAttachments:
    def test_attachments_lists_names_json(self):
        r = CliRunner().invoke(cli, ["attachments", "mock-003", "--json-output"], env=ENV)
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == ["経費精算書.xlsx", "領収書.pdf"]

    def test_attachments_text_mode(self):
        r = CliRunner().invoke(cli, ["attachments", "mock-002"], env=ENV)
        assert r.exit_code == 0, r.output
        assert "Q1レポート様式.xlsx" in r.output

    def test_attachments_none(self):
        r = CliRunner().invoke(cli, ["attachments", "mock-001"], env=ENV)
        assert r.exit_code == 0, r.output
        assert "添付ファイルなし" in r.output

    def test_save_attachment_one_file_json(self, tmp_path):
        r = CliRunner().invoke(
            cli,
            ["save-attachment", "mock-003", "領収書.pdf", "--dest", str(tmp_path), "--json-output"],
            env=ENV,
        )
        assert r.exit_code == 0, r.output
        saved = json.loads(r.output)
        assert saved == [{"filename": "領収書.pdf", "path": str(tmp_path / "領収書.pdf")}]
        assert (tmp_path / "領収書.pdf").exists()

    def test_save_attachment_all(self, tmp_path):
        r = CliRunner().invoke(
            cli, ["save-attachment", "mock-003", "--all", "--dest", str(tmp_path), "--json-output"],
            env=ENV,
        )
        assert r.exit_code == 0, r.output
        saved = json.loads(r.output)
        assert [s["filename"] for s in saved] == ["経費精算書.xlsx", "領収書.pdf"]
        assert {p.name for p in tmp_path.iterdir()} == {"経費精算書.xlsx", "領収書.pdf"}

    def test_save_attachment_requires_name_or_all(self, tmp_path):
        r = CliRunner().invoke(
            cli, ["save-attachment", "mock-003", "--dest", str(tmp_path)], env=ENV,
        )
        assert r.exit_code != 0
        assert "--all" in r.output

    def test_save_attachment_unknown_name_fails(self, tmp_path):
        r = CliRunner().invoke(
            cli, ["save-attachment", "mock-002", "nothing.zip", "--dest", str(tmp_path)], env=ENV,
        )
        assert r.exit_code != 0
        assert "nothing.zip" in r.output


class TestCliMoveAndFolders:
    def test_move_command(self):
        r = CliRunner().invoke(cli, ["move", "mock-002", "--folder", "プロジェクトA"], env=ENV)
        assert r.exit_code == 0, r.output
        assert "移動しました" in r.output

    def test_move_command_json_returns_id_and_folder(self):
        r = CliRunner().invoke(
            cli, ["move", "mock-002", "--folder", "プロジェクトA", "--json-output"], env=ENV,
        )
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == {"id": "mock-002", "folder": "プロジェクトA"}

    def test_move_unknown_mail_fails_cleanly(self):
        r = CliRunner().invoke(cli, ["move", "no-such", "--folder", "trash"], env=ENV)
        assert r.exit_code != 0
        assert "no-such" in r.output

    def test_folders_json(self):
        r = CliRunner().invoke(cli, ["folders", "--json-output"], env=ENV)
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == ["プロジェクトA", "社内連絡"]

    def test_folders_recursive_json(self):
        r = CliRunner().invoke(cli, ["folders", "--recursive", "--json-output"], env=ENV)
        assert r.exit_code == 0, r.output
        assert json.loads(r.output) == ["プロジェクトA", "社内連絡"]
