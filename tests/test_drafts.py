import pytest
from outlook_cli.mock import OutlookMock


def test_read_drafts():
    """drafts フォルダのメールを read で読み込めること"""
    client = OutlookMock()
    # 下書きデータを追加
    test_draft = {
        "id": "draft-001",
        "subject": "下書きテスト",
        "from": "me@company.com",
        "to": "you@company.com",
        "date": "2023-01-01T10:00:00",
        "unread": False,
        "body": "これは下書きです。",
    }
    client._data.setdefault("drafts", []).append(test_draft)

    # read で下書きを取得
    result = client.read("draft-001")

    # 正しく取得できていることを確認
    assert result["id"] == "draft-001"
    assert result["subject"] == "下書きテスト"
    assert "下書き" in result["body"]


def test_reply_on_draft():
    """drafts フォルダのメールに対して reply できること"""
    client = OutlookMock()
    # 下書きデータを追加
    test_draft = {
        "id": "draft-002",
        "subject": "リプライテスト",
        "from": "other@company.com",
        "to": "me@company.com",
        "date": "2023-01-01T10:00:00",
        "unread": False,
        "body": "元の本文。",
    }
    client._data.setdefault("drafts", []).append(test_draft)

    # reply を実行
    client.reply("draft-002", "返信本文")

    # エラーが発生しないことを確認 (送信済みフォルダにメールが追加されるはず)
    assert True
