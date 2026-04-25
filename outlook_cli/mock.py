import json
from datetime import datetime, timedelta
from pathlib import Path

from .base import OutlookBase

MOCK_DATA_PATH = Path(__file__).parent.parent / "mock_data.json"


class OutlookMock(OutlookBase):

    def __init__(self):
        with open(MOCK_DATA_PATH, encoding="utf-8") as f:
            self._data = json.load(f)

    def _get_folder_mails(self, folder: str) -> list:
        if folder in ("inbox", "sent", "drafts", "trash"):
            return self._data.get(folder, [])
        return self._data.get("folders", {}).get(folder, [])

    def list_mails(self, folder: str = "inbox", limit: int = 20) -> list:
        mails = sorted(self._get_folder_mails(folder), key=lambda m: m["date"], reverse=True)
        return [{k: v for k, v in m.items() if k != "body"} for m in mails[:limit]]

    def search(self, keyword: str, days: int = 7, sender: str = None) -> list:
        cutoff = datetime.now() - timedelta(days=days)
        keyword_lower = keyword.lower()
        result = []
        all_mails = (
            self._data.get("inbox", [])
            + self._data.get("sent", [])
            + sum(self._data.get("folders", {}).values(), [])
        )
        for mail in all_mails:
            if datetime.fromisoformat(mail["date"]) < cutoff:
                continue
            if sender and sender.lower() not in mail.get("from", "").lower():
                continue
            if keyword_lower in mail.get("subject", "").lower() or keyword_lower in mail.get("body", "").lower():
                result.append({k: v for k, v in mail.items() if k != "body"})
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def read(self, mail_id: str) -> dict:
        all_mails = (
            self._data.get("inbox", [])
            + self._data.get("sent", [])
            + self._data.get("trash", [])
            + sum(self._data.get("folders", {}).values(), [])
        )
        for mail in all_mails:
            if mail["id"] == mail_id:
                return mail
        raise ValueError(f"メールが見つかりません: {mail_id}")

    def send(self, to: str, subject: str, body: str) -> None:
        mail = {
            "id": f"mock-s{len(self._data.get('sent', [])) + 1:02d}",
            "subject": subject,
            "from": "you@company.com",
            "from_name": "自分",
            "to": to,
            "cc": "",
            "date": datetime.now().isoformat(),
            "unread": False,
            "body": body,
        }
        self._data.setdefault("sent", []).append(mail)

    def reply(self, mail_id: str, body: str) -> None:
        original = self.read(mail_id)
        self.send(
            to=original["from"],
            subject=f"Re: {original['subject']}",
            body=body,
        )

    def list_subfolders(self) -> list[str]:
        return list(self._data.get("folders", {}).keys())

    def get_contacts(self) -> list[dict]:
        return self._data.get("contacts", [])

    def delete(self, mail_id: str) -> None:
        for folder_key in ("inbox", "sent", "drafts"):
            mails = self._data.get(folder_key, [])
            for i, m in enumerate(mails):
                if m["id"] == mail_id:
                    self._data.setdefault("trash", []).append(mails.pop(i))
                    return
        for mails in self._data.get("folders", {}).values():
            for i, m in enumerate(mails):
                if m["id"] == mail_id:
                    self._data.setdefault("trash", []).append(mails.pop(i))
                    return

    def get_unread_count(self, folder: str) -> int:
        return sum(1 for m in self._get_folder_mails(folder) if m.get("unread"))

    def unread_summary(self, limit: int = 10, folder: str = "inbox") -> list:
        mails = sorted(
            [m for m in self._get_folder_mails(folder) if m.get("unread")],
            key=lambda m: m["date"], reverse=True,
        )
        return [
            {
                "subject":   m["subject"],
                "from":      m["from"],
                "from_name": m.get("from_name", ""),
                "date":      m["date"],
                "preview":   m.get("body", "")[:100],
            }
            for m in mails[:limit]
        ]

    def sent_today(self, date: str = None) -> list:
        target = date or datetime.now().date().isoformat()
        mails = [m for m in self._data.get("sent", []) if m["date"].startswith(target)]
        return sorted(mails, key=lambda m: m["date"], reverse=True)

    def unread_count(self, folder: str = None) -> dict:
        result = {}
        if folder is None:
            for f in ("inbox", "sent", "drafts", "trash"):
                count = self.get_unread_count(f)
                if count > 0 or f == "inbox":
                    result[f] = count
            for subfolder in self._data.get("folders", {}):
                count = self.get_unread_count(subfolder)
                if count > 0:
                    result[subfolder] = count
        else:
            result[folder] = self.get_unread_count(folder)
        result["total"] = sum(v for k, v in result.items() if k != "total")
        return result
