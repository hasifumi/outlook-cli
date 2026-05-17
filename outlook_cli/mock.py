import json
from datetime import datetime, timedelta
from pathlib import Path

from .base import OutlookBase

# (weekday 0=Mon..4=Fri, start_hour, start_min, duration_min, busy_status)
_MOCK_BUSY_SCHEDULE = {
    "tanaka@company.com": [
        (0, 10,  0, 60, 2),  # Mon 10:00-11:00 Busy
        (0, 15,  0, 30, 2),  # Mon 15:00-15:30 Busy
        (2, 10,  0, 60, 2),  # Wed 10:00-11:00 Busy
        (3,  9, 30, 30, 1),  # Thu 09:30-10:00 Tentative
    ],
    "sato@company.com": [
        (2, 13,  0, 90, 2),  # Wed 13:00-14:30 Busy
        (1, 14,  0, 60, 1),  # Tue 14:00-15:00 Tentative
        (4, 10,  0, 60, 2),  # Fri 10:00-11:00 Busy
    ],
}

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

    def flagged_or_due(self, days: int = 7, folder: str = "inbox") -> list:
        cutoff = datetime.now() - timedelta(days=days)
        result = []
        for mail in self._get_folder_mails(folder):
            if datetime.fromisoformat(mail["date"]) < cutoff:
                continue
            flag_status = mail.get("flag_status", 0)
            due_date = mail.get("due_date")
            if flag_status == 1 or due_date:
                result.append({
                    "id":          mail["id"],
                    "subject":     mail["subject"],
                    "from":        mail["from"],
                    "from_name":   mail.get("from_name", ""),
                    "date":        mail["date"],
                    "flag_status": flag_status,
                    "due_date":    due_date,
                })
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def get_calendar(self, start: datetime, end: datetime) -> list:
        start_iso = start.isoformat()[:19]
        end_iso = end.isoformat()[:19]
        items = self._data.get("calendar", [])
        result = [
            item for item in items
            if item["start"] < end_iso and item["end"] > start_iso
        ]
        return sorted(result, key=lambda x: x["start"])

    def get_freebusy(self, email: str, start: datetime, minutes: int = 30) -> str:
        schedule = _MOCK_BUSY_SCHEDULE.get(email)
        total_slots = (7 * 24 * 60) // minutes
        slots = ["0"] * total_slots
        if schedule is None:
            return "".join(slots)
        for weekday, sh, sm, dur_min, status in schedule:
            for slot_offset in range(total_slots):
                slot_dt = start + timedelta(minutes=slot_offset * minutes)
                if slot_dt.weekday() != weekday:
                    continue
                slot_minutes_from_midnight = slot_dt.hour * 60 + slot_dt.minute
                event_start_min = sh * 60 + sm
                event_end_min = event_start_min + dur_min
                if event_start_min <= slot_minutes_from_midnight < event_end_min:
                    slots[slot_offset] = str(status)
        return "".join(slots)

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
