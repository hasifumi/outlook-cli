import hashlib
import json
import os
import re
from datetime import datetime, timedelta
from pathlib import Path

from ._files import safe_filename, unique_path
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
    "suzuki@company.com": [
        (3,  0,  0, 1440, 3),  # Thu 終日 Out of Office（休暇）
    ],
    "yamada@company.com": [
        (3, 14,  0, 30, 1),  # Thu 14:00-14:30 Tentative（室定例と一部重複、判断が割れるケース）
    ],
    "kimura@company.com": [
        (3, 14,  0, 60, 2),  # Thu 14:00-15:00 Busy（室定例と完全重複、確定予定とバッティング）
    ],
}

# 件名先頭の Re: / Fw: / 返信: / 転送: 等（全角コロン・繰り返しにも対応）
_SUBJECT_PREFIX = re.compile(r"^\s*(re|fw|fwd|返信|転送)\s*[:：]\s*", re.IGNORECASE)


def _conversation_topic(subject: str) -> str:
    """Outlook の ConversationTopic 相当（件名から返信・転送の接頭辞を除去）"""
    prev = None
    while prev != subject:
        prev = subject
        subject = _SUBJECT_PREFIX.sub("", subject)
    return subject.strip()


def _with_conversation(mail: dict, include_body: bool = True) -> dict:
    """スレッド整理用フィールドを付与したコピーを返す"""
    topic = _conversation_topic(mail.get("subject", ""))
    # 本文なしの一覧では添付ファイル名（"attachments"）も出さない（COM 側と揃える）
    result = {k: v for k, v in mail.items() if include_body or k not in ("body", "attachments")}
    if include_body:
        result["attachments"] = list(mail.get("attachments", []))
    result["conversation_topic"] = topic
    result["conversation_id"] = "conv-" + hashlib.md5(topic.encode("utf-8")).hexdigest()[:12]
    return result


MOCK_DATA_PATH = Path(__file__).parent.parent / "mock_data.json"

# Outlook の新規メールボックスが既定で持つ 6 色カテゴリ（ロケールや版で多少異なる）
_MOCK_CATEGORIES = ["赤カテゴリ", "オレンジカテゴリ", "黄カテゴリ", "緑カテゴリ", "青カテゴリ", "紫カテゴリ"]


class OutlookMock(OutlookBase):

    def __init__(self):
        with open(MOCK_DATA_PATH, encoding="utf-8") as f:
            self._data = json.load(f)
        # Outlook 本体を開く操作はモックでは実行できないので、呼び出し履歴だけ記録する
        self.opened_mail: list = []
        self.opened_mail_attachments: list = []
        self.opened_in_outlook: list = []
        self.opened_attachments: list = []

    def _get_folder_mails(self, folder: str) -> list:
        if folder in ("inbox", "sent", "drafts", "outbox", "trash"):
            return self._data.get(folder, [])
        return self._data.get("folders", {}).get(folder, [])

    def list_mails(self, folder: str = "inbox", limit: int = 20, include_body: bool = False) -> list:
        mails = sorted(self._get_folder_mails(folder), key=lambda m: m["date"], reverse=True)
        return [_with_conversation(m, include_body) for m in mails[:limit]]

    def search(self, keyword: str, days: int = 7, sender: str = None, include_body: bool = False) -> list:
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
                result.append(_with_conversation(mail, include_body))
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def read(self, mail_id: str) -> dict:
        all_mails = (
            self._data.get("inbox", [])
            + self._data.get("sent", [])
            + self._data.get("trash", [])
            + self._data.get("drafts", [])
            + self._data.get("outbox", [])
            + sum(self._data.get("folders", {}).values(), [])
        )
        for mail in all_mails:
            if mail["id"] == mail_id:
                return _with_conversation(mail)
        raise ValueError(f"メールが見つかりません: {mail_id}")

    def _new_mail(self, mail_id: str, to: str, subject: str, body: str, cc: str,
                  attachments: list[str] | None) -> dict:
        # 実際の Outlook に合わせ、添付はパスではなくファイル名で持つ
        names = [os.path.basename(p) for p in (attachments or [])]
        return {
            "id": mail_id,
            "subject": subject,
            "from": "you@company.com",
            "from_name": "自分",
            "to": to,
            "cc": cc,
            "date": datetime.now().isoformat(),
            "unread": False,
            "body": body,
            "attachments": names,
            "has_attachments": bool(names),
        }

    def send(self, to: str, subject: str, body: str, cc: str = "",
             attachments: list[str] | None = None) -> None:
        mail_id = f"mock-s{len(self._data.get('sent', [])) + 1:02d}"
        self._data.setdefault("sent", []).append(
            self._new_mail(mail_id, to, subject, body, cc, attachments))

    def reply(self, mail_id: str, body: str, reply_all: bool = False,
              attachments: list[str] | None = None) -> None:
        original = self.read(mail_id)
        self.send(
            to=original["from"],
            subject=f"Re: {original['subject']}",
            body=body,
            cc=original.get("cc", "") if reply_all else "",
            attachments=attachments,
        )

    def forward(self, mail_id: str, to: str, body: str = "",
                attachments: list[str] | None = None) -> None:
        original = self.read(mail_id)
        # 実際の Outlook の Forward() は元メールの添付も引き継ぐ
        combined = list(original.get("attachments", [])) + list(attachments or [])
        self.send(to=to, subject=f"Fwd: {original['subject']}", body=body,
                  attachments=combined or None)

    def save_draft(self, to: str, subject: str, body: str, cc: str = "",
                   attachments: list[str] | None = None) -> str:
        draft_id = f"mock-d{len(self._data.get('drafts', [])) + 1:02d}"
        self._data.setdefault("drafts", []).append(
            self._new_mail(draft_id, to, subject, body, cc, attachments))
        return draft_id

    def open_mail(self, mail_id: str | None = None) -> None:
        self.opened_mail.append(mail_id)

    def open_mail_attachment(self, mail_id: str, filename: str) -> None:
        self.opened_mail_attachments.append((mail_id, filename))

    def list_subfolders(self, recursive: bool = False) -> list[str]:
        # モックのフォルダは 1 階層（キー名がそのままパス）なので recursive でも同じ
        return list(self._data.get("folders", {}).keys())

    def save_mail_attachment(self, mail_id: str, filename: str, dest_dir: str) -> str:
        """モックの添付には実体が無いため、目印の内容を持つダミーファイルを書き出す"""
        mail = self.read(mail_id)
        if filename not in mail.get("attachments", []):
            raise ValueError(f"添付ファイルが見つかりません: {filename}")
        dest = Path(dest_dir)
        dest.mkdir(parents=True, exist_ok=True)
        path = unique_path(dest, safe_filename(filename))
        path.write_bytes(f"[mock attachment] {mail_id}: {filename}\n".encode("utf-8"))
        return str(path)

    def _pop_mail(self, mail_id: str):
        for folder_key in ("inbox", "sent", "drafts", "outbox", "trash"):
            mails = self._data.get(folder_key, [])
            for i, m in enumerate(mails):
                if m["id"] == mail_id:
                    return mails.pop(i)
        for mails in self._data.get("folders", {}).values():
            for i, m in enumerate(mails):
                if m["id"] == mail_id:
                    return mails.pop(i)
        return None

    def move_mail(self, mail_id: str, folder: str) -> str:
        mail = self._pop_mail(mail_id)
        if mail is None:
            raise ValueError(f"メールが見つかりません: {mail_id}")
        if folder in ("inbox", "sent", "drafts", "outbox", "trash"):
            self._data.setdefault(folder, []).append(mail)
        else:
            self._data.setdefault("folders", {}).setdefault(folder, []).append(mail)
        return mail_id

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

    @staticmethod
    def _event_view(item: dict) -> dict:
        """予定に body 等の既定値を補い、COM と同じ形（キー）で返す"""
        return {
            "body": "", "reminder_minutes": None, "category": "", "attachments": [],
            **item,
        }

    def get_calendar(self, start: datetime, end: datetime, mailbox: str | None = None) -> list:
        start_iso = start.isoformat()[:19]
        end_iso = end.isoformat()[:19]
        if mailbox:
            items = self._mock_calendar_for_mailbox(mailbox, start)
        else:
            items = self._data.get("calendar", [])
        result = [
            self._event_view(item) for item in items
            if item["start"] < end_iso and item["end"] > start_iso
        ]
        return sorted(result, key=lambda x: x["start"])

    def _mock_calendar_for_mailbox(self, mailbox: str, start: datetime) -> list[dict]:
        """他人の予定表のダミー。空き状況スケジュール（_MOCK_BUSY_SCHEDULE）をその週の予定として生成する"""
        schedule = _MOCK_BUSY_SCHEDULE.get(mailbox, [])
        monday = start - timedelta(days=start.weekday())
        monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
        items = []
        for i, (weekday, sh, sm, dur_min, status) in enumerate(schedule):
            ev_start = monday + timedelta(days=weekday, hours=sh, minutes=sm)
            ev_end = ev_start + timedelta(minutes=dur_min)
            items.append({
                "id": f"shared-{mailbox}-{i}",
                "subject": f"{mailbox}の予定",
                "start": ev_start.isoformat()[:19],
                "end": ev_end.isoformat()[:19],
                "location": "",
                "organizer": mailbox,
                "attendees": "",
                "all_day": dur_min >= 1440,
                "busy_status": status,
            })
        return items

    def get_event(self, event_id: str) -> dict:
        for item in self._data.get("calendar", []):
            if item["id"] == event_id:
                return self._event_view(item)
        raise ValueError(f"予定が見つかりません: {event_id}")

    def get_categories(self) -> list[str]:
        return list(_MOCK_CATEGORIES)

    def _next_event_id(self) -> str:
        used = [
            int(e["id"].rsplit("-", 1)[1]) for e in self._data.get("calendar", [])
            if e["id"].startswith("cal-mock-") and e["id"].rsplit("-", 1)[1].isdigit()
        ]
        return f"cal-mock-{max(used, default=0) + 1}"

    @staticmethod
    def _reminder(reminder_minutes):
        return reminder_minutes if reminder_minutes is not None and reminder_minutes >= 0 else None

    def create_event(self, subject: str, start: datetime, end: datetime,
                     attendees: list[str] | None = None, location: str = "",
                     body: str = "", busy_status: int = 2,
                     reminder_minutes: int | None = 15, category: str = "",
                     attachments: list[str] | None = None) -> str:
        event_id = self._next_event_id()
        self._data.setdefault("calendar", []).append({
            "id": event_id,
            "subject": subject,
            "start": start.isoformat()[:19],
            "end": end.isoformat()[:19],
            "location": location,
            "organizer": "自分",
            "attendees": ";".join(attendees) if attendees else "",
            "all_day": False,
            "busy_status": busy_status,
            "body": body,
            "reminder_minutes": self._reminder(reminder_minutes),
            "category": category,
            "attachments": [os.path.basename(p) for p in (attachments or [])],
        })
        return event_id

    def update_event(self, event_id: str, subject: str, start: datetime, end: datetime,
                     attendees: list[str] | None = None, location: str = "",
                     body: str = "", busy_status: int = 2,
                     reminder_minutes: int | None = 15, category: str = "",
                     attachments: list[str] | None = None) -> None:
        for event in self._data.get("calendar", []):
            if event["id"] != event_id:
                continue
            event.update({
                "subject": subject,
                "start": start.isoformat()[:19],
                "end": end.isoformat()[:19],
                "location": location,
                "busy_status": busy_status,
                "body": body,
                "reminder_minutes": self._reminder(reminder_minutes),
                "category": category,
            })
            if attendees:                                  # 実機と同じく None なら出席者は触らない
                event["attendees"] = ";".join(attendees)
            current = event.get("attachments", [])
            added = [os.path.basename(p) for p in (attachments or [])]
            event["attachments"] = current + [n for n in added if n not in current]
            return
        raise ValueError(f"予定が見つかりません: {event_id}")

    def delete_event(self, event_id: str) -> None:
        calendar = self._data.get("calendar", [])
        for i, event in enumerate(calendar):
            if event["id"] == event_id:
                calendar.pop(i)
                return
        raise ValueError(f"予定が見つかりません: {event_id}")

    def open_in_outlook(self, event_id: str | None = None) -> None:
        self.opened_in_outlook.append(event_id)

    def open_attachment(self, event_id: str, filename: str) -> None:
        self.opened_attachments.append((event_id, filename))

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
