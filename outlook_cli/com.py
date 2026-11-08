import os
import sys
from datetime import datetime, timedelta

if sys.platform != "win32":
    raise ImportError("OutlookCOMはWindowsのみ対応しています")

import win32com.client

from pathlib import Path

from ._files import safe_filename, unique_path
from .base import OutlookBase

FOLDER_IDS = {
    "inbox":  6,
    "sent":   5,
    "drafts": 16,
    "outbox": 4,
    "trash":  3,
}


class OutlookCOM(OutlookBase):

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def _get_folder(self, folder: str):
        if folder in FOLDER_IDS:
            return self.namespace.GetDefaultFolder(FOLDER_IDS[folder])
        # 受信トレイ配下のフォルダ。"親/子" のように "/" で階層を指定できる
        current = self.namespace.GetDefaultFolder(6)
        for part in folder.split("/"):
            found = None
            for f in current.Folders:
                if f.Name == part:
                    found = f
                    break
            if found is None:
                raise ValueError(f"フォルダが見つかりません: {folder}")
            current = found
        return current

    @staticmethod
    def _attachment_names(item) -> list:
        names = []
        try:
            for att in item.Attachments:
                try:
                    names.append(att.FileName)
                except Exception:
                    continue
        except Exception:
            pass
        return names

    def _mail_to_dict(self, mail, include_body: bool = False) -> dict:
        result = {
            "id":        mail.EntryID,
            "subject":   mail.Subject,
            "from":      mail.SenderEmailAddress,
            "from_name": mail.SenderName,
            "to":        mail.To,
            "cc":        mail.CC,
            "date":      mail.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S"),
            "unread":          mail.UnRead,
            "has_attachments": mail.Attachments.Count > 0,
            "conversation_id":    getattr(mail, "ConversationID", "") or "",
            "conversation_topic": getattr(mail, "ConversationTopic", "") or mail.Subject,
        }
        if include_body:
            result["body"] = mail.Body
            result["attachments"] = self._attachment_names(mail)
        return result

    def list_mails(self, folder: str = "inbox", limit: int = 20, include_body: bool = False) -> list:
        items = self._get_folder(folder).Items
        items.Sort("[ReceivedTime]", True)
        result = []
        for i, mail in enumerate(items):
            if i >= limit:
                break
            try:
                result.append(self._mail_to_dict(mail, include_body=include_body))
            except Exception:
                continue
        return result

    def search(self, keyword: str, days: int = 7, sender: str = None, include_body: bool = False) -> list:
        cutoff = (datetime.now() - timedelta(days=days)).strftime("%m/%d/%Y")
        folder = self._get_folder("inbox")
        items = folder.Items
        filter_str = f"[ReceivedTime] >= '{cutoff}'"
        if sender:
            filter_str += f" AND [SenderEmailAddress] = '{sender}'"
        restricted = items.Restrict(filter_str)
        keyword_lower = keyword.lower()
        result = []
        for mail in restricted:
            try:
                if keyword_lower in mail.Subject.lower() or keyword_lower in mail.Body.lower():
                    result.append(self._mail_to_dict(mail, include_body=include_body))
            except Exception:
                continue
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def read(self, mail_id: str) -> dict:
        mail = self.namespace.GetItemFromID(mail_id)
        return self._mail_to_dict(mail, include_body=True)

    @staticmethod
    def _attach_files(item, attachments) -> None:
        for path in attachments or []:
            item.Attachments.Add(path)

    def send(self, to: str, subject: str, body: str, cc: str = "",
             attachments: list[str] | None = None) -> None:
        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.Body = body
        self._attach_files(mail, attachments)
        mail.Send()

    def reply(self, mail_id: str, body: str, reply_all: bool = False,
              attachments: list[str] | None = None) -> None:
        original = self.namespace.GetItemFromID(mail_id)
        reply = original.ReplyAll() if reply_all else original.Reply()
        reply.Body = body + "\n\n" + reply.Body
        self._attach_files(reply, attachments)
        reply.Send()

    def forward(self, mail_id: str, to: str, body: str = "",
                attachments: list[str] | None = None) -> None:
        original = self.namespace.GetItemFromID(mail_id)
        fwd = original.Forward()      # 元メールの添付は引き継がれる
        fwd.To = to
        fwd.Body = body + "\n\n" + fwd.Body
        self._attach_files(fwd, attachments)
        fwd.Send()

    def save_draft(self, to: str, subject: str, body: str, cc: str = "",
                   attachments: list[str] | None = None) -> str:
        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.Subject = subject
        mail.Body = body
        self._attach_files(mail, attachments)
        mail.Save()
        return mail.EntryID

    def open_mail(self, mail_id: str | None = None) -> None:
        if mail_id:
            self.namespace.GetItemFromID(mail_id).Display()
        else:
            self.namespace.GetDefaultFolder(6).Display()

    def open_mail_attachment(self, mail_id: str, filename: str) -> None:
        os.startfile(self.save_mail_attachment_to_temp(mail_id, filename))

    def _walk_folders(self, folder, prefix: str = "") -> list[str]:
        paths = []
        for f in folder.Folders:
            path = f"{prefix}{f.Name}"
            paths.append(path)
            paths.extend(self._walk_folders(f, prefix=f"{path}/"))
        return paths

    def list_subfolders(self, recursive: bool = False) -> list[str]:
        inbox = self.namespace.GetDefaultFolder(6)
        if recursive:
            return self._walk_folders(inbox)
        return [f.Name for f in inbox.Folders]

    def save_mail_attachment(self, mail_id: str, filename: str, dest_dir: str) -> str:
        mail = self.namespace.GetItemFromID(mail_id)
        dest = Path(dest_dir)
        dest.mkdir(parents=True, exist_ok=True)
        for att in mail.Attachments:
            if att.FileName == filename:
                path = unique_path(dest, safe_filename(att.FileName))
                att.SaveAsFile(str(path.resolve()))  # SaveAsFile は絶対パスが必要
                return str(path)
        raise ValueError(f"添付ファイルが見つかりません: {filename}")

    def move_mail(self, mail_id: str, folder: str) -> str:
        mail = self.namespace.GetItemFromID(mail_id)
        moved = mail.Move(self._get_folder(folder))
        # 移動後は EntryID が変わることがあるので、移動後の ID を返す
        return getattr(moved, "EntryID", mail_id)

    def delete(self, mail_id: str) -> None:
        mail = self.namespace.GetItemFromID(mail_id)
        mail.Delete()

    def get_unread_count(self, folder: str) -> int:
        return self._get_folder(folder).UnReadItemCount

    def unread_summary(self, limit: int = 10, folder: str = "inbox") -> list:
        items = self._get_folder(folder).Items
        items.Sort("[ReceivedTime]", True)
        restricted = items.Restrict("[UnRead] = True")
        result = []
        for i, mail in enumerate(restricted):
            if i >= limit:
                break
            try:
                result.append({
                    "subject":   mail.Subject,
                    "from":      mail.SenderEmailAddress,
                    "from_name": mail.SenderName,
                    "date":      mail.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S"),
                    "preview":   mail.Body[:100],
                })
            except Exception:
                continue
        return result

    def sent_today(self, date: str = None) -> list:
        target = datetime.fromisoformat(date).date() if date else datetime.now().date()
        date_str = target.strftime("%m/%d/%Y")
        next_str = (target + timedelta(days=1)).strftime("%m/%d/%Y")
        items = self._get_folder("sent").Items
        restricted = items.Restrict(f"[SentOn] >= '{date_str}' AND [SentOn] < '{next_str}'")
        result = []
        for mail in restricted:
            try:
                result.append({
                    "subject": mail.Subject,
                    "to":      mail.To,
                    "date":    mail.SentOn.strftime("%Y-%m-%dT%H:%M:%S"),
                })
            except Exception:
                continue
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def flagged_or_due(self, days: int = 7, folder: str = "inbox") -> list:
        cutoff = (datetime.now() - timedelta(days=days)).strftime("%m/%d/%Y")
        items = self._get_folder(folder).Items
        # FlagStatus=1 (フラグあり) OR TaskDueDate が設定済み の両方を拾うため日付フィルタのみ Restrict
        restricted = items.Restrict(f"[ReceivedTime] >= '{cutoff}'")
        result = []
        for mail in restricted:
            try:
                flag = getattr(mail, "FlagStatus", 0)
                due = getattr(mail, "TaskDueDate", None)
                due_str = due.strftime("%Y-%m-%d") if due and due.year > 4000 is False and due.year < 4500 else None
                if flag == 1 or due_str:
                    result.append({
                        "id":          mail.EntryID,
                        "subject":     mail.Subject,
                        "from":        mail.SenderEmailAddress,
                        "from_name":   mail.SenderName,
                        "date":        mail.ReceivedTime.strftime("%Y-%m-%dT%H:%M:%S"),
                        "flag_status": flag,
                        "due_date":    due_str,
                    })
            except Exception:
                continue
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def unread_count(self, folder: str = None) -> dict:
        result = {}
        if folder is None:
            for f in ("inbox", "sent", "drafts", "trash"):
                try:
                    count = self.get_unread_count(f)
                except Exception:
                    count = 0
                if count > 0 or f == "inbox":
                    result[f] = count
            for subfolder_name in self.list_subfolders():
                try:
                    count = self.get_unread_count(subfolder_name)
                except Exception:
                    count = 0
                if count > 0:
                    result[subfolder_name] = count
        else:
            try:
                result[folder] = self.get_unread_count(folder)
            except Exception:
                result[folder] = 0
        result["total"] = sum(v for k, v in result.items() if k != "total")
        return result

    def _get_calendar_folder(self):
        return self.namespace.GetDefaultFolder(9)

    def _appointment_to_dict(self, item) -> dict:
        start = item.Start
        end = item.End
        return {
            "id":          item.EntryID,
            "subject":     item.Subject,
            "start":       start.strftime("%Y-%m-%dT%H:%M:%S"),
            "end":         end.strftime("%Y-%m-%dT%H:%M:%S"),
            "location":    item.Location or "",
            "organizer":   item.Organizer or "",
            "attendees":   item.RequiredAttendees or "",
            "all_day":     item.AllDayEvent,
            "busy_status": item.BusyStatus,
        }

    def get_calendar(self, start: datetime, end: datetime) -> list:
        items = self._get_calendar_folder().Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
        start_str = start.strftime("%m/%d/%Y %I:%M %p")
        end_str = end.strftime("%m/%d/%Y %I:%M %p")
        restricted = items.Restrict(f"[Start] >= '{start_str}' AND [Start] < '{end_str}'")
        result = []
        for item in restricted:
            try:
                d = self._appointment_to_dict(item)
                # Python側で再フィルタ（Restrictが漏らすケースへの防御）
                if d["start"] >= start.isoformat()[:19] and d["start"] < end.isoformat()[:19]:
                    result.append(d)
            except Exception:
                continue
        return result

    def get_freebusy(self, email: str, start: datetime, minutes: int = 30) -> str:
        try:
            r = self.namespace.CreateRecipient(email)
            r.Resolve()
            if not r.Resolved:
                return ""
            return r.FreeBusy(start, minutes, True)
        except Exception:
            return ""

    def get_contacts(self) -> list[dict]:
        contacts = []
        try:
            gal = self.namespace.GetGlobalAddressList()
            for entry in gal.AddressEntries:
                try:
                    contacts.append({
                        "name":  entry.Name,
                        "email": entry.GetExchangeUser().PrimarySmtpAddress,
                    })
                except Exception:
                    continue
        except Exception:
            pass
        return contacts
