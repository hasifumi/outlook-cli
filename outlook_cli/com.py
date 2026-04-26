import sys
from datetime import datetime, timedelta

if sys.platform != "win32":
    raise ImportError("OutlookCOMはWindowsのみ対応しています")

import win32com.client

from .base import OutlookBase

FOLDER_IDS = {
    "inbox":  6,
    "sent":   5,
    "drafts": 16,
    "trash":  3,
}


class OutlookCOM(OutlookBase):

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

    def _get_folder(self, folder: str):
        if folder in FOLDER_IDS:
            return self.namespace.GetDefaultFolder(FOLDER_IDS[folder])
        inbox = self.namespace.GetDefaultFolder(6)
        for f in inbox.Folders:
            if f.Name == folder:
                return f
        raise ValueError(f"フォルダが見つかりません: {folder}")

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
        }
        if include_body:
            result["body"] = mail.Body
        return result

    def list_mails(self, folder: str = "inbox", limit: int = 20) -> list:
        items = self._get_folder(folder).Items
        items.Sort("[ReceivedTime]", True)
        result = []
        for i, mail in enumerate(items):
            if i >= limit:
                break
            try:
                result.append(self._mail_to_dict(mail))
            except Exception:
                continue
        return result

    def search(self, keyword: str, days: int = 7, sender: str = None) -> list:
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
                    result.append(self._mail_to_dict(mail))
            except Exception:
                continue
        return sorted(result, key=lambda m: m["date"], reverse=True)

    def read(self, mail_id: str) -> dict:
        mail = self.namespace.GetItemFromID(mail_id)
        return self._mail_to_dict(mail, include_body=True)

    def send(self, to: str, subject: str, body: str) -> None:
        mail = self.outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        mail.Send()

    def reply(self, mail_id: str, body: str) -> None:
        original = self.namespace.GetItemFromID(mail_id)
        reply = original.Reply()
        reply.Body = body + "\n\n" + reply.Body
        reply.Send()

    def list_subfolders(self) -> list[str]:
        inbox = self.namespace.GetDefaultFolder(6)
        return [f.Name for f in inbox.Folders]

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
