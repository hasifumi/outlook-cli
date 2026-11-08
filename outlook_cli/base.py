from abc import ABC, abstractmethod
from datetime import datetime


class OutlookBase(ABC):

    @abstractmethod
    def list_mails(self, folder: str = "inbox", limit: int = 20, include_body: bool = False) -> list:
        """メール一覧取得（include_body=True で本文全文を "body" に含める）"""
        ...

    @abstractmethod
    def search(self, keyword: str, days: int = 7, sender: str = None, include_body: bool = False) -> list:
        """メール検索（include_body=True で本文全文を "body" に含める）"""
        ...

    @abstractmethod
    def read(self, mail_id: str) -> dict:
        """メール本文取得"""
        ...

    @abstractmethod
    def send(self, to: str, subject: str, body: str) -> None:
        """メール送信"""
        ...

    @abstractmethod
    def reply(self, mail_id: str, body: str) -> None:
        """メール返信"""
        ...

    @abstractmethod
    def save_mail_attachment(self, mail_id: str, filename: str, dest_dir: str) -> str:
        """添付ファイルを dest_dir に保存して保存先パスを返す。
        dest_dir は無ければ作成。同名ファイルがあれば上書きせず a_1.xlsx のように連番を付ける。
        添付が無ければ ValueError"""
        ...

    def save_mail_attachment_to_temp(self, mail_id: str, filename: str) -> str:
        """添付ファイルを一時フォルダに保存してパスを返す（開いて確認する用途）"""
        import tempfile
        return self.save_mail_attachment(mail_id, filename, tempfile.mkdtemp(prefix="outlook_cli_"))

    @abstractmethod
    def move_mail(self, mail_id: str, folder: str) -> str:
        """メールを別フォルダへ移動し、移動後のメール ID を返す
        （実 Outlook では移動で EntryID が変わることがある）。メールが無ければ ValueError"""
        ...

    @abstractmethod
    def list_subfolders(self, recursive: bool = False) -> list[str]:
        """サブフォルダ一覧取得（recursive=True で孫フォルダ以下も "親/子" 形式で含める）"""
        ...

    @abstractmethod
    def get_contacts(self) -> list[dict]:
        """連絡先一覧取得"""
        ...

    @abstractmethod
    def delete(self, mail_id: str) -> None:
        """メール削除（ゴミ箱へ移動）"""
        ...

    @abstractmethod
    def get_unread_count(self, folder: str) -> int:
        """フォルダの未読件数取得"""
        ...

    @abstractmethod
    def unread_count(self, folder: str = None) -> dict:
        """未読件数集計（folder=Noneで全フォルダ）→ {"inbox": 2, "total": 2}"""
        ...

    @abstractmethod
    def unread_summary(self, limit: int = 10, folder: str = "inbox") -> list:
        """未読メールのサマリー取得 → [{"subject", "from", "from_name", "date", "preview"}, ...]"""
        ...

    @abstractmethod
    def sent_today(self, date: str = None) -> list:
        """当日の送信メール一覧 (date="YYYY-MM-DD", 省略時は今日) → [{"subject", "to", "date"}, ...]"""
        ...

    @abstractmethod
    def flagged_or_due(self, days: int = 7, folder: str = "inbox") -> list:
        """重要フラグまたは期限設定のあるメール (days以内受信) → [{"id", "subject", "from", "date", "flag_status", "due_date"}, ...]"""
        ...

    @abstractmethod
    def get_calendar(self, start: datetime, end: datetime) -> list:
        """期間内の予定一覧 → [{"id", "subject", "start", "end", "location", "organizer", "attendees", "all_day", "busy_status"}, ...]"""
        ...

    @abstractmethod
    def get_freebusy(self, email: str, start: datetime, minutes: int = 30) -> str:
        """空き状況ビット文字列取得（30分単位のbusy_statusを連結）"""
        ...
