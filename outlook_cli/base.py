from abc import ABC, abstractmethod


class OutlookBase(ABC):

    @abstractmethod
    def list_mails(self, folder: str = "inbox", limit: int = 20) -> list:
        """メール一覧取得"""
        ...

    @abstractmethod
    def search(self, keyword: str, days: int = 7, sender: str = None) -> list:
        """メール検索"""
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
    def list_subfolders(self) -> list[str]:
        """サブフォルダ一覧取得"""
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
