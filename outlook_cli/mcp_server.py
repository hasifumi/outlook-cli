import os
from datetime import datetime, timedelta

from dotenv import load_dotenv
from mcp.server.mcpserver import MCPServer

from .cli import _compute_free_slots
from .mock import OutlookMock

load_dotenv()


def get_client():
    if os.getenv("OUTLOOK_MOCK"):
        return OutlookMock()
    else:
        from .com import OutlookCOM
        return OutlookCOM()


mcp = MCPServer("outlook-cli")


@mcp.tool(name="list")
def list_mails_tool(folder: str = "inbox", limit: int = 20) -> list:
    """メール一覧取得"""
    return get_client().list_mails(folder=folder, limit=limit)


@mcp.tool()
def search(keyword: str, days: int = 7, sender: str | None = None) -> list:
    """メール検索"""
    return get_client().search(keyword=keyword, days=days, sender=sender)


@mcp.tool()
def read(mail_id: str) -> dict:
    """メール本文取得"""
    return get_client().read(mail_id)


@mcp.tool()
def send(to: str, subject: str, body: str) -> dict:
    """メール送信"""
    get_client().send(to=to, subject=subject, body=body)
    return {"status": "sent"}


@mcp.tool()
def reply(mail_id: str, body: str) -> dict:
    """メール返信"""
    get_client().reply(mail_id=mail_id, body=body)
    return {"status": "replied"}


@mcp.tool()
def unread_count(folder: str | None = None) -> dict:
    """フォルダ別未読件数"""
    return get_client().unread_count(folder=folder)


@mcp.tool()
def unread_summary(limit: int = 10, folder: str = "inbox") -> list:
    """未読メール本文冒頭サマリー"""
    return get_client().unread_summary(limit=limit, folder=folder)


@mcp.tool()
def sent_today(date: str | None = None) -> list:
    """当日の送信メール一覧"""
    return get_client().sent_today(date=date)


@mcp.tool()
def flagged(folder: str = "inbox", days: int = 7) -> list:
    """重要フラグまたは期限設定のあるメール"""
    return get_client().flagged_or_due(days=days, folder=folder)


@mcp.tool()
def cal_today(date: str | None = None) -> list:
    """今日の予定一覧"""
    target = datetime.fromisoformat(date).date() if date else datetime.now().date()
    start = datetime(target.year, target.month, target.day)
    end = start + timedelta(days=1)
    return get_client().get_calendar(start, end)


@mcp.tool()
def cal_week(date: str | None = None) -> list:
    """今週（月〜金）の予定一覧"""
    base = datetime.fromisoformat(date).date() if date else datetime.now().date()
    monday = base - timedelta(days=base.weekday())
    start = datetime(monday.year, monday.month, monday.day)
    end = start + timedelta(days=5)
    return get_client().get_calendar(start, end)


@mcp.tool()
def find_slot(
    attendees: str,
    duration: int = 60,
    days: int = 5,
    work_start: int = 9,
    work_end: int = 18,
) -> list:
    """複数人の空き時間候補を検索"""
    client = get_client()
    emails = [e.strip() for e in attendees.split(";") if e.strip()]
    search_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    freebusy_map = {}
    for email in emails:
        fb = client.get_freebusy(email, search_start, 30)
        if fb != "":
            freebusy_map[email] = fb

    if not freebusy_map:
        raise ValueError("空き情報を取得できる参加者がいません")

    return _compute_free_slots(freebusy_map, search_start, days, duration, work_start, work_end)


def main():
    mcp.run(
        transport="streamable-http",
        host=os.getenv("OUTLOOK_MCP_HOST", "0.0.0.0"),
        port=int(os.getenv("OUTLOOK_MCP_PORT", "8764")),
    )


if __name__ == "__main__":
    main()
