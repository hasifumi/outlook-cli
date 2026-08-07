import os

from dotenv import load_dotenv
from mcp.server.mcpserver import MCPServer

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


def main():
    mcp.run(
        transport="streamable-http",
        host=os.getenv("OUTLOOK_MCP_HOST", "0.0.0.0"),
        port=int(os.getenv("OUTLOOK_MCP_PORT", "8764")),
    )


if __name__ == "__main__":
    main()
