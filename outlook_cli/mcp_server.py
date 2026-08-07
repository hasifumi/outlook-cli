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


def main():
    mcp.run(
        transport="streamable-http",
        host=os.getenv("OUTLOOK_MCP_HOST", "0.0.0.0"),
        port=int(os.getenv("OUTLOOK_MCP_PORT", "8764")),
    )


if __name__ == "__main__":
    main()
