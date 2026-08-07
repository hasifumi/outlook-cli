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


def main():
    mcp.run(
        transport="streamable-http",
        host=os.getenv("OUTLOOK_MCP_HOST", "0.0.0.0"),
        port=int(os.getenv("OUTLOOK_MCP_PORT", "8764")),
    )


if __name__ == "__main__":
    main()
