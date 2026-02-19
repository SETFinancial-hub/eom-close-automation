"""
QBO MCP Server
--------------
Exposes QuickBooks Online financial data as MCP tools
so Claude can query QBO reports alongside portfolio data.

Usage:
    python src/qbo/mcp_server.py

Then configure in ~/.claude/claude_desktop_config.json:
    {
      "mcpServers": {
        "qbo": {
          "command": "python",
          "args": ["src/qbo/mcp_server.py"],
          "cwd": "/path/to/eom-close-automation"
        }
      }
    }
"""
import json
import sys
import os
from datetime import date, timedelta
from pathlib import Path

# Ensure src/ is on the path when run directly
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from dotenv import load_dotenv
load_dotenv()

from src.qbo.client import QBOClient
from src.qbo.auth import QBOAuth

# MCP protocol uses stdio JSON-RPC
import mcp.server.stdio
import mcp.types as types
from mcp.server import NotificationOptions, Server
from mcp.server.models import InitializationOptions

server = Server("qbo-set-financial")
_client: QBOClient | None = None


def get_client() -> QBOClient:
    global _client
    if _client is None:
        _client = QBOClient()
    return _client


def _last_month_range() -> tuple[str, str]:
    today = date.today()
    first_of_month = today.replace(day=1)
    last_month_end = first_of_month - timedelta(days=1)
    last_month_start = last_month_end.replace(day=1)
    return last_month_start.isoformat(), last_month_end.isoformat()


# ------------------------------------------------------------------
# MCP Tool definitions
# ------------------------------------------------------------------

@server.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="qbo_profit_and_loss",
            description=(
                "Pull SET Financial's QuickBooks Online Profit & Loss report. "
                "Returns income and expense accounts with balances for the period."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "start_date": {"type": "string", "description": "YYYY-MM-DD (default: first of last month)"},
                    "end_date": {"type": "string", "description": "YYYY-MM-DD (default: last day of last month)"},
                },
            },
        ),
        types.Tool(
            name="qbo_balance_sheet",
            description=(
                "Pull SET Financial's QuickBooks Online Balance Sheet. "
                "Returns assets, liabilities, and equity as of a given date."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "as_of_date": {"type": "string", "description": "YYYY-MM-DD (default: last day of last month)"},
                },
            },
        ),
        types.Tool(
            name="qbo_general_ledger",
            description=(
                "Pull SET Financial's QuickBooks Online General Ledger detail. "
                "Use for transaction-level reconciliation against Nortridge data."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "start_date": {"type": "string", "description": "YYYY-MM-DD"},
                    "end_date": {"type": "string", "description": "YYYY-MM-DD"},
                },
                "required": ["start_date", "end_date"],
            },
        ),
        types.Tool(
            name="qbo_chart_of_accounts",
            description="Return SET Financial's full QuickBooks Chart of Accounts with current balances.",
            inputSchema={"type": "object", "properties": {}},
        ),
        types.Tool(
            name="qbo_cash_flow",
            description="Pull SET Financial's QuickBooks Online Cash Flow Statement for a date range.",
            inputSchema={
                "type": "object",
                "properties": {
                    "start_date": {"type": "string", "description": "YYYY-MM-DD"},
                    "end_date": {"type": "string", "description": "YYYY-MM-DD"},
                },
            },
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:
    client = get_client()
    start_default, end_default = _last_month_range()

    try:
        if name == "qbo_profit_and_loss":
            df = client.profit_and_loss(
                start_date=arguments.get("start_date", start_default),
                end_date=arguments.get("end_date", end_default),
            )
            return [types.TextContent(type="text", text=df.to_markdown(index=False))]

        elif name == "qbo_balance_sheet":
            df = client.balance_sheet(
                as_of_date=arguments.get("as_of_date", end_default),
            )
            return [types.TextContent(type="text", text=df.to_markdown(index=False))]

        elif name == "qbo_general_ledger":
            df = client.general_ledger(
                start_date=arguments["start_date"],
                end_date=arguments["end_date"],
            )
            return [types.TextContent(type="text", text=df.to_markdown(index=False))]

        elif name == "qbo_chart_of_accounts":
            df = client.chart_of_accounts()
            return [types.TextContent(type="text", text=df.to_markdown(index=False))]

        elif name == "qbo_cash_flow":
            df = client.cash_flow(
                start_date=arguments.get("start_date", start_default),
                end_date=arguments.get("end_date", end_default),
            )
            return [types.TextContent(type="text", text=df.to_markdown(index=False))]

        else:
            return [types.TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        return [types.TextContent(type="text", text=f"Error calling {name}: {e}")]


# ------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------

async def main():
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name="qbo-set-financial",
                server_version="0.1.0",
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
