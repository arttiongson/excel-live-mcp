"""
Excel MCP Server: exposes live Excel session to Claude via stdio MCP.

Read-only by design. Server-side aggregation tools so Claude asks
QUESTIONS, not pulls DATA. See client.py for guardrails.

Run:
    python -m excel_mcp.server
    excel-live-mcp           # after pip install

Claude Code config: see README.md
"""
from __future__ import annotations

import asyncio
import json
import logging
from concurrent.futures import ThreadPoolExecutor
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from . import client as ec

logging.basicConfig(level=logging.INFO, format="%(asctime)s [excel-mcp] %(message)s")
log = logging.getLogger("excel-mcp")

server: Server = Server("excel-live-mcp")

# COM is not thread-safe. Serialize xlwings access through one worker thread
# and gate concurrent tool calls with an asyncio lock.
_COM_LOCK = asyncio.Lock()


def _com_thread_init() -> None:
    try:
        import pythoncom  # type: ignore[import-not-found]
        pythoncom.CoInitialize()
    except Exception:
        pass


_EXECUTOR = ThreadPoolExecutor(
    max_workers=1,
    thread_name_prefix="excel-com",
    initializer=_com_thread_init,
)


# ---------- Tool registry ----------
TOOLS: list[Tool] = [
    Tool(
        name="list_workbooks",
        description=(
            "List all open Excel workbooks with their sheet tabs and active sheet. "
            "Cheap. Always call this first to see what's available."
        ),
        inputSchema={"type": "object", "properties": {}, "additionalProperties": False},
    ),
    Tool(
        name="peek_sheet",
        description=(
            "Headers + row/col count + first 3 + last 3 rows of a sheet. "
            "Constant token cost regardless of sheet size. Use to understand "
            "structure before deciding what to query."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string", "description": "Workbook name (e.g. 'MyWorkbook.xlsx')"},
                "sheet": {"type": "string", "description": "Sheet/tab name"},
            },
            "required": ["workbook", "sheet"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="get_selection",
        description=(
            "Return the cell/range the user has currently selected in Excel. "
            "Use this when the user says 'look at what I have selected' or 'this row'."
        ),
        inputSchema={"type": "object", "properties": {}, "additionalProperties": False},
    ),
    Tool(
        name="read_range",
        description=(
            "Read computed values from a bounded range, returned as 2D list. "
            "Hard cap: 5,000 cells. Whole-column refs (A:A) rejected. "
            "Prefer query-shaped tools (groupby_sum, count_where) for big sheets."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "range": {"type": "string", "description": "A1-notation range, e.g. 'A1:G10'"},
            },
            "required": ["workbook", "sheet", "range"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="read_formulas",
        description=(
            "Read formulas (not computed values) from a bounded range, returned as 2D list. "
            "Use this to learn the calc layer (XLOOKUPs, etc). 5K cell cap."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "range": {"type": "string"},
            },
            "required": ["workbook", "sheet", "range"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="read_cell",
        description="Single cell read. Cheap probe. Optionally returns formula.",
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "cell": {"type": "string", "description": "A1 ref, e.g. 'AB2'"},
                "formula": {"type": "boolean", "default": False},
            },
            "required": ["workbook", "sheet", "cell"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="unique_values",
        description=(
            "Distinct values in a column (server-side). Capped at 500. "
            "Column can be header name (case-insensitive) or Excel letter (A, B, AA...). "
            "Set has_header=false for sheets without a header row."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "column": {"type": "string"},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["workbook", "sheet", "column"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="count_where",
        description=(
            "Count rows where column matches predicate. Pass exactly one of: "
            "equals (exact match, numeric AND string both checked, so 42 matches cell 42.0), "
            "contains (substring, case-insensitive), or not_null=true."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "column": {"type": "string"},
                "equals": {},
                "contains": {"type": "string"},
                "not_null": {"type": "boolean", "default": False},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["workbook", "sheet", "column"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="groupby_sum",
        description=(
            "Pivot: sum(sum_col) grouped by group_by. Server-side pandas. "
            "Null group keys come back as JSON null (not the string 'nan'). "
            "Use this instead of reading whole sheets to compute totals."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "group_by": {"type": "string"},
                "sum_col": {"type": "string"},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["workbook", "sheet", "group_by", "sum_col"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="column_stats",
        description=(
            "min/max/mean/median/null-count for numeric columns, or "
            "unique-count + top-5 for text columns. Auto-detects type."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "column": {"type": "string"},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["workbook", "sheet", "column"],
            "additionalProperties": False,
        },
    ),
    Tool(
        name="find_rows",
        description=(
            "Return SHEET ROW NUMBERS (1-indexed) where column matches predicate. "
            "Pass equals or contains. Includes row data only if matches <= 200. "
            "Correctly handles sheets where the data block starts below row 1."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "workbook": {"type": "string"},
                "sheet": {"type": "string"},
                "column": {"type": "string"},
                "equals": {},
                "contains": {"type": "string"},
                "has_header": {"type": "boolean", "default": True},
            },
            "required": ["workbook", "sheet", "column"],
            "additionalProperties": False,
        },
    ),
]

DISPATCH = {
    "list_workbooks": lambda **_: ec.list_workbooks(),
    "peek_sheet": lambda **kw: ec.peek_sheet(kw["workbook"], kw["sheet"]),
    "get_selection": lambda **_: ec.get_selection(),
    "read_range": lambda **kw: ec.read_range(kw["workbook"], kw["sheet"], kw["range"]),
    "read_formulas": lambda **kw: ec.read_formulas(kw["workbook"], kw["sheet"], kw["range"]),
    "read_cell": lambda **kw: ec.read_cell(kw["workbook"], kw["sheet"], kw["cell"], kw.get("formula", False)),
    "unique_values": lambda **kw: ec.unique_values(
        kw["workbook"], kw["sheet"], kw["column"], has_header=kw.get("has_header", True),
    ),
    "count_where": lambda **kw: ec.count_where(
        kw["workbook"], kw["sheet"], kw["column"],
        equals=kw.get("equals"), contains=kw.get("contains"),
        not_null=kw.get("not_null", False), has_header=kw.get("has_header", True),
    ),
    "groupby_sum": lambda **kw: ec.groupby_sum(
        kw["workbook"], kw["sheet"], kw["group_by"], kw["sum_col"],
        has_header=kw.get("has_header", True),
    ),
    "column_stats": lambda **kw: ec.column_stats(
        kw["workbook"], kw["sheet"], kw["column"], has_header=kw.get("has_header", True),
    ),
    "find_rows": lambda **kw: ec.find_rows(
        kw["workbook"], kw["sheet"], kw["column"],
        equals=kw.get("equals"), contains=kw.get("contains"),
        has_header=kw.get("has_header", True),
    ),
}


# ---------- MCP plumbing ----------
@server.list_tools()
async def _list_tools() -> list[Tool]:
    return TOOLS


@server.call_tool()
async def _call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    log.info(f"call {name} args={list(arguments.keys())}")
    try:
        fn = DISPATCH.get(name)
        if fn is None:
            return [TextContent(type="text", text=f"ERROR: unknown tool {name!r}")]
        async with _COM_LOCK:
            loop = asyncio.get_event_loop()
            result = await loop.run_in_executor(_EXECUTOR, lambda: fn(**arguments))
        cleaned = ec._clean_for_json(result)
        text = json.dumps(cleaned, default=str, allow_nan=False)
        return [TextContent(type="text", text=text)]
    except ec.ExcelError as e:
        return [TextContent(type="text", text=f"ERROR: {e}")]
    except Exception as e:
        log.exception("tool failed")
        return [TextContent(type="text", text=f"ERROR ({type(e).__name__}): {e}")]


async def main() -> None:
    log.info("excel-mcp server starting (stdio)")
    async with stdio_server() as (read, write):
        await server.run(read, write, server.create_initialization_options())


def cli_main() -> None:
    """Console script entry point."""
    asyncio.run(main())


if __name__ == "__main__":
    cli_main()
