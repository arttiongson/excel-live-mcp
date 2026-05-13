# excel-live-mcp

### Excel Claude Code Token Optimized MCP Server

> A **read-only MCP server for Microsoft Excel** designed to **minimize token usage** through `peek_sheet`, `count_where`, `groupby_sum`, and other server-side aggregation tools. Operates on your **live Excel session** via xlwings COM, no file paths required. Works with **Claude Code**, **Claude Desktop**, and any **Model Context Protocol** client.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![MCP](https://img.shields.io/badge/MCP-1.0+-purple.svg)](https://modelcontextprotocol.io/)
[![Platform: Windows](https://img.shields.io/badge/platform-windows-lightgrey.svg)](#platform)
[![Status: Beta](https://img.shields.io/badge/status-beta-orange.svg)](#)

**Also indexed as:** Excel Claude Code Token Optimized MCP, Token Optimized Excel MCP, Excel MCP server, Claude Excel MCP, Claude Excel integration, MCP server for Excel, xlwings MCP, Anthropic MCP Excel, Claude Desktop Excel MCP, live Excel session MCP, read-only Excel MCP, token efficient Excel MCP.

---

## Why a Token Optimized Excel MCP?

Most existing Excel MCP servers (`negokaz/excel-mcp-server`, `sbroenne/mcp-server-excel`, `mort-lab/excel-mcp`, `haris-musa/excel-mcp-server`, `sbraind/excel-mcp-server`) target feature completeness. They give the LLM 20+ tools that read, write, format, run VBA, and manipulate PivotTables.

That's the wrong tradeoff when **context window cost** is your bottleneck. Returning raw rows blows your token budget on data the model has to re-read on every turn.

**excel-live-mcp is designed to minimize token usage:**

- `peek_sheet` returns headers plus the first 3 and last 3 rows at constant cost, regardless of sheet size
- `count_where`, `groupby_sum`, `column_stats` run server-side and return a single number or small list, not the underlying rows
- `unique_values` caps at 500 distinct values, so a dirty 50K-row column doesn't blow your context
- `read_range` hard-caps at 5,000 cells; whole-column refs like `A:A` are rejected before they bloat context
- `find_rows` returns sheet row numbers, not row data, when matches exceed 200
- TTL DataFrame cache so repeat queries on the same sheet are nearly free
- Read-only by design. No `write_range`, no `set_formula`, no `save`. Zero risk to your data.
- COM concurrency safe. All tool calls serialized through one dedicated worker thread.

If you need to write to Excel, look at `negokaz`, `sbroenne`, or `sbraind`. If you care about not blowing your Claude context window on raw data, use this.

## Platform

- **OS:** Windows 10 / 11 (uses pywin32 COM bindings)
- **Excel:** Microsoft Excel installed and running
- **Python:** 3.10 or newer

Mac support is on the roadmap if there's demand (xlwings supports Mac via AppleScript bridge).

## Install

### Recommended: install from GitHub via pip

```powershell
pip install git+https://github.com/arttiongson/excel-live-mcp.git
```

### Or with uvx (no Python env needed in your project)

```powershell
uvx --from git+https://github.com/arttiongson/excel-live-mcp excel-live-mcp
```

### Local development

```powershell
git clone https://github.com/arttiongson/excel-live-mcp.git
cd excel-live-mcp
pip install -e ".[dev]"
pytest tests/ -v
```

## Wire into Claude Code (or Claude Desktop)

Add to your MCP client's config (`~/.claude/settings.json` for Claude Code, similar path for other clients):

```json
{
  "mcpServers": {
    "excel": {
      "command": "excel-live-mcp",
      "args": [],
      "env": {}
    }
  }
}
```

Restart the client. Tools appear with names like `mcp__excel__list_workbooks`, `mcp__excel__peek_sheet`, etc.

## Tools

11 query-shaped tools, designed so the model asks questions instead of pulling data.

| Tool | Returns | Token cost |
|---|---|---|
| `list_workbooks` | Open workbooks, tabs, active context | tiny |
| `peek_sheet` | Headers, shape, first 3 / last 3 rows | constant (~500 tokens) |
| `get_selection` | Whatever cell or range is highlighted | tiny to small |
| `read_range` | 2D values from bounded range (5K cap) | proportional |
| `read_formulas` | 2D formulas from bounded range | proportional |
| `read_cell` | Single-cell read, optional formula | tiny |
| `unique_values` | Distinct values in a column (cap 500) | small |
| `count_where` | Predicate count (equals / contains / not_null) | tiny |
| `groupby_sum` | Pivot sum by column | small |
| `column_stats` | min/max/mean/median/null counts or top-5 categorical | tiny |
| `find_rows` | Sheet row numbers matching predicate | small |

All aggregation tools accept `has_header` (default `True`) for sheets where row 1 isn't a header.

## Guardrails

- **5,000 cell cap** on `read_range` and `read_formulas`
- Whole-column refs (`A:A`, `$A:$A`) and whole-row refs (`1:1`) rejected
- `unique_values` capped at 500 distinct values
- `find_rows` returns row data only if matches <= 200 (otherwise just row numbers)
- Sheets with > 250,000 cells refuse server-side aggregation. Slice first.
- All COM access serialized through a single worker thread (safe under parallel tool calls)
- 30-second TTL cache on used-range pulls

## Correctness notes

- **Numeric equality** works across types. `equals=42` matches cells containing `42`, `42.0`, or `"42"`. (Excel returns ints as floats via xlwings, so naive string compare would miss these.)
- **Sheet row numbers** from `find_rows` correctly account for used_range offset. If your data starts at row 5, returned row numbers are correct, not off-by-4.
- **Excel column letters** in `column` params are strict uppercase (`A`, `AA`, `XFD`, up to Excel's max). Header names are case-insensitive.
- **Null group keys** in `groupby_sum` come back as JSON `null`, not the literal string `"nan"`.

## FAQ

### What is excel-live-mcp?

A token optimized, read-only Excel MCP server. It's a Python Model Context Protocol (MCP) server that lets Claude (and any MCP-compatible client) read from a live Microsoft Excel session through peek tools, predicate counts, and server-side aggregation, all designed to minimize how many tokens get spent reading spreadsheet data.

### How does it minimize token usage?

Three mechanisms:

1. **Peek tools** like `peek_sheet` return headers plus a handful of head/tail rows at constant token cost, so the model can understand a sheet's structure without reading the whole thing.
2. **Server-side aggregation** in `count_where`, `groupby_sum`, `column_stats`, and `unique_values` runs the computation in Python and returns just the answer (a number or small list), instead of dumping rows for the model to add up.
3. **Hard caps** on raw reads: 5,000 cells per `read_range` call, 500 distinct values in `unique_values`, 200 matches before `find_rows` drops the row payload. Whole-column references like `A:A` are rejected entirely.

### Does it work with Claude Desktop?

Yes. Any MCP-compatible client works: Claude Desktop, Claude Code, Cline, Continue.dev, Cursor, OpenAI agents with MCP adapters, etc.

### Can it write to my workbook?

No. The server is read-only by design. There's no `write_range`, `set_formula`, or `save` tool. If you need write capability, look at `negokaz/excel-mcp-server`, `sbroenne/mcp-server-excel`, or `sbraind/excel-mcp-server`.

### Does it require Excel to be installed?

Yes. The server uses xlwings COM bindings to talk to a running Excel instance. It does NOT read `.xlsx` files from disk directly; the workbook must be open in Excel.

### Does it work on Mac or Linux?

Currently Windows-only. The codebase uses pywin32 for COM. Mac support is feasible via xlwings' AppleScript bridge and is on the roadmap.

### How does the DataFrame cache work?

The first call that triggers a sheet read pulls the used range via COM and caches the resulting pandas DataFrame in memory for 30 seconds, keyed on workbook name, sheet name, used range address, and shape. Subsequent calls within that window skip the COM roundtrip entirely. Cell-value changes within the same range may be served from stale cache for up to 30 seconds; structural changes (rows/cols added) invalidate immediately.

### Why server-side aggregation instead of just pulling the data?

Returning raw rows wastes the LLM's context window. A `groupby_sum` over a 50K-row sheet pulls 50K rows of context if done naively. Doing the pivot server-side returns one number per group. The tool design biases toward "ask a question, get an answer" rather than "pull data, reason over it."

### How is this different from openpyxl-based Excel MCPs?

openpyxl reads `.xlsx` files from disk. It doesn't see live changes, doesn't know which workbook is active, and can't read your current selection. excel-live-mcp uses xlwings + COM, so it operates on the Excel session that's actually open on your machine.

## Troubleshooting

- **"No Excel instance running"**: open Excel before starting the MCP client.
- **"Workbook not open"**: name must match exactly (e.g. `MyWorkbook.xlsx`). Call `list_workbooks` to see open names.
- **COM modal block**: if Excel pops a "Recover unsaved" dialog, the server hangs until dismissed. Click through it in Excel.
- **"Column 'X' not found"**: use exact header name, case-insensitive name, or uppercase Excel letter (`A`, `AA`, `XFD`). Pass `has_header=false` if your sheet has no header row.
- **"Sheet used range is too large"**: sheet exceeds 250K cells. Slice down in Excel or work on a smaller tab.

## Development

```powershell
git clone https://github.com/arttiongson/excel-live-mcp.git
cd excel-live-mcp
pip install -e ".[dev]"
pytest tests/ -v
```

The test suite uses mocks and does not require Excel to be installed or running. 54 tests cover column resolution, numeric equality, JSON cleaning, predicate validation, sheet-row math, and tool-level behavior.

## Roadmap

- Mac support (xlwings AppleScript backend)
- PyPI release (currently install via `pip install git+https://...`)
- MCP registry submission once PyPI is live
- Optional read-write mode with explicit confirmation gating (post-1.0)

## License

MIT. See [LICENSE](LICENSE).

---

**Keywords:** Excel Claude Code Token Optimized MCP, Token Optimized Excel MCP, MCP, Model Context Protocol, MCP server, Excel, Microsoft Excel, xlwings, Claude, Claude Code, Claude Desktop, Anthropic, LLM tools, AI agents, spreadsheet automation, pandas, COM, Windows, token efficiency, token optimization, context efficiency, read-only Excel, live Excel session, Python MCP server, minimize token usage, peek_sheet, server-side aggregation.
