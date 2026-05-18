"""
Excel client wrapper around xlwings COM access.

Read-only. Hard caps on cell counts. Server-side aggregation via pandas
so we return ANSWERS, not DATA.
"""
from __future__ import annotations

import math
import re
import time
from typing import Any

import pandas as pd
import xlwings as xw

# ---------- Guardrails ----------
MAX_CELLS_PER_CALL = 5_000
MAX_UNIQUE_VALUES = 500
MAX_FIND_ROWS = 200
PEEK_HEAD_ROWS = 3
PEEK_TAIL_ROWS = 3
MAX_USED_CELLS_FOR_AGG = 250_000
CACHE_TTL_SEC = 30


class ExcelError(Exception):
    """Surfaced verbatim to the MCP tool response."""


# ---------- DataFrame cache ----------
# Key: (book.name, sheet.name, used.address, used.shape, has_header)
# Value: (timestamp, df, used_row, used_col)
_DF_CACHE: dict[tuple, tuple[float, pd.DataFrame, int, int]] = {}


def _cache_clear() -> None:
    """Test/manual reset."""
    _DF_CACHE.clear()


# ---------- Helpers ----------
def _safe_fullname(bk: xw.Book) -> str | None:
    """bk.fullname raises XlwingsError for OneDrive-for-Business / SharePoint
    backed workbooks (xlwings can't map the cloud URL to a local path).

    We never actually need the filesystem path; workbooks are identified by
    name. Return None instead of letting the exception kill every tool.
    """
    try:
        return bk.fullname
    except Exception:
        return None


def _get_book(workbook: str) -> xw.Book:
    """Match by exact name, fullname, or unique case-insensitive basename.

    Detects ambiguity across multiple Excel app instances instead of returning
    the first found. Tolerates OneDrive-for-Business workbooks where .fullname
    is unresolvable (matches on name only in that case).
    """
    if not xw.apps:
        raise ExcelError("No Excel instance running. Open Excel first.")
    exact_matches: list[xw.Book] = []
    case_matches: list[xw.Book] = []
    for app in xw.apps:
        for bk in app.books:
            fn = _safe_fullname(bk)
            if bk.name == workbook or (fn is not None and fn == workbook):
                exact_matches.append(bk)
            elif bk.name.lower() == workbook.lower():
                case_matches.append(bk)
    if len(exact_matches) == 1:
        return exact_matches[0]
    if len(exact_matches) > 1:
        raise ExcelError(
            f"Multiple workbooks named {workbook!r} across Excel instances. Use full path."
        )
    if len(case_matches) == 1:
        return case_matches[0]
    if len(case_matches) > 1:
        raise ExcelError(
            f"Multiple workbooks match {workbook!r} (case-insensitive). Use exact name or full path."
        )
    open_names = [bk.name for app in xw.apps for bk in app.books]
    raise ExcelError(
        f"Workbook {workbook!r} not open. Currently open: {open_names}"
    )


def _get_sheet(workbook: str, sheet: str) -> xw.Sheet:
    bk = _get_book(workbook)
    try:
        return bk.sheets[sheet]
    except Exception:
        names = [s.name for s in bk.sheets]
        raise ExcelError(f"Sheet {sheet!r} not in {bk.name!r}. Tabs: {names}")


def _range_cell_count(rng: xw.Range) -> int:
    return rng.shape[0] * rng.shape[1]


def _enforce_cell_cap(rng: xw.Range, op: str) -> None:
    n = _range_cell_count(rng)
    if n > MAX_CELLS_PER_CALL:
        raise ExcelError(
            f"{op}: {n:,} cells exceeds cap of {MAX_CELLS_PER_CALL:,}. "
            f"Narrow the range, or use a query-shaped tool "
            f"(unique_values / count_where / groupby_sum / column_stats)."
        )


def _normalize_range(sheet: xw.Sheet, range_: str) -> xw.Range:
    # Reject open-ended whole-column / whole-row to prevent bloat (incl. absolute $ refs)
    if re.match(r"^\$?[A-Z]+:\$?[A-Z]+$", range_, re.I):
        raise ExcelError(
            f"Whole-column reference {range_!r} not allowed. Use bounded range like A1:A100."
        )
    if re.match(r"^\$?\d+:\$?\d+$", range_):
        raise ExcelError(
            f"Whole-row reference {range_!r} not allowed. Use bounded range like A1:Z1."
        )
    return sheet.range(range_)


def _normalize_2d(values: Any, shape: tuple[int, int]) -> list:
    """Coerce xlwings .value / .formula return into list-of-lists matching shape."""
    rows, cols = shape
    if rows == 1 and cols == 1:
        if isinstance(values, (list, tuple)):
            v = values
            while isinstance(v, (list, tuple)) and v:
                v = v[0]
            return [[v]]
        return [[values]]
    if rows == 1:
        if values is None:
            return [[None] * cols]
        if not isinstance(values, (list, tuple)):
            return [[values]]
        return [list(values)]
    if cols == 1:
        if values is None:
            return [[None] for _ in range(rows)]
        if not isinstance(values, (list, tuple)):
            return [[values]]
        return [[v] for v in values]
    if values is None:
        return [[None] * cols for _ in range(rows)]
    return [list(r) for r in values]


def _read_sheet_cached(sheet: xw.Sheet, has_header: bool = True) -> tuple[pd.DataFrame, int, int]:
    """Pull used range into DataFrame with TTL caching.

    Returns (df, used_row, used_col), both 1-indexed.
    """
    used = sheet.used_range
    addr = used.address
    shape = tuple(used.shape)
    book_name = sheet.book.name
    sheet_name = sheet.name
    key = (book_name, sheet_name, addr, shape, has_header)

    now = time.time()
    cached = _DF_CACHE.get(key)
    if cached is not None:
        ts, df, urow, ucol = cached
        if now - ts < CACHE_TTL_SEC:
            return df.copy(), urow, ucol

    n = shape[0] * shape[1]
    if n > MAX_USED_CELLS_FOR_AGG:
        raise ExcelError(
            f"Sheet used range is {n:,} cells. Too large for server-side aggregation. "
            f"Slice to a sub-range first, or pre-filter in Excel."
        )

    header_arg = 1 if has_header else 0
    df = used.options(pd.DataFrame, header=header_arg, index=False).value
    if df is None:
        df = pd.DataFrame()
    elif isinstance(df, pd.Series):
        df = df.to_frame()

    urow, ucol = int(used.row), int(used.column)
    _DF_CACHE[key] = (now, df, urow, ucol)
    return df.copy(), urow, ucol


def _excel_col_letter_to_num(letter: str) -> int:
    """A=1, Z=26, AA=27, AZ=52, BA=53. 1-indexed Excel column."""
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _resolve_col(df: pd.DataFrame, col: str, used_col: int = 1) -> str:
    """Match by exact name, case-insensitive name, or Excel letter.

    used_col is the 1-indexed Excel column where the DataFrame starts.
    So df.columns[0] corresponds to Excel column `used_col`.
    """
    if col in df.columns:
        return col
    lower = {str(c).lower(): c for c in df.columns}
    if col.lower() in lower:
        return lower[col.lower()]
    # Only treat as Excel column letter if strictly uppercase A-Z, 1-3 chars
    # (Excel max column is XFD = 16384). Avoids parsing header names like 'Cat'.
    if re.match(r"^[A-Z]{1,3}$", col):
        excel_col_num = _excel_col_letter_to_num(col)
        df_col_idx = excel_col_num - used_col
        if 0 <= df_col_idx < len(df.columns):
            return df.columns[df_col_idx]
        raise ExcelError(
            f"Column letter {col!r} (Excel col {excel_col_num}) is outside used range "
            f"(starts at col {used_col}, width {len(df.columns)})."
        )
    raise ExcelError(
        f"Column {col!r} not found. Available: {list(df.columns)[:20]}"
        + (" ..." if len(df.columns) > 20 else "")
    )


def _build_equals_mask(series: pd.Series, value: Any) -> pd.Series:
    """Equality with string + numeric fallback.

    Excel ints come back as floats via xlwings (42 -> 42.0). User passing
    equals=42 needs the numeric path or the comparison silently fails.
    """
    str_mask = series.astype(str) == str(value)
    if isinstance(value, bool):
        return str_mask
    if isinstance(value, (int, float)):
        numeric = pd.to_numeric(series, errors="coerce")
        try:
            num_mask = numeric == float(value)
        except (TypeError, ValueError):
            num_mask = pd.Series(False, index=series.index)
        return str_mask | num_mask
    return str_mask


def _build_contains_mask(series: pd.Series, substring: str) -> pd.Series:
    return series.astype(str).str.contains(re.escape(substring), case=False, na=False)


def _clean_for_json(obj: Any) -> Any:
    """Recursively replace NaN/Inf/pd.NA with None for valid JSON."""
    if isinstance(obj, dict):
        return {k: _clean_for_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_clean_for_json(v) for v in obj]
    if isinstance(obj, float) and (math.isnan(obj) or math.isinf(obj)):
        return None
    if obj is pd.NA or obj is getattr(pd, "NaT", None):
        return None
    return obj


# ---------- Tools ----------
def list_workbooks() -> dict:
    if not xw.apps:
        return {"workbooks": [], "note": "No Excel instance running."}
    out = []
    for app in xw.apps:
        for bk in app.books:
            try:
                active_sheet = bk.sheets.active.name
            except Exception:
                active_sheet = None
            out.append({
                "name": bk.name,
                "fullname": _safe_fullname(bk),
                "sheets": [s.name for s in bk.sheets],
                "active_sheet": active_sheet,
            })
    try:
        active_app = xw.apps.active
        active_book = active_app.books.active.name if active_app else None
    except Exception:
        active_book = None
    return {"workbooks": out, "active_workbook": active_book}


def peek_sheet(workbook: str, sheet: str) -> dict:
    """Headers + shape + first/last rows. Batched into 2 COM calls."""
    ws = _get_sheet(workbook, sheet)
    used = ws.used_range
    rows, cols = used.shape
    if rows == 0:
        return {"workbook": workbook, "sheet": sheet, "rows": 0, "cols": 0,
                "headers": [], "head": [], "tail": []}

    top_n = min(PEEK_HEAD_ROWS + 1, rows)
    top_block_raw = ws.range(
        (used.row, used.column),
        (used.row + top_n - 1, used.column + cols - 1),
    ).value
    top_block = _normalize_2d(top_block_raw, (top_n, cols))
    headers = top_block[0] if top_block else []
    head_rows = top_block[1:] if len(top_block) > 1 else []

    tail_rows: list = []
    if rows > PEEK_HEAD_ROWS + PEEK_TAIL_ROWS + 1:
        tail_block_raw = ws.range(
            (used.row + rows - PEEK_TAIL_ROWS, used.column),
            (used.row + rows - 1, used.column + cols - 1),
        ).value
        tail_rows = _normalize_2d(tail_block_raw, (PEEK_TAIL_ROWS, cols))

    return {
        "workbook": workbook,
        "sheet": sheet,
        "rows": rows,
        "cols": cols,
        "data_rows": rows - 1,
        "anchor": used.address,
        "headers": headers,
        "head": head_rows,
        "tail": tail_rows,
    }


def get_selection() -> dict:
    if not xw.apps:
        raise ExcelError("No Excel instance running. Open Excel first.")
    try:
        active = xw.apps.active
        if active is None:
            raise ExcelError("No active Excel application.")
        sel = active.selection
    except ExcelError:
        raise
    except Exception as e:
        raise ExcelError(f"Could not read selection: {type(e).__name__}: {e}")
    n = _range_cell_count(sel)
    if n > MAX_CELLS_PER_CALL:
        return {
            "workbook": sel.sheet.book.name,
            "sheet": sel.sheet.name,
            "address": sel.address,
            "shape": list(sel.shape),
            "values": None,
            "note": f"Selection is {n:,} cells (cap {MAX_CELLS_PER_CALL:,}). Narrow it.",
        }
    return {
        "workbook": sel.sheet.book.name,
        "sheet": sel.sheet.name,
        "address": sel.address,
        "shape": list(sel.shape),
        "values": _normalize_2d(sel.value, tuple(sel.shape)) if n > 1 else sel.value,
    }


def read_range(workbook: str, sheet: str, range_: str) -> dict:
    ws = _get_sheet(workbook, sheet)
    rng = _normalize_range(ws, range_)
    _enforce_cell_cap(rng, "read_range")
    shape = tuple(rng.shape)
    values = _normalize_2d(rng.value, shape)
    return {
        "workbook": workbook, "sheet": sheet, "range": rng.address,
        "shape": list(shape), "values": values,
    }


def read_formulas(workbook: str, sheet: str, range_: str) -> dict:
    ws = _get_sheet(workbook, sheet)
    rng = _normalize_range(ws, range_)
    _enforce_cell_cap(rng, "read_formulas")
    shape = tuple(rng.shape)
    formulas = _normalize_2d(rng.formula, shape)
    return {
        "workbook": workbook, "sheet": sheet, "range": rng.address,
        "shape": list(shape), "formulas": formulas,
    }


def read_cell(workbook: str, sheet: str, cell: str, formula: bool = False) -> dict:
    ws = _get_sheet(workbook, sheet)
    rng = ws.range(cell)
    if _range_cell_count(rng) != 1:
        raise ExcelError(f"read_cell expects single cell. Got {rng.address!r}.")
    return {
        "workbook": workbook, "sheet": sheet, "cell": rng.address,
        "value": rng.value, "formula": rng.formula if formula else None,
    }


def unique_values(workbook: str, sheet: str, column: str, has_header: bool = True) -> dict:
    df, urow, ucol = _read_sheet_cached(_get_sheet(workbook, sheet), has_header)
    col = _resolve_col(df, column, ucol)
    vals = df[col].dropna().unique().tolist()
    truncated = len(vals) > MAX_UNIQUE_VALUES
    return {
        "workbook": workbook, "sheet": sheet, "column": str(col),
        "count": len(vals),
        "values": vals[:MAX_UNIQUE_VALUES],
        "truncated": truncated,
    }


def count_where(workbook: str, sheet: str, column: str, equals: Any = None,
                contains: str | None = None, not_null: bool = False,
                has_header: bool = True) -> dict:
    predicates_given = sum([equals is not None, contains is not None, bool(not_null)])
    if predicates_given == 0:
        raise ExcelError("count_where: pass one of equals / contains / not_null=True")
    if predicates_given > 1:
        raise ExcelError("count_where: pass only ONE of equals / contains / not_null")

    df, urow, ucol = _read_sheet_cached(_get_sheet(workbook, sheet), has_header)
    col = _resolve_col(df, column, ucol)
    if equals is not None:
        mask = _build_equals_mask(df[col], equals)
    elif contains is not None:
        mask = _build_contains_mask(df[col], contains)
    else:
        mask = df[col].notna()
    return {"column": str(col), "matches": int(mask.sum()), "total_rows": len(df)}


def groupby_sum(workbook: str, sheet: str, group_by: str, sum_col: str,
                has_header: bool = True) -> dict:
    df, urow, ucol = _read_sheet_cached(_get_sheet(workbook, sheet), has_header)
    g = _resolve_col(df, group_by, ucol)
    s = _resolve_col(df, sum_col, ucol)
    df[s] = pd.to_numeric(df[s], errors="coerce")
    out = df.groupby(g, dropna=False)[s].sum().sort_values(ascending=False)
    groups = []
    for k, v in out.items():
        key_is_na = (isinstance(k, float) and math.isnan(k)) or k is pd.NA
        key = None if key_is_na else str(k)
        val_is_na = isinstance(v, float) and math.isnan(v)
        val = None if val_is_na else float(v)
        groups.append({"key": key, "sum": val})
    nonnan_total = out.dropna().sum()
    grand_total = float(nonnan_total) if pd.notna(nonnan_total) else 0.0
    return {
        "group_by": str(g), "sum_col": str(s),
        "groups": groups, "grand_total": grand_total,
    }


def column_stats(workbook: str, sheet: str, column: str, has_header: bool = True) -> dict:
    df, urow, ucol = _read_sheet_cached(_get_sheet(workbook, sheet), has_header)
    col = _resolve_col(df, column, ucol)
    s = df[col]
    numeric = pd.to_numeric(s, errors="coerce")
    non_null_total = int(s.notna().sum())
    numeric_count = int(numeric.notna().sum())
    is_numeric = numeric_count > 0 and numeric_count >= non_null_total * 0.5
    out: dict[str, Any] = {
        "column": str(col),
        "n": int(len(s)),
        "non_null": non_null_total,
        "null": int(s.isna().sum()),
        "is_numeric": bool(is_numeric),
    }
    if is_numeric:
        out["numeric_coverage"] = numeric_count
        if numeric.notna().any():
            out["min"] = float(numeric.min())
            out["max"] = float(numeric.max())
            out["mean"] = float(numeric.mean())
            out["median"] = float(numeric.median())
            out["sum"] = float(numeric.sum())
        else:
            out.update({"min": None, "max": None, "mean": None, "median": None, "sum": 0.0})
    else:
        out["unique"] = int(s.nunique(dropna=True))
        top = s.value_counts(dropna=True).head(5)
        out["top_5"] = [{"value": str(k), "count": int(v)} for k, v in top.items()]
    return out


def find_rows(workbook: str, sheet: str, column: str, equals: Any = None,
              contains: str | None = None, has_header: bool = True) -> dict:
    predicates_given = sum([equals is not None, contains is not None])
    if predicates_given == 0:
        raise ExcelError("find_rows: pass equals OR contains")
    if predicates_given > 1:
        raise ExcelError("find_rows: pass only ONE of equals / contains")

    ws = _get_sheet(workbook, sheet)
    df, urow, ucol = _read_sheet_cached(ws, has_header)
    col = _resolve_col(df, column, ucol)
    if equals is not None:
        mask = _build_equals_mask(df[col], equals)
    else:
        mask = _build_contains_mask(df[col], contains)
    idxs = df.index[mask].tolist()

    header_offset = 1 if has_header else 0
    sheet_rows = [urow + header_offset + i for i in idxs]
    truncated = len(sheet_rows) > MAX_FIND_ROWS
    rows_returned = None
    if not truncated:
        rows_returned = df.loc[idxs].to_dict(orient="records")
    return {
        "column": str(col),
        "match_count": len(sheet_rows),
        "sheet_rows": sheet_rows[:MAX_FIND_ROWS],
        "rows": rows_returned,
        "truncated": truncated,
    }
