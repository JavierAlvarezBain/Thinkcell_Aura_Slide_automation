from __future__ import annotations

import json
import logging
import re
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

from shared import BuildConfig, clean_text, require_file, stage, to_int_or_none, to_number_or_none


DATA_MAX_ROWS = 1600
DATA_MAX_COLS = 180


@dataclass
class CellProxy:
    value: Any


@dataclass
class ValueSheet:
    title: str
    rows: list[list[Any]]
    max_row: int
    max_column: int

    def value(self, row: int, column: int) -> Any:
        if row < 1 or column < 1:
            return None
        row_idx = row - 1
        col_idx = column - 1
        if row_idx >= len(self.rows):
            return None
        row_values = self.rows[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]

    def cell(self, row: int, column: int) -> CellProxy:
        return CellProxy(self.value(row, column))


@dataclass
class ValueWorkbook:
    sheets: dict[str, ValueSheet]
    sheetnames: list[str]

    def __getitem__(self, name: str) -> ValueSheet:
        return self.sheets[name]

    def close(self) -> None:
        return None


@dataclass
class SourceLayout:
    sheet_name: str
    header_row: int
    function_col: int
    company_start_col: int
    company_end_col: int
    quartile_start_col: int
    quartile_end_col: int
    revenue_row: int
    revenue_label_col: int
    fte_start_row: int
    fte_end_row: int
    l0_col: int | None = None
    quartile_columns: dict[str, int] = field(default_factory=dict)
    fte_section_row: int | None = None
    confidence: float = 0.0
    # The first table on Aura Summary is the canonical company/revenue source.
    # Its columns can differ from the later FTE/$B table, so keep both layouts.
    company_header_row: int | None = None
    company_header_start_col: int | None = None
    company_header_end_col: int | None = None


def load_workbook_fast(path):
    try:
        return load_workbook(path, data_only=True, read_only=True, keep_links=False)
    except TypeError:
        return load_workbook(path, data_only=True, read_only=True)


def _mapped_workbook_bounds(
    payload: dict[str, Any],
) -> dict[str, tuple[int, int]]:
    """Return the maximum row/column needed on each mapped worksheet.

    The mapped-input reader only needs explicitly selected rectangles plus any
    function-label range stored in legacy metadata. Reading from A1 through the
    furthest required cell lets read-only openpyxl stream each relevant sheet
    once while preserving O(1) random access afterward via ``ValueSheet``.
    """
    bounds: dict[str, tuple[int, int]] = {}

    def include(sheet_name: str, range_a1: str) -> None:
        sheet = clean_text(sheet_name)
        cell_range = clean_text(range_a1)
        if not sheet or not cell_range:
            return
        try:
            _min_col, _min_row, max_col, max_row = range_boundaries(cell_range)
        except Exception:
            return
        old_row, old_col = bounds.get(sheet, (0, 0))
        bounds[sheet] = (max(old_row, int(max_row)), max(old_col, int(max_col)))

    inputs = payload.get("inputs") or {}
    if isinstance(inputs, dict):
        for mapping in inputs.values():
            if isinstance(mapping, dict):
                include(mapping.get("sheet_name"), mapping.get("range_a1"))
                include(
                    mapping.get("row_label_sheet") or mapping.get("sheet_name"),
                    mapping.get("row_label_range"),
                )

    function_lookup = payload.get("function_lookup") or {}
    if isinstance(function_lookup, dict):
        fte_mapping = inputs.get("fte_values") if isinstance(inputs, dict) else {}
        fte_sheet = (fte_mapping or {}).get("sheet_name") if isinstance(fte_mapping, dict) else ""
        include(
            function_lookup.get("function_label_sheet") or fte_sheet,
            function_lookup.get("function_label_range"),
        )

    return bounds


def load_workbook_for_mapped_ranges(
    path,
    payload: dict[str, Any],
    logger: logging.Logger | None = None,
) -> ValueWorkbook:
    """Stream only the worksheet area required by explicit Aura mappings.

    A mapped build already knows the exact cells it needs. Fully materializing a
    large workbook (``read_only=False``) can take minutes because openpyxl parses
    every worksheet, style, formula cell, and relationship before returning.
    Instead we open in read-only/data-only mode, stream each referenced sheet
    once from A1 through its furthest mapped cell, and retain that small value
    window in memory. Existing mapped-range validation then runs unchanged on
    ``ValueSheet`` objects with constant-time ``.cell(...)`` access.
    """
    bounds = _mapped_workbook_bounds(payload)
    if not bounds:
        raise ValueError("Aura input map contains no readable worksheet ranges.")

    if logger is not None:
        logger.info(
            "Opening mapped workbook in streaming values mode; only %s referenced worksheet(s) will be materialized.",
            len(bounds),
        )

    raw_workbook = load_workbook_fast(path)
    available_sheetnames = list(raw_workbook.sheetnames)
    sheets: dict[str, ValueSheet] = {}
    try:
        for sheet_name, (max_row, max_col) in bounds.items():
            if sheet_name not in raw_workbook.sheetnames:
                raise ValueError(
                    f"Mapped input refers to missing worksheet {sheet_name!r}. "
                    f"Available sheets: {raw_workbook.sheetnames}"
                )
            if logger is not None:
                logger.info(
                    "Reading mapped source window: %s!A1:%s%s",
                    sheet_name,
                    col_letter(max_col),
                    max_row,
                )
            sheets[sheet_name] = materialize_values_sheet(
                raw_workbook[sheet_name],
                max_rows=max_row,
                max_cols=max_col,
            )
    finally:
        try:
            raw_workbook.close()
        except Exception:
            pass

    return ValueWorkbook(sheets=sheets, sheetnames=available_sheetnames)

def materialize_values_sheet(
    ws,
    max_rows: int = DATA_MAX_ROWS,
    max_cols: int = DATA_MAX_COLS,
) -> ValueSheet:
    rows: list[list[Any]] = []
    actual_max_row = 0
    actual_max_col = 0

    for row_idx, row in enumerate(
        ws.iter_rows(
            min_row=1,
            max_row=max_rows,
            min_col=1,
            max_col=max_cols,
            values_only=True,
        ),
        start=1,
    ):
        row_values = list(row)
        rows.append(row_values)

        if any(value is not None and clean_text(value) != "" for value in row_values):
            actual_max_row = row_idx
            for col_idx in range(len(row_values), 0, -1):
                value = row_values[col_idx - 1]
                if value is not None and clean_text(value) != "":
                    actual_max_col = max(actual_max_col, col_idx)
                    break

    return ValueSheet(
        title=ws.title,
        rows=rows,
        max_row=max(actual_max_row, len(rows)),
        max_column=max(actual_max_col, min(max_cols, len(rows[0]) if rows else max_cols)),
    )


def normalized_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", clean_text(value).lower()).strip()


def compact_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def is_blank(value: Any) -> bool:
    return clean_text(value) == ""


_MISSING_NUMERIC_TOKENS = {
    "", "-", "na", "n/a", "n\\a", "nd", "n/d", "n.d", "n.d.",
    "nan", "none", "<no data>", "not disclosed", "not available",
}
_EXCEL_ERROR_TOKENS = {
    "#div/0!", "#n/a", "#value!", "#ref!", "#name?", "#num!", "#null!",
}


def _is_missing_numeric_token(value: Any, *, include_blank: bool = True) -> bool:
    text = clean_text(value).casefold()
    if not include_blank and not text:
        return False
    return text in _MISSING_NUMERIC_TOKENS or text in _EXCEL_ERROR_TOKENS


def is_numeric_like(value: Any) -> bool:
    return to_number_or_none(value) is not None


def col_letter(col_num: int) -> str:
    letters = ""
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def looks_like_function_value(value: Any) -> bool:
    text = clean_text(value)
    if not text or text.startswith(">>"):
        return False
    if _is_missing_numeric_token(text, include_blank=False):
        return False
    if to_number_or_none(text) is not None:
        return False
    return bool(re.search(r"[A-Za-z]", text))



def _is_combined_hierarchy_header(value: Any) -> bool:
    compact = compact_text(value)
    return compact in {
        "l1l0", "l0l1", "level1level0", "level0level1",
        "l1andl0", "l0andl1", "level1andlevel0", "level0andlevel1",
    }


def _contiguous_sum_match(
    ws: ValueSheet,
    target_row: int,
    source_first_row: int,
    source_last_row: int,
    company_start_col: int,
    company_end_col: int,
) -> bool:
    company_count = company_end_col - company_start_col + 1
    required_comparisons = max(3, min(6, max(1, company_count // 3)))
    compared = 0
    matched = 0
    for col in range(company_start_col, company_end_col + 1):
        target = to_number_or_none(ws.value(target_row, col))
        if target is None:
            continue
        source_values = [
            to_number_or_none(ws.value(source_row, col))
            for source_row in range(source_first_row, source_last_row + 1)
        ]
        if any(value is None for value in source_values):
            continue
        compared += 1
        source_sum = sum(float(value) for value in source_values if value is not None)
        tolerance = max(1e-6, 1e-6 * max(1.0, abs(float(target)), abs(source_sum)))
        if abs(float(target) - source_sum) <= tolerance:
            matched += 1
    return compared >= required_comparisons and matched / compared >= 0.90


def _find_contiguous_aggregate_source(
    ws: ValueSheet,
    target_row: int,
    detail_first_row: int,
    detail_last_row: int,
    company_start_col: int,
    company_end_col: int,
) -> tuple[int, int] | None:
    for source_first in range(detail_first_row, detail_last_row + 1):
        for source_last in range(source_first, detail_last_row + 1):
            if _contiguous_sum_match(
                ws,
                target_row,
                source_first,
                source_last,
                company_start_col,
                company_end_col,
            ):
                return source_first, source_last
    return None


def _detect_embedded_aggregate_tail(
    ws: ValueSheet,
    header_row: int,
    function_col: int,
    company_start_col: int,
    company_end_col: int,
    candidate_end_row: int,
) -> tuple[int, int, int] | None:
    """Split a combined L1/L0 table from structural aggregation evidence."""
    if not _is_combined_hierarchy_header(ws.value(header_row, function_col)):
        return None

    detail_first = header_row + 1
    if candidate_end_row - detail_first + 1 < 4:
        return None

    for aggregate_start in range(detail_first + 2, candidate_end_row):
        detail_last = aggregate_start - 1
        matches: list[tuple[int, int]] = []
        valid = True
        for target_row in range(aggregate_start, candidate_end_row + 1):
            source = _find_contiguous_aggregate_source(
                ws,
                target_row,
                detail_first,
                detail_last,
                company_start_col,
                company_end_col,
            )
            if source is None:
                valid = False
                break
            matches.append(source)
        if not valid:
            continue
        if not any(source_last > source_first for source_first, source_last in matches):
            continue
        return detail_last, aggregate_start, candidate_end_row
    return None

def contiguous_nonblank_runs(
    ws: ValueSheet,
    row: int,
    start_col: int,
    end_col: int,
) -> list[tuple[int, int]]:
    runs: list[tuple[int, int]] = []
    run_start = None

    for col in range(start_col, end_col + 1):
        if not is_blank(ws.value(row, col)):
            if run_start is None:
                run_start = col
        else:
            if run_start is not None:
                runs.append((run_start, col - 1))
                run_start = None

    if run_start is not None:
        runs.append((run_start, end_col))

    return runs


RAW_OR_NON_SUMMARY_SHEET_TERMS = {
    "rds",
    "raw",
    "data",
    "received",
    "pivot",
    "dump",
    "export",
    "input",
    "mapping",
    "lookup",
    "calc",
    "calculation",
    "working",
}


def aura_summary_sheet_name_score(sheet_name: str) -> int:
    text = normalized_text(sheet_name)
    compact = compact_text(sheet_name)
    tokens = set(text.split())

    if not text:
        return -10_000

    raw_penalty = -500 if tokens & RAW_OR_NON_SUMMARY_SHEET_TERMS else 0

    if compact == "aurasummary":
        return 1200 + raw_penalty
    if compact == "summaryaura":
        return 1150 + raw_penalty
    if compact == "summary":
        return 1100 + raw_penalty
    if "summary" in tokens and "aura" in tokens:
        return 950 + raw_penalty
    if "summary" in tokens:
        return 850 + raw_penalty

    return -10_000


def find_matching_sheet_name(sheetnames: list[str], requested: str) -> str | None:
    requested_clean = clean_text(requested)
    requested_compact = compact_text(requested_clean)

    for sheet_name in sheetnames:
        if sheet_name == requested_clean or sheet_name.lower() == requested_clean.lower():
            return sheet_name

    for sheet_name in sheetnames:
        if compact_text(sheet_name) == requested_compact:
            return sheet_name

    return None


def candidate_aura_summary_sheet_names(sheetnames: list[str]) -> list[tuple[int, str]]:
    candidates = [(aura_summary_sheet_name_score(name), name) for name in sheetnames]
    candidates = [(score, name) for score, name in candidates if score > 0]
    candidates.sort(key=lambda item: (item[0], item[1].lower()), reverse=True)
    return candidates


def validate_sheet_has_table(
    wb,
    sheet_name: str,
    config: BuildConfig,
    logger: logging.Logger,
) -> ValueSheet | None:
    ws = materialize_values_sheet(wb[sheet_name])
    try:
        detect_fte_table_layout(ws, config, logger)
        return ws
    except Exception as exc:
        logger.debug("Rejected Aura Summary candidate %s: %s", sheet_name, exc)
        return None


def choose_aura_summary_sheet(wb, config: BuildConfig, logger: logging.Logger) -> ValueSheet:
    requested = clean_text(getattr(config, "sheet_name", None))

    tried: set[str] = set()
    rejected: list[str] = []

    if requested:
        match = find_matching_sheet_name(wb.sheetnames, requested)
        if match is not None:
            tried.add(match)
            logger.info("Checking requested Aura Summary sheet: %s", match)
            ws = validate_sheet_has_table(wb, match, config, logger)
            if ws is not None:
                logger.info("Using requested Aura Summary sheet: %s", match)
                return ws
            rejected.append(f"{match}: requested sheet did not contain a detectable FTE/$B quartile table")
        else:
            logger.warning(
                "Requested sheet %r was not found. Available sheets: %s",
                requested,
                ", ".join(wb.sheetnames),
            )

    candidates = candidate_aura_summary_sheet_names(wb.sheetnames)
    if not candidates:
        raise ValueError(
            "Could not find an Aura Summary tab by name. Expected a sheet named like "
            "Summary, Aura Summary, aura_summary, Summary_Aura, etc. "
            f"Available sheets: {', '.join(wb.sheetnames)}"
        )

    for score, sheet_name in candidates:
        if sheet_name in tried:
            continue

        logger.info("Checking Aura Summary candidate sheet: %s (name_score=%s)", sheet_name, score)
        tried.add(sheet_name)

        ws = validate_sheet_has_table(wb, sheet_name, config, logger)
        if ws is not None:
            logger.info("Detected Aura Summary sheet: %s", sheet_name)
            return ws

        rejected.append(f"{sheet_name}: no detectable FTE/$B quartile table")

    raise ValueError(
        "Found Aura Summary-like sheet name(s), but none contained a detectable FTE/$B quartile table. "
        f"Rejected candidates: {'; '.join(rejected)}"
    )


def select_summary_sheet(wb, config: BuildConfig, logger: logging.Logger) -> ValueSheet:
    return choose_aura_summary_sheet(wb, config, logger)


NEGATIVE_REVENUE_TERMS = {
    "source",
    "comment",
    "coverage",
    "employee",
    "employees",
    "fte",
    "headcount",
    "margin",
    "growth",
    "gap",
    "percent",
    "percentage",
    "cagr",
    "multiple",
    "ebitda",
    "opex",
}


def revenue_label_score(value: Any) -> int:
    """Score a first-table revenue label without confusing it with Sales L1 rows.

    Valid examples include Revenue, Revenue ($), Revenue ($B), Revenue ($B)
    Latest Year, LTM Revenue, Total Revenue, and Net Sales.  A bare ``Sales``
    function row is deliberately not treated as a revenue label.
    """
    raw = clean_text(value)
    text = normalized_text(raw)
    compact = compact_text(raw)
    tokens = set(text.split())

    if not text:
        return 0
    if tokens & NEGATIVE_REVENUE_TERMS:
        return 0

    has_revenue_word = bool({"revenue", "revenues", "rev"} & tokens)
    has_qualified_sales = "sales" in tokens and bool(
        {"net", "total", "latest", "ltm", "ttm", "b", "bn", "billion", "m", "mm", "million", "usd"}
        & tokens
        or "$" in raw
    )
    if not (has_revenue_word or has_qualified_sales):
        return 0

    score = 40
    if "ltm" in tokens or "ttm" in tokens or "l12m" in compact:
        score += 90
    if (
        "last" in tokens
        and ("twelve" in tokens or "12" in tokens)
        and ("months" in tokens or "month" in tokens)
    ):
        score += 90
    if "total" in tokens:
        score += 40
    if "net" in tokens and "sales" in tokens:
        score += 30
    if "latest" in tokens:
        score += 30

    if "$" in raw or "usd" in tokens:
        score += 15
    has_billions = bool({"b", "bn", "billion", "billions"} & tokens)
    has_millions = bool({"m", "mm", "million", "millions"} & tokens)
    if has_billions:
        score += 25
    if has_millions:
        score += 15

    # A row such as ``Total Revenue Year`` often contains years rather than
    # revenue values. Keep it as a fallback, but rank a unit-bearing/latest row
    # above it.
    if "year" in tokens and not ("latest" in tokens or "$" in raw or has_billions or has_millions):
        score -= 35
    return max(score, 0)


def _contains_fte_token(value: Any) -> bool:
    raw = clean_text(value)
    text = normalized_text(raw)
    compact = compact_text(raw)
    tokens = set(text.split())
    if {"fte", "ftes"} & tokens:
        return True
    if "fulltimeequivalent" in compact or "fulltimeequivalents" in compact:
        return True
    return compact.startswith("fteper") or compact.startswith("ftesper")


def fte_per_revenue_score(value: Any) -> int:
    """Score the section marker above the actual FTE/$B table.

    Token-based FTE detection avoids the historical false positive where the
    word ``after`` happened to contain the letters ``fte``.
    """
    raw = clean_text(value)
    text = normalized_text(raw)
    compact = compact_text(raw)
    tokens = set(text.split())

    if not _contains_fte_token(raw):
        return 0
    if "%" in raw or "percent" in tokens or "percentage" in tokens or "tfte" in compact:
        return 0

    score = 35
    if "revenue" in tokens or "revenues" in tokens or "rev" in tokens or "sales" in tokens:
        score += 45
    if "$" in raw or "usd" in tokens:
        score += 20
    if "per" in tokens or "/" in raw:
        score += 30
    if "1b" in compact or "billion" in tokens or "b" in tokens or "bn" in tokens:
        score += 50
    if "avg" in tokens or "average" in tokens:
        score += 20
    if raw.lstrip().startswith(">>"):
        score += 10
    return score

def quartile_kind(value: Any) -> str | None:
    text = normalized_text(value)
    compact = compact_text(value)
    tokens = set(text.split())

    if not text:
        return None

    if compact in {
        "median",
        "med",
        "middleq",
        "midq",
        "middlequartile",
        "midquartile",
        "q2",
        "p50",
        "50th",
        "50thpercentile",
    }:
        return "median"
    if "median" in tokens:
        return "median"
    if ("middle" in tokens or "mid" in tokens) and ("q" in tokens or "quartile" in tokens):
        return "median"

    if compact in {
        "top",
        "topq",
        "topquartile",
        "topquart",
        "upperq",
        "upperquartile",
        "q1",
        "p25",
        "25th",
        "25thpercentile",
        "firstquartile",
        "1stquartile",
    }:
        return "top"
    if "top" in tokens and ("q" in tokens or "quartile" in tokens):
        return "top"
    if "upper" in tokens and ("q" in tokens or "quartile" in tokens):
        return "top"

    if compact in {
        "bottom",
        "bottomq",
        "bottomquartile",
        "bottomquart",
        "lowerq",
        "lowerquartile",
        "q3",
        "p75",
        "75th",
        "75thpercentile",
        "thirdquartile",
        "3rdquartile",
    }:
        return "bottom"
    if "bottom" in tokens and ("q" in tokens or "quartile" in tokens):
        return "bottom"
    if "lower" in tokens and ("q" in tokens or "quartile" in tokens):
        return "bottom"

    return None


def is_l0_header(value: Any) -> bool:
    compact = compact_text(value)
    return compact in {
        "l0",
        "level0",
        "levelzero",
        "l0level",
        "l0function",
        "l0category",
        "categoryl0",
    }


def is_l1_header(value: Any) -> bool:
    compact = compact_text(value)
    text = normalized_text(value)
    return compact in {
        "l1",
        "level1",
        "levelone",
        "l1level",
        "l1function",
        "function",
        "functions",
        "functionname",
        "category",
        "workstream",
        "department",
        "area",
    } or text in {"level 1", "l 1"}


def function_header_score(value: Any) -> int:
    if is_l1_header(value):
        return 45
    if is_l0_header(value):
        return 5
    compact = compact_text(value)
    if compact.startswith("l1") or "function" in compact or "category" in compact:
        return 25
    return 0


def l0_display_text(value: Any) -> str:
    raw = clean_text(value)
    compact = compact_text(raw)
    if not raw:
        return ""

    mapping = {
        "ga": "G&A",
        "ganda": "G&A",
        "generalandadministrative": "G&A",
        "generaladministrative": "G&A",
        "generaladmin": "G&A",
        "corporate": "G&A",
        "sm": "S&M",
        "salesmarketing": "S&M",
        "salesandmarketing": "S&M",
        "commercial": "S&M",
        "ops": "Ops",
        "operations": "Ops",
        "rd": "R&D",
        "randd": "R&D",
        "researchdevelopment": "R&D",
        "researchanddevelopment": "R&D",
    }
    return mapping.get(compact, raw)


KNOWN_L0_VALUE_COMPACTS = {
    "ga",
    "ganda",
    "generalandadministrative",
    "generaladministrative",
    "generaladmin",
    "corporate",
    "sm",
    "salesmarketing",
    "salesandmarketing",
    "commercial",
    "ops",
    "operations",
    "rd",
    "randd",
    "researchdevelopment",
    "researchanddevelopment",
}


def is_l0_value(value: Any) -> bool:
    return compact_text(value) in KNOWN_L0_VALUE_COMPACTS


def is_l0_text_candidate(value: Any) -> bool:
    text = clean_text(value)
    if not text:
        return False
    if is_numeric_like(value):
        return False
    if len(text) > 60:
        return False
    return bool(re.search(r"[A-Za-z&]", text))


def detect_l0_value_column(
    ws: ValueSheet,
    header_row: int,
    fte_start_row: int,
    fte_end_row: int,
    function_col: int,
    company_start_col: int,
) -> int | None:
    """
    Detect L0 by values, not by header.

    Works when the L0 header is blank/merged and rows contain:
      G&A
      S&M
      Ops
      R&D
    """
    if function_col <= 1:
        return None

    scored: list[tuple[int, int, dict[str, int]]] = []

    for col in range(1, function_col):
        known_hits = 0
        text_hits = 0
        numeric_hits = 0
        blank_hits = 0
        sampled_rows = 0

        for row in range(fte_start_row, fte_end_row + 1):
            function_name = ws.value(row, function_col)
            if not looks_like_function_value(function_name):
                continue

            sampled_rows += 1
            value = ws.value(row, col)

            if is_blank(value):
                blank_hits += 1
                continue
            if is_numeric_like(value):
                numeric_hits += 1
                continue
            if is_l0_value(value):
                known_hits += 1
            elif is_l0_text_candidate(value):
                text_hits += 1

        if sampled_rows == 0:
            continue

        distance_from_l1 = function_col - col
        proximity_bonus = max(0, 30 - distance_from_l1 * 3)

        score = (
            known_hits * 120
            + text_hits * 8
            + proximity_bonus
            - numeric_hits * 80
        )

        if known_hits >= 1 and score > 0:
            scored.append(
                (
                    score,
                    col,
                    {
                        "known_hits": known_hits,
                        "text_hits": text_hits,
                        "numeric_hits": numeric_hits,
                        "blank_hits": blank_hits,
                        "sampled_rows": sampled_rows,
                    },
                )
            )

    if not scored:
        return None

    scored.sort(key=lambda item: item[0], reverse=True)
    return scored[0][1]


def l0_from_row_or_above(
    ws: ValueSheet,
    row: int,
    l0_col: int | None,
    header_row: int,
    max_lookback: int = 60,
) -> str:
    if not l0_col:
        return ""

    start_row = max(header_row + 1, row - max_lookback)

    for scan_row in range(row, start_row - 1, -1):
        value = ws.value(scan_row, l0_col)
        if is_l0_text_candidate(value):
            return l0_display_text(value)

    return ""


def infer_l0_from_l1(function_name: Any) -> str:
    compact = compact_text(function_name)

    if compact in {
        "ehs",
        "environmenthealthsafety",
        "environmentalhealthsafety",
        "facilitiesmanagement",
        "facilities",
        "finance",
        "hr",
        "humanresources",
        "it",
        "informationtechnology",
        "leadership",
        "legal",
        "otherga",
        "otherganda",
        "procurement",
        "supplychainmanagement",
        "admin",
        "administration",
        "corporate",
    }:
        return "G&A"

    if compact in {
        "sales",
        "marketing",
        "merchandising",
        "commercialexcellence",
        "commercial",
        "commercialstrategy",
        "commercialoperations",
        "customerservice",
        "customerexperience",
        "customeroperations",
        "categorymanagement",
        "pricing",
        "revenuegrowthmanagement",
        "salesoperations",
        "salesops",
        "ecommerce",
        "digitalcommerce",
        "businessdevelopment",
    }:
        return "S&M"

    if compact in {
        "operations",
        "ops",
        "manufacturing",
        "logisticssupplychain",
        "logisticsandsupplychain",
        "supplychain",
        "maintenance",
        "ehsops",
        "subcontractors",
        "otherops",
        "warehousing",
        "distribution",
        "fulfillment",
    }:
        return "Ops"

    if compact in {
        "rd",
        "randd",
        "researchdevelopment",
        "researchanddevelopment",
        "cybersecurity",
        "softwaredevelopment",
        "engineering",
        "productdevelopment",
    }:
        return "R&D"

    return ""


def quartile_columns_for_row(ws: ValueSheet, row: int) -> dict[str, int]:
    found: dict[str, int] = {}

    for col in range(1, ws.max_column + 1):
        kind = quartile_kind(ws.value(row, col))
        if kind and kind not in found:
            found[kind] = col

    return found


def count_numeric_values(ws: ValueSheet, row: int, start_col: int, end_col: int) -> int:
    return sum(
        1
        for col in range(start_col, end_col + 1)
        if is_numeric_like(ws.value(row, col))
    )


def count_valid_data_rows(
    ws: ValueSheet,
    header_row: int,
    function_col: int,
    company_start_col: int,
    company_end_col: int,
    quartile_cols: list[int],
    sample_rows: int = 36,
) -> int:
    company_count = company_end_col - company_start_col + 1
    min_company_numbers = max(2, min(4, company_count // 2 if company_count else 2))
    min_quartile_numbers = min(2, len(quartile_cols))
    valid_rows = 0

    for row in range(header_row + 1, min(ws.max_row, header_row + sample_rows) + 1):
        if not looks_like_function_value(ws.value(row, function_col)):
            continue

        company_numbers = count_numeric_values(ws, row, company_start_col, company_end_col)
        quartile_numbers = sum(1 for col in quartile_cols if is_numeric_like(ws.value(row, col)))

        if company_numbers >= min_company_numbers and quartile_numbers >= min_quartile_numbers:
            valid_rows += 1

    return valid_rows


def find_fte_section_row(ws: ValueSheet, header_row: int) -> tuple[int | None, int]:
    best_row = None
    best_score = 0

    for row in range(max(1, header_row - 16), header_row + 1):
        row_score = 0
        for col in range(1, ws.max_column + 1):
            row_score = max(row_score, fte_per_revenue_score(ws.value(row, col)))
        if row_score > best_score:
            best_row = row
            best_score = row_score

    return best_row, best_score



def _blank_or_dash(value: Any) -> bool:
    return clean_text(value).lower() in {"", "-", "na", "n/a"}


def _revenue_value_span(
    ws: ValueSheet,
    row: int,
    label_col: int,
) -> tuple[int, int] | None:
    """Return the contiguous revenue-value span immediately right of a label."""
    start_col: int | None = None
    for col in range(label_col + 1, min(ws.max_column, label_col + 4) + 1):
        value = ws.value(row, col)
        if is_numeric_like(value) or _is_missing_numeric_token(value, include_blank=False):
            start_col = col
            break
    if start_col is None:
        return None

    end_col = start_col - 1
    blank_streak = 0
    for col in range(start_col, ws.max_column + 1):
        value = ws.value(row, col)
        text = clean_text(value)
        if is_numeric_like(value) or _is_missing_numeric_token(value, include_blank=False):
            end_col = col
            blank_streak = 0
            continue
        if not text:
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        break

    if end_col - start_col + 1 < 3:
        return None
    return start_col, end_col


def _company_header_row_score(
    ws: ValueSheet,
    row: int,
    start_col: int,
    end_col: int,
) -> int | None:
    values = [clean_text(ws.value(row, col)) for col in range(start_col, end_col + 1)]
    nonblank = [value for value in values if value]
    expected_count = end_col - start_col + 1
    if len(nonblank) < max(3, int(expected_count * 0.80)):
        return None
    if any(is_numeric_like(value) for value in nonblank):
        return None

    keys = [compact_text(value) for value in nonblank if compact_text(value)]
    unique_count = len(set(keys))
    unique_ratio = unique_count / max(1, len(keys))
    if unique_ratio < 0.60:
        return None

    generic_keys = {
        "public", "private", "annualreport", "companywebsite", "rds",
        "source", "reported", "client", "peer", "ips", "amc", "pes",
    }
    generic_hits = sum(1 for key in keys if key in generic_keys)
    if generic_hits > len(keys) * 0.25:
        return None

    preceding_key = compact_text(ws.value(row, start_col - 1)) if start_col > 1 else ""
    explicit_label_bonus = 80 if preceding_key in {
        "company", "companies", "companyname", "companynames",
        "peercompanies", "peers",
    } else 0

    return (
        100
        + len(nonblank) * 5
        + unique_count * 4
        + int(unique_ratio * 40)
        + explicit_label_bonus
    )


def detect_primary_company_revenue_table(ws: ValueSheet) -> dict[str, Any]:
    """Detect the canonical company row and revenue row in the first Aura table.

    The first table is authoritative for company order.  Later tables repeat
    the same companies, but may append helper columns such as Client Actuals.
    """
    marker_rows = [
        row
        for row in range(1, ws.max_row + 1)
        if max(
            (fte_per_revenue_score(ws.value(row, col)) for col in range(1, ws.max_column + 1)),
            default=0,
        ) >= 80
    ]
    search_end_row = min(marker_rows) - 1 if marker_rows else ws.max_row
    candidates: list[dict[str, Any]] = []

    for revenue_row in range(1, search_end_row + 1):
        for label_col in range(1, ws.max_column + 1):
            label = ws.value(revenue_row, label_col)
            label_score = revenue_label_score(label)
            if label_score <= 0:
                continue
            span = _revenue_value_span(ws, revenue_row, label_col)
            if span is None:
                continue
            company_start_col, company_end_col = span
            numeric_values = [
                to_number_or_none(ws.value(revenue_row, col))
                for col in range(company_start_col, company_end_col + 1)
            ]
            numeric_values = [value for value in numeric_values if value is not None]
            numeric_count = len(numeric_values)
            if numeric_count < 3:
                continue
            year_like_count = sum(
                1
                for value in numeric_values
                if 1900 <= float(value) <= 2100 and abs(float(value) - round(float(value))) < 1e-9
            )
            label_tokens = set(normalized_text(label).split())
            label_has_units = bool(
                "$" in clean_text(label)
                or {"b", "bn", "billion", "billions", "m", "mm", "million", "millions", "usd"}
                & label_tokens
            )
            if (
                year_like_count / numeric_count >= 0.60
                and "year" in label_tokens
                and not label_has_units
            ):
                continue

            for company_row in range(max(1, revenue_row - 30), revenue_row):
                row_score = _company_header_row_score(
                    ws,
                    company_row,
                    company_start_col,
                    company_end_col,
                )
                if row_score is None:
                    continue
                distance = revenue_row - company_row
                score = (
                    label_score
                    + numeric_count * 12
                    + row_score
                    + max(0, 35 - distance)
                    - revenue_row * 0.05
                )
                company_headers = [
                    clean_text(ws.value(company_row, col))
                    for col in range(company_start_col, company_end_col + 1)
                ]
                candidates.append(
                    {
                        "score": score,
                        "company_header_row": company_row,
                        "company_start_col": company_start_col,
                        "company_end_col": company_end_col,
                        "company_headers": company_headers,
                        "revenue_row": revenue_row,
                        "revenue_label_col": label_col,
                        "revenue_label": clean_text(label),
                    }
                )

    if not candidates:
        raise ValueError(
            "Could not detect the first Aura company/revenue table. Expected a unique "
            "company-name row above a labeled Revenue row with aligned numeric values."
        )
    candidates.sort(
        key=lambda item: (float(item["score"]), -int(item["revenue_row"])),
        reverse=True,
    )
    return candidates[0]


def _find_company_sequence_in_header(
    ws: ValueSheet,
    row: int,
    company_headers: list[str],
    first_col: int,
    last_col: int,
) -> tuple[int, int] | None:
    expected = [compact_text(value) for value in company_headers]
    if not expected or any(not key for key in expected):
        return None
    width = len(expected)
    if last_col - first_col + 1 < width:
        return None

    for start_col in range(first_col, last_col - width + 2):
        observed = [
            compact_text(ws.value(row, col))
            for col in range(start_col, start_col + width)
        ]
        if observed == expected:
            return start_col, start_col + width - 1
    return None


def candidate_function_and_company_ranges(
    ws: ValueSheet,
    header_row: int,
    first_quartile_col: int,
) -> list[tuple[int | None, int, int, int, int]]:
    candidates: list[tuple[int | None, int, int, int, int]] = []
    runs = contiguous_nonblank_runs(ws, header_row, 1, first_quartile_col - 1)

    for run_start, run_end in runs:
        run_len = run_end - run_start + 1
        if run_len < 3:
            continue

        if run_len >= 5 and is_l0_header(ws.value(header_row, run_start)) and is_l1_header(ws.value(header_row, run_start + 1)):
            candidates.append((run_start, run_start + 1, run_start + 2, run_end, 180 + run_len))
            continue

        if run_len >= 4 and is_l1_header(ws.value(header_row, run_start)):
            l0_col = run_start - 1 if run_start > 1 and is_l0_header(ws.value(header_row, run_start - 1)) else None
            candidates.append((l0_col, run_start, run_start + 1, run_end, 125 + run_len))
            continue

        if run_start > 1:
            function_col = run_start - 1
            l0_col = function_col - 1 if function_col > 1 and is_l0_header(ws.value(header_row, function_col - 1)) else None
            score = 45 + run_len
            if is_l1_header(ws.value(header_row, function_col)):
                score += 55
            elif function_header_score(ws.value(header_row, function_col)) > 0:
                score += 25
            candidates.append((l0_col, function_col, run_start, run_end, score))

        if run_len >= 4 and not is_l0_header(ws.value(header_row, run_start)):
            candidates.append((None, run_start, run_start + 1, run_end, 40 + run_len))

    for function_col in range(1, first_quartile_col):
        if not is_l1_header(ws.value(header_row, function_col)):
            continue
        l0_col = function_col - 1 if function_col > 1 and is_l0_header(ws.value(header_row, function_col - 1)) else None
        possible_runs = contiguous_nonblank_runs(
            ws,
            header_row,
            function_col + 1,
            first_quartile_col - 1,
        )
        for company_start, company_end in possible_runs:
            if company_end - company_start + 1 >= 3:
                candidates.append(
                    (
                        l0_col,
                        function_col,
                        company_start,
                        company_end,
                        110 + company_end - company_start + 1,
                    )
                )

    best_by_key: dict[tuple[int | None, int, int, int], int] = {}
    for l0_col, function_col, company_start, company_end, score in candidates:
        if company_start <= function_col or company_end < company_start:
            continue
        if is_l0_header(ws.value(header_row, function_col)) and not is_l1_header(ws.value(header_row, function_col)):
            continue
        key = (l0_col, function_col, company_start, company_end)
        best_by_key[key] = max(best_by_key.get(key, -1), score)

    return [(a, b, c, d, score) for (a, b, c, d), score in best_by_key.items()]


def detect_fte_table_end(
    ws: ValueSheet,
    header_row: int,
    function_col: int,
    company_start_col: int,
    company_end_col: int,
    quartile_cols: list[int],
) -> int:
    end_row = header_row
    blank_streak = 0

    for row in range(header_row + 1, ws.max_row + 1):
        function_name = clean_text(ws.value(row, function_col))
        company_numbers = count_numeric_values(ws, row, company_start_col, company_end_col)
        quartile_numbers = sum(1 for col in quartile_cols if is_numeric_like(ws.value(row, col)))

        if not function_name:
            blank_streak += 1
            if end_row > header_row and blank_streak >= 1:
                break
            continue

        blank_streak = 0

        function_key = compact_text(function_name)
        if function_key == "total" or function_key.startswith("totalexcluding"):
            break
        if function_name.startswith(">>") and end_row > header_row:
            break
        if company_numbers == 0 and quartile_numbers == 0 and end_row > header_row:
            break

        if looks_like_function_value(function_name):
            end_row = row

    combined_split = _detect_embedded_aggregate_tail(
        ws,
        header_row,
        function_col,
        company_start_col,
        company_end_col,
        end_row,
    )
    if combined_split is not None:
        end_row = combined_split[0]

    return end_row


def _detect_fte_table_layout_legacy(
    ws: ValueSheet,
    config: BuildConfig,
    logger: logging.Logger,
) -> SourceLayout:
    candidates: list[dict[str, Any]] = []

    for row in range(1, ws.max_row + 1):
        quartile_by_kind = quartile_columns_for_row(ws, row)
        required = {"top", "median", "bottom"}

        if not required.issubset(quartile_by_kind):
            continue

        quartile_cols = [
            quartile_by_kind["top"],
            quartile_by_kind["median"],
            quartile_by_kind["bottom"],
        ]
        first_quartile_col = min(quartile_cols)

        for l0_col, function_col, company_start_col, company_end_col, structural_score in candidate_function_and_company_ranges(
            ws,
            row,
            first_quartile_col,
        ):
            company_count = company_end_col - company_start_col + 1

            if company_count < 3:
                continue

            data_rows = count_valid_data_rows(
                ws,
                header_row=row,
                function_col=function_col,
                company_start_col=company_start_col,
                company_end_col=company_end_col,
                quartile_cols=quartile_cols,
            )

            if data_rows < 2:
                continue

            fte_section_row, fte_section_score = find_fte_section_row(ws, row)

            if fte_section_score < 40 and data_rows < 5:
                continue

            header_score = function_header_score(ws.value(row, function_col))
            l0_score = 35 if l0_col is not None else 0
            quartile_score = 120
            company_score = min(company_count, 20) * 4
            data_score = min(data_rows, 30) * 10
            total_score = (
                structural_score
                + header_score
                + l0_score
                + quartile_score
                + company_score
                + data_score
                + fte_section_score
            )

            candidates.append(
                {
                    "score": total_score,
                    "header_row": row,
                    "l0_col": l0_col,
                    "function_col": function_col,
                    "company_start_col": company_start_col,
                    "company_end_col": company_end_col,
                    "quartile_by_kind": dict(quartile_by_kind),
                    "quartile_cols": quartile_cols,
                    "fte_section_row": fte_section_row,
                    "fte_section_score": fte_section_score,
                    "data_rows": data_rows,
                    "company_count": company_count,
                }
            )

    if not candidates:
        raise ValueError(
            "Could not detect the FTE/$B quartile table on the Aura Summary tab. "
            "Expected a table with company headers plus quartile headers like Top Q/Top Quartile, "
            "Median/Middle Q, and Bottom Q/Bottom Quartile."
        )

    candidates.sort(key=lambda item: item["score"], reverse=True)
    best = candidates[0]

    quartile_by_kind = best["quartile_by_kind"]
    quartile_cols = [
        quartile_by_kind["top"],
        quartile_by_kind["median"],
        quartile_by_kind["bottom"],
    ]

    revenue_row, revenue_label_col = find_revenue_row(
        ws,
        int(best["company_start_col"]),
        int(best["company_end_col"]),
    )

    fte_end_row = detect_fte_table_end(
        ws,
        int(best["header_row"]),
        int(best["function_col"]),
        int(best["company_start_col"]),
        int(best["company_end_col"]),
        quartile_cols,
    )

    detected_l0_col = int(best["l0_col"]) if best.get("l0_col") is not None else None

    value_detected_l0_col = detect_l0_value_column(
        ws=ws,
        header_row=int(best["header_row"]),
        fte_start_row=int(best["header_row"]) + 1,
        fte_end_row=fte_end_row,
        function_col=int(best["function_col"]),
        company_start_col=int(best["company_start_col"]),
    )

    if value_detected_l0_col is not None:
        if detected_l0_col != value_detected_l0_col:
            logger.info(
                "Detected L0 value column by row values: %s%s",
                col_letter(value_detected_l0_col),
                int(best["header_row"]),
            )
        detected_l0_col = value_detected_l0_col

    layout = SourceLayout(
        sheet_name=ws.title,
        header_row=int(best["header_row"]),
        l0_col=detected_l0_col,
        function_col=int(best["function_col"]),
        company_start_col=int(best["company_start_col"]),
        company_end_col=int(best["company_end_col"]),
        quartile_start_col=min(quartile_cols),
        quartile_end_col=max(quartile_cols),
        revenue_row=revenue_row,
        revenue_label_col=revenue_label_col,
        fte_start_row=int(best["header_row"]) + 1,
        fte_end_row=fte_end_row,
        quartile_columns={
            "top": quartile_by_kind["top"],
            "median": quartile_by_kind["median"],
            "bottom": quartile_by_kind["bottom"],
        },
        fte_section_row=best["fte_section_row"],
        confidence=float(best["score"]),
    )

    logger.debug("Best FTE/$B table candidate: %s", best)
    logger.debug("Top FTE/$B table candidates: %s", candidates[:8])

    return layout



def detect_fte_table_layout(
    ws: ValueSheet,
    config: BuildConfig,
    logger: logging.Logger,
) -> SourceLayout:
    """Detect the canonical first table and the marker-anchored FTE/$B table.

    The first table supplies company order and Revenue.  The later table is
    accepted only when it sits beneath an FTE/$B section label, repeats that
    company sequence, and exposes semantic Top/Median/Bottom headers in any
    left-to-right order.
    """
    try:
        primary = detect_primary_company_revenue_table(ws)
    except Exception as exc:
        logger.warning(
            "Canonical Aura company/revenue table detection failed; using the "
            "legacy structural detector. Reason: %s",
            exc,
        )
        return _detect_fte_table_layout_legacy(ws, config, logger)

    canonical_headers = list(primary["company_headers"])
    canonical_count = len(canonical_headers)
    candidates: list[dict[str, Any]] = []

    for header_row in range(1, ws.max_row + 1):
        quartile_by_kind = quartile_columns_for_row(ws, header_row)
        if not {"top", "median", "bottom"}.issubset(quartile_by_kind):
            continue
        quartile_cols = [
            quartile_by_kind["top"],
            quartile_by_kind["median"],
            quartile_by_kind["bottom"],
        ]
        first_quartile_col = min(quartile_cols)

        company_sequence = _find_company_sequence_in_header(
            ws,
            header_row,
            canonical_headers,
            1,
            first_quartile_col - 1,
        )
        if company_sequence is None:
            continue
        company_start_col, company_end_col = company_sequence
        if company_end_col - company_start_col + 1 != canonical_count:
            continue

        function_col = company_start_col - 1
        if function_col < 1 or function_header_score(ws.value(header_row, function_col)) <= 0:
            continue
        l0_col = (
            function_col - 1
            if function_col > 1 and is_l0_header(ws.value(header_row, function_col - 1))
            else None
        )

        fte_section_row, fte_section_score = find_fte_section_row(ws, header_row)
        if fte_section_row is None or fte_section_score < 80:
            continue

        valid_rows = count_valid_data_rows(
            ws,
            header_row=header_row,
            function_col=function_col,
            company_start_col=company_start_col,
            company_end_col=company_end_col,
            quartile_cols=quartile_cols,
            sample_rows=60,
        )
        if valid_rows < 2:
            continue

        score = (
            400
            + fte_section_score
            + function_header_score(ws.value(header_row, function_col))
            + (35 if l0_col is not None else 0)
            + min(canonical_count, 30) * 5
            + min(valid_rows, 40) * 10
            + max(0, 20 - (header_row - fte_section_row))
        )
        candidates.append(
            {
                "score": score,
                "header_row": header_row,
                "l0_col": l0_col,
                "function_col": function_col,
                "company_start_col": company_start_col,
                "company_end_col": company_end_col,
                "quartile_by_kind": dict(quartile_by_kind),
                "quartile_cols": quartile_cols,
                "fte_section_row": fte_section_row,
                "fte_section_score": fte_section_score,
                "valid_rows": valid_rows,
            }
        )

    if not candidates:
        logger.warning(
            "No FTE/$B table matched the canonical first-table company sequence "
            "beneath an FTE/$B marker; using the legacy structural detector."
        )
        legacy = _detect_fte_table_layout_legacy(ws, config, logger)
        if legacy.company_end_col - legacy.company_start_col + 1 == canonical_count:
            legacy.company_header_row = int(primary["company_header_row"])
            legacy.company_header_start_col = int(primary["company_start_col"])
            legacy.company_header_end_col = int(primary["company_end_col"])
            legacy.revenue_row = int(primary["revenue_row"])
            legacy.revenue_label_col = int(primary["revenue_label_col"])
        return legacy

    candidates.sort(key=lambda item: float(item["score"]), reverse=True)
    best = candidates[0]
    quartile_by_kind = dict(best["quartile_by_kind"])
    quartile_cols = [
        quartile_by_kind["top"],
        quartile_by_kind["median"],
        quartile_by_kind["bottom"],
    ]
    fte_end_row = detect_fte_table_end(
        ws,
        int(best["header_row"]),
        int(best["function_col"]),
        int(best["company_start_col"]),
        int(best["company_end_col"]),
        quartile_cols,
    )
    if fte_end_row <= int(best["header_row"]):
        raise ValueError("The marker-anchored FTE/$B table contains no function rows.")

    detected_l0_col = int(best["l0_col"]) if best.get("l0_col") is not None else None
    value_detected_l0_col = detect_l0_value_column(
        ws=ws,
        header_row=int(best["header_row"]),
        fte_start_row=int(best["header_row"]) + 1,
        fte_end_row=fte_end_row,
        function_col=int(best["function_col"]),
        company_start_col=int(best["company_start_col"]),
    )
    if value_detected_l0_col is not None:
        detected_l0_col = value_detected_l0_col

    layout = SourceLayout(
        sheet_name=ws.title,
        header_row=int(best["header_row"]),
        l0_col=detected_l0_col,
        function_col=int(best["function_col"]),
        company_start_col=int(best["company_start_col"]),
        company_end_col=int(best["company_end_col"]),
        quartile_start_col=min(quartile_cols),
        quartile_end_col=max(quartile_cols),
        revenue_row=int(primary["revenue_row"]),
        revenue_label_col=int(primary["revenue_label_col"]),
        fte_start_row=int(best["header_row"]) + 1,
        fte_end_row=fte_end_row,
        quartile_columns={
            "top": quartile_by_kind["top"],
            "median": quartile_by_kind["median"],
            "bottom": quartile_by_kind["bottom"],
        },
        fte_section_row=int(best["fte_section_row"]),
        confidence=float(best["score"]),
        company_header_row=int(primary["company_header_row"]),
        company_header_start_col=int(primary["company_start_col"]),
        company_header_end_col=int(primary["company_end_col"]),
    )
    logger.debug("Canonical first Aura table: %s", primary)
    logger.debug("Best marker-anchored FTE/$B candidate: %s", best)
    return layout

def find_revenue_row(ws: ValueSheet, company_start_col: int, company_end_col: int) -> tuple[int, int]:
    candidates: list[tuple[int, int, int, str]] = []

    for row in range(1, ws.max_row + 1):
        for col in range(1, min(ws.max_column, company_start_col) + 1):
            label = ws.value(row, col)
            base_score = revenue_label_score(label)

            if base_score <= 0:
                continue

            numeric_count = 0
            blank_or_dash_count = 0

            for value_col in range(company_start_col, company_end_col + 1):
                raw_value = ws.value(row, value_col)

                if is_numeric_like(raw_value):
                    numeric_count += 1
                elif _is_missing_numeric_token(raw_value):
                    blank_or_dash_count += 1

            if numeric_count == 0:
                continue

            numeric_values = [
                to_number_or_none(ws.value(row, value_col))
                for value_col in range(company_start_col, company_end_col + 1)
            ]
            numeric_values = [value for value in numeric_values if value is not None]
            year_like_count = sum(
                1
                for value in numeric_values
                if 1900 <= float(value) <= 2100 and abs(float(value) - round(float(value))) < 1e-9
            )
            label_tokens = set(normalized_text(label).split())
            label_has_units = bool(
                "$" in clean_text(label)
                or {"b", "bn", "billion", "billions", "m", "mm", "million", "millions", "usd"}
                & label_tokens
            )
            if (
                numeric_values
                and year_like_count / len(numeric_values) >= 0.60
                and "year" in label_tokens
                and not label_has_units
            ):
                continue

            score = base_score + numeric_count * 14 + blank_or_dash_count
            candidates.append((score, row, col, clean_text(label)))

    if not candidates:
        raise ValueError(
            "Could not find a Revenue row on the Aura Summary tab aligned to the detected company columns. "
            "Accepted labels include LTM Revenue, Last Twelve Months Revenue, Total Revenue(s), "
            "Revenue ($B), Revenue ($M), Rev, Sales, or Net Sales."
        )

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, row, col, _ = candidates[0]

    return row, col


AURA_INPUT_MAP_SCHEMA_VERSION = 9
AURA_SHARED_INPUT_KEYS = (
    "taxonomy",
    "company_headers",
    "revenue_values",
    "fte_values",
)
AURA_L0_QUARTILE_INPUT_KEYS = ("l0_quartile_table",)
AURA_L1_BENCHMARK_INPUT_KEYS = ("l1_quartile_table",)
AURA_INPUT_KEYS = (
    AURA_SHARED_INPUT_KEYS
    + AURA_L0_QUARTILE_INPUT_KEYS
    + AURA_L1_BENCHMARK_INPUT_KEYS
)

# Legacy schema 3-8 stored three separate L1 benchmark columns.
LEGACY_L1_BENCHMARK_INPUT_KEYS = (
    "top_quartile",
    "median",
    "bottom_quartile",
)
LEGACY_AURA_INPUT_KEYS = AURA_SHARED_INPUT_KEYS + LEGACY_L1_BENCHMARK_INPUT_KEYS

# Legacy aliases retained for older manifests and imports.
L1_INPUT_MAP_SCHEMA_VERSION = AURA_INPUT_MAP_SCHEMA_VERSION
L1_REQUIRED_INPUT_KEYS = AURA_INPUT_KEYS

_TAXONOMY_L0_HEADER_KEYS = {
    "l0",
    "level0",
    "levelzero",
    "l0name",
    "l0category",
    "level0name",
}
_TAXONOMY_L1_HEADER_KEYS = {
    "l1",
    "level1",
    "levelone",
    "l1name",
    "l1function",
    "level1name",
    "function",
}


def _load_l1_input_map(path: Path) -> dict[str, Any]:
    """Load either the legacy L1 map or the standardized Aura input map."""
    require_file(path, "Aura input map")
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Aura input map is not readable JSON: {path}") from exc

    schema_version = int(payload.get("schema_version", -1))
    supported_versions = {3, 4, 5, 6, 7, 8, AURA_INPUT_MAP_SCHEMA_VERSION}
    if schema_version not in supported_versions:
        raise ValueError(
            "Unsupported Aura input-map schema version: "
            f"{payload.get('schema_version')!r}. Expected one of "
            f"{sorted(supported_versions)}."
        )

    slide_family = clean_text(payload.get("slide_family")).lower()
    if slide_family not in {"l1", "aura"}:
        raise ValueError(
            "The supplied input map is not an Aura mapping manifest."
        )

    selected_sections = {
        clean_text(value).lower()
        for value in (payload.get("selected_sections") or [])
        if clean_text(value)
    }
    requires_l1_benchmarks = bool(
        payload.get("requires_l1_benchmarks")
        if "requires_l1_benchmarks" in payload
        else (slide_family == "l1" or "l1" in selected_sections)
    )
    requires_l0_quartiles = bool(
        payload.get("requires_l0_quartiles")
        if "requires_l0_quartiles" in payload
        else ("l0" in selected_sections)
    )
    if schema_version >= 9:
        required_list = list(AURA_SHARED_INPUT_KEYS)
        if requires_l0_quartiles:
            required_list.extend(AURA_L0_QUARTILE_INPUT_KEYS)
        if requires_l1_benchmarks:
            required_list.extend(AURA_L1_BENCHMARK_INPUT_KEYS)
        required_keys = tuple(required_list)
    else:
        required_keys = (
            LEGACY_AURA_INPUT_KEYS
            if requires_l1_benchmarks
            else AURA_SHARED_INPUT_KEYS
        )

    inputs = payload.get("inputs")
    if not isinstance(inputs, dict):
        raise ValueError("Aura input map has no inputs object.")
    missing = [
        key for key in required_keys if not isinstance(inputs.get(key), dict)
    ]
    if missing:
        raise ValueError(
            "Aura input map is missing required input(s): " + ", ".join(missing)
        )

    payload["requires_l1_benchmarks"] = requires_l1_benchmarks
    payload["requires_l0_quartiles"] = requires_l0_quartiles
    payload["required_input_keys"] = list(required_keys)
    return payload


def _mapped_range_matrix(
    workbook,
    mapping: dict[str, Any],
    input_label: str,
) -> list[list[Any]]:
    sheet_name = clean_text(mapping.get("sheet_name"))
    range_a1 = clean_text(mapping.get("range_a1"))
    if not sheet_name or not range_a1:
        raise ValueError(f"{input_label} mapping must contain a sheet name and A1 range.")
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"{input_label} refers to missing worksheet {sheet_name!r}. "
            f"Available sheets: {workbook.sheetnames}"
        )
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    except Exception as exc:
        raise ValueError(f"{input_label} has an invalid A1 range: {range_a1!r}.") from exc
    worksheet = workbook[sheet_name]
    return [
        [worksheet.cell(row=row, column=column).value for column in range(min_col, max_col + 1)]
        for row in range(min_row, max_row + 1)
    ]


def _mapped_vector(matrix: list[list[Any]], input_label: str) -> list[Any]:
    row_count = len(matrix)
    column_count = max((len(row) for row in matrix), default=0)
    if row_count == 1 and column_count >= 1:
        return list(matrix[0])
    if column_count == 1 and row_count >= 1:
        return [row[0] if row else None for row in matrix]
    raise ValueError(
        f"{input_label} must be a single row or a single column; "
        f"received {row_count} x {column_count}."
    )


def _mapped_plain_row(
    matrix: list[list[Any]],
    input_label: str,
) -> list[Any]:
    """Read a horizontal value-only row; the first selected cell is data."""
    if len(matrix) != 1 or len(matrix[0]) < 1:
        rows = len(matrix)
        cols = max((len(row) for row in matrix), default=0)
        raise ValueError(
            f"{input_label} must be one horizontal value-only row; "
            f"received {rows} x {cols}."
        )
    return list(matrix[0])


def _mapped_labeled_row(
    matrix: list[list[Any]],
    input_label: str,
) -> tuple[str, list[Any]]:
    """Read one horizontal range whose first cell is a required master label."""
    if len(matrix) != 1 or len(matrix[0]) < 2:
        rows = len(matrix)
        cols = max((len(row) for row in matrix), default=0)
        raise ValueError(
            f"{input_label} must be one horizontal row containing a master label "
            f"followed by data; received {rows} x {cols}."
        )
    label = clean_text(matrix[0][0])
    if not label:
        raise ValueError(f"{input_label} is missing its first-cell master label.")
    return label, list(matrix[0][1:])


def _mapped_labeled_column(
    matrix: list[list[Any]],
    input_label: str,
) -> tuple[str, list[Any]]:
    """Read one vertical range whose first cell is a required master label."""
    if len(matrix) < 2 or any(len(row) != 1 for row in matrix):
        rows = len(matrix)
        cols = max((len(row) for row in matrix), default=0)
        raise ValueError(
            f"{input_label} must be one vertical column containing its header "
            f"followed by values; received {rows} x {cols}."
        )
    label = clean_text(matrix[0][0])
    if not label:
        raise ValueError(f"{input_label} is missing its top master label/header.")
    return label, [row[0] for row in matrix[1:]]


def _mapped_quartile_table(
    workbook,
    mapping: dict[str, Any],
    input_label: str,
) -> dict[str, Any]:
    """Read one three-column Top/Median/Bottom table in any column order.

    The selected range may include the header row. For lower L0 benchmark
    tables whose headers are shared with an earlier table, the nearest matching
    header in the same three columns is resolved above the selected range.
    """
    sheet_name = clean_text(mapping.get("sheet_name"))
    range_a1 = clean_text(mapping.get("range_a1"))
    if not sheet_name or not range_a1:
        raise ValueError(f"{input_label} mapping must contain a sheet name and A1 range.")
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"{input_label} refers to missing worksheet {sheet_name!r}. "
            f"Available sheets: {workbook.sheetnames}"
        )
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    except Exception as exc:
        raise ValueError(f"{input_label} has an invalid A1 range: {range_a1!r}.") from exc
    if max_col - min_col + 1 != 3:
        raise ValueError(
            f"{input_label} must contain exactly three consecutive columns; "
            f"received {max_col - min_col + 1}."
        )

    worksheet = workbook[sheet_name]

    def header_map(row_number: int) -> dict[str, int]:
        found: dict[str, int] = {}
        for column in range(min_col, max_col + 1):
            kind = quartile_kind(worksheet.cell(row=row_number, column=column).value)
            if kind and kind not in found:
                found[kind] = column
        return found

    headers = header_map(min_row)
    header_in_selection = {"top", "median", "bottom"}.issubset(headers)
    if header_in_selection:
        header_row = min_row
        data_start_row = min_row + 1
    else:
        header_row = 0
        data_start_row = min_row
        for row_number in range(min_row - 1, max(0, min_row - 80), -1):
            candidate = header_map(row_number)
            if {"top", "median", "bottom"}.issubset(candidate):
                header_row = row_number
                headers = candidate
                break
    if not header_row:
        raise ValueError(
            f"{input_label} could not find Top, Median, and Bottom headers in "
            "the selected columns or immediately above them."
        )

    data_rows = list(range(data_start_row, max_row + 1))
    if not data_rows:
        raise ValueError(f"{input_label} contains no benchmark data rows.")

    raw_by_kind: dict[str, list[Any]] = {"top": [], "median": [], "bottom": []}
    for row_number in data_rows:
        for kind in ("top", "median", "bottom"):
            raw_by_kind[kind].append(
                worksheet.cell(row=row_number, column=int(headers[kind])).value
            )

    return {
        "sheet_name": sheet_name,
        "range_a1": range_a1,
        "min_col": min_col,
        "max_col": max_col,
        "min_row": min_row,
        "max_row": max_row,
        "header_row": header_row,
        "header_in_selection": header_in_selection,
        "headers": {
            kind: {
                "column": int(headers[kind]),
                "label": clean_text(worksheet.cell(row=header_row, column=int(headers[kind])).value),
            }
            for kind in ("top", "median", "bottom")
        },
        "data_rows": data_rows,
        "data_row_count": len(data_rows),
        "raw_top": raw_by_kind["top"],
        "raw_median": raw_by_kind["median"],
        "raw_bottom": raw_by_kind["bottom"],
    }


def _validate_company_master_label(label: str) -> None:
    descriptor = _taxonomy_header_descriptor(label)
    if (descriptor is not None and descriptor[0] == 1) or is_l1_header(label) or function_header_score(label) > 0:
        return
    compact = compact_text(label)
    if compact in {"company", "companies", "companyname", "companynames", "peers", "peerset"}:
        return
    raise ValueError(
        f"Company header row begins with {label!r}; expected a Function/L1 master "
        "label immediately before the company names."
    )


def _normalized_header_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).casefold())


def _mapped_numeric_value(value: Any, input_label: str, position: int) -> float | None:
    number = to_number_or_none(value)
    if number is not None:
        return number
    if _is_missing_numeric_token(value):
        return None
    raise ValueError(
        f"{input_label} contains a non-numeric value at position {position}: {value!r}."
    )




def _l1_function_lookup_key(value: Any) -> str:
    """Normalize L1 labels for Taxonomy-to-source-row matching."""
    raw = clean_text(value).casefold()
    tokens = re.findall(r"[a-z0-9]+", raw)
    aliases = {
        "mgmt": "management",
        "mgt": "management",
    }
    normalized = [aliases.get(token, token) for token in tokens]
    normalized = [token for token in normalized if token not in {"and", "the"}]
    return "".join(normalized)


def _mapped_range_bounds(
    mapping: dict[str, Any],
    input_label: str,
) -> tuple[str, int, int, int, int]:
    sheet_name = clean_text(mapping.get("sheet_name"))
    range_a1 = clean_text(mapping.get("range_a1"))
    if not sheet_name or not range_a1:
        raise ValueError(f"{input_label} mapping must contain a sheet name and A1 range.")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    except Exception as exc:
        raise ValueError(f"{input_label} has an invalid A1 range: {range_a1!r}.") from exc
    return sheet_name, min_col, min_row, max_col, max_row


def _explicit_function_label_range(
    workbook,
    *,
    fte_sheet: str,
    fte_min_row: int,
    fte_max_row: int,
    function_lookup: dict[str, Any],
    fte_mapping: dict[str, Any],
) -> tuple[list[str], str, int] | None:
    label_sheet = clean_text(
        function_lookup.get("function_label_sheet")
        or fte_mapping.get("row_label_sheet")
        or fte_sheet
    )
    label_range = clean_text(
        function_lookup.get("function_label_range")
        or fte_mapping.get("row_label_range")
    )
    if not label_range or label_sheet != fte_sheet or label_sheet not in workbook.sheetnames:
        return None
    try:
        min_col, min_row, max_col, max_row = range_boundaries(label_range)
    except Exception:
        return None
    if min_col != max_col or min_row != fte_min_row or max_row != fte_max_row:
        return None
    worksheet = workbook[label_sheet]
    labels = [
        clean_text(worksheet.cell(row=row, column=min_col).value)
        for row in range(min_row, max_row + 1)
    ]
    normalized_range = f"{col_letter(min_col)}{min_row}:{col_letter(min_col)}{max_row}"
    return labels, normalized_range, min_col


def _detect_function_label_column_in_bounds(
    workbook,
    *,
    fte_sheet: str,
    fte_min_row: int,
    fte_max_row: int,
    candidate_start_column: int,
    candidate_end_column: int,
    taxonomy_records: list[dict[str, Any]],
) -> tuple[list[str], str, int, dict[str, Any]]:
    """Find the source L1/function column inside a selected descriptor block."""
    worksheet = workbook[fte_sheet]
    taxonomy_keys = {
        _l1_function_lookup_key(record.get("function"))
        for record in taxonomy_records
        if _l1_function_lookup_key(record.get("function"))
    }
    candidates: list[
        tuple[tuple[int, int, int, int, int], int, list[str], dict[str, Any]]
    ] = []
    for column in range(candidate_start_column, candidate_end_column + 1):
        labels = [
            clean_text(worksheet.cell(row=row, column=column).value)
            for row in range(fte_min_row, fte_max_row + 1)
        ]
        keys = [_l1_function_lookup_key(value) for value in labels]
        if not any(keys):
            continue
        match_count = sum(1 for key in keys if key and key in taxonomy_keys)
        unique_match_count = len({key for key in keys if key in taxonomy_keys})
        text_count = sum(1 for value in labels if value and re.search(r"[A-Za-z]", value))
        numeric_count = sum(
            1
            for value in labels
            if value
            and not re.search(r"[A-Za-z]", value)
            and to_number_or_none(value) is not None
        )
        header_hint = 0
        for row in range(max(1, fte_min_row - 12), fte_min_row):
            raw_header = clean_text(worksheet.cell(row=row, column=column).value)
            descriptor = _taxonomy_header_descriptor(raw_header)
            compact = compact_text(raw_header)
            if descriptor is not None and descriptor[0] == 1:
                header_hint = max(header_hint, 3)
            elif "function" in compact or compact in {"l1", "level1", "levelone"}:
                header_hint = max(header_hint, 2)
        distance_from_values = candidate_end_column - column
        score = (
            match_count,
            unique_match_count,
            header_hint,
            text_count - numeric_count * 2,
            -distance_from_values,
        )
        candidates.append(
            (
                score,
                column,
                labels,
                {
                    "match_count": match_count,
                    "unique_match_count": unique_match_count,
                    "header_hint": header_hint,
                    "text_count": text_count,
                    "numeric_count": numeric_count,
                    "distance_from_values": distance_from_values,
                },
            )
        )

    if not candidates:
        raise ValueError(
            "The selected Function + FTE / $B table contains no usable text "
            "column before the company-value columns. Include the L1/function "
            "labels on the left of the numeric matrix."
        )
    candidates.sort(key=lambda item: item[0], reverse=True)
    score, column, labels, details = candidates[0]
    if int(details.get("match_count") or 0) < 1:
        raise ValueError(
            "No function/L1 column inside the selected Function + FTE / $B "
            "table matches the selected Taxonomy L1 values."
        )
    details["score"] = list(score)
    details["candidate_start_column"] = candidate_start_column
    details["candidate_end_column"] = candidate_end_column
    label_range = f"{col_letter(column)}{fte_min_row}:{col_letter(column)}{fte_max_row}"
    return labels, label_range, column, details


def _detect_function_label_column_for_fte(
    workbook,
    *,
    fte_sheet: str,
    fte_min_row: int,
    fte_max_row: int,
    fte_min_col: int,
    taxonomy_records: list[dict[str, Any]],
) -> tuple[list[str], str, int, dict[str, Any]]:
    """Legacy lookup for schema-3/4 numeric-only FTE mappings."""
    scan_start = max(1, fte_min_col - 40)
    if scan_start >= fte_min_col:
        raise ValueError(
            "Could not find an L1/function-name column to the left of the selected "
            "FTE / $B matrix."
        )
    return _detect_function_label_column_in_bounds(
        workbook,
        fte_sheet=fte_sheet,
        fte_min_row=fte_min_row,
        fte_max_row=fte_max_row,
        candidate_start_column=scan_start,
        candidate_end_column=fte_min_col - 1,
        taxonomy_records=taxonomy_records,
    )


def _build_taxonomy_source_crosswalk(
    workbook,
    *,
    taxonomy_records: list[dict[str, Any]],
    fte_mapping: dict[str, Any],
    function_lookup: dict[str, Any] | None,
    logger: logging.Logger,
    schema_version: int,
    company_count: int,
) -> dict[str, Any]:
    function_lookup = dict(function_lookup or {})
    fte_sheet, fte_min_col, fte_min_row, fte_max_col, fte_max_row = (
        _mapped_range_bounds(fte_mapping, "Function + FTE / $B table")
    )
    if fte_sheet not in workbook.sheetnames:
        raise ValueError(
            f"Function + FTE / $B table refers to missing worksheet {fte_sheet!r}. "
            f"Available sheets: {workbook.sheetnames}"
        )

    fte_header_row = fte_min_row if schema_version >= 6 else None
    fte_data_min_row = fte_min_row + 1 if schema_version >= 6 else fte_min_row
    if fte_data_min_row > fte_max_row:
        raise ValueError(
            "Function + FTE / $B table must include its header row and at least "
            "one function data row beneath it."
        )

    selected_column_count = fte_max_col - fte_min_col + 1
    detection_details: dict[str, Any]
    value_start_column = fte_min_col
    descriptor_column_count = 0

    if schema_version >= 7:
        descriptor_column_count = selected_column_count - int(company_count)
        if descriptor_column_count != 1:
            raise ValueError(
                "Function + FTE / $B must start with exactly one L1/function "
                f"column followed by {company_count} company-value columns. "
                "Do not include L0. "
                f"The selected range has {selected_column_count} columns."
            )
        worksheet = workbook[fte_sheet]
        function_header = clean_text(
            worksheet.cell(row=fte_header_row, column=fte_min_col).value
        )
        descriptor = _taxonomy_header_descriptor(function_header)
        if not (
            (descriptor is not None and descriptor[0] == 1)
            or function_header_score(function_header) > 0
        ):
            raise ValueError(
                f"Function + FTE / $B begins with {function_header!r}; the first "
                "selected column must be L1/Function."
            )
        label_column = fte_min_col
        source_labels = [
            clean_text(worksheet.cell(row=row, column=label_column).value)
            for row in range(fte_data_min_row, fte_max_row + 1)
        ]
        label_range = (
            f"{col_letter(label_column)}{fte_data_min_row}:"
            f"{col_letter(label_column)}{fte_max_row}"
        )
        value_start_column = fte_min_col + 1
        detection_details = {
            "origin": "first_selected_l1_column",
            "function_header": function_header,
        }
    elif schema_version >= 5:
        descriptor_column_count = selected_column_count - int(company_count)
        if descriptor_column_count < 1:
            raise ValueError(
                "Function + FTE / $B table must include at least one function/L1 "
                f"column followed by {company_count} company-value columns. "
                f"The selected range has {selected_column_count} columns."
            )
        value_start_column = fte_max_col - int(company_count) + 1
        source_labels, label_range, label_column, detection_details = (
            _detect_function_label_column_in_bounds(
                workbook,
                fte_sheet=fte_sheet,
                fte_min_row=fte_data_min_row,
                fte_max_row=fte_max_row,
                candidate_start_column=fte_min_col,
                candidate_end_column=value_start_column - 1,
                taxonomy_records=taxonomy_records,
            )
        )
        detection_details["origin"] = "inside_selected_fte_table"
    else:
        # Preserve old manifests that stored only the numeric matrix and kept
        # function labels in hidden metadata or an adjacent worksheet column.
        explicit = _explicit_function_label_range(
            workbook,
            fte_sheet=fte_sheet,
            fte_min_row=fte_min_row,
            fte_max_row=fte_max_row,
            function_lookup=function_lookup,
            fte_mapping=fte_mapping,
        )
        if explicit is not None:
            source_labels, label_range, label_column = explicit
            detection_details = {"origin": "manifest_or_mapping_metadata"}
        else:
            source_labels, label_range, label_column, detection_details = (
                _detect_function_label_column_for_fte(
                    workbook,
                    fte_sheet=fte_sheet,
                    fte_min_row=fte_min_row,
                    fte_max_row=fte_max_row,
                    fte_min_col=fte_min_col,
                    taxonomy_records=taxonomy_records,
                )
            )
            detection_details["origin"] = "adjacent_column_detection"

    source_positions: dict[str, list[int]] = {}
    for source_index, source_label in enumerate(source_labels):
        key = _l1_function_lookup_key(source_label)
        if key:
            source_positions.setdefault(key, []).append(source_index)

    taxonomy_key_counts: dict[str, int] = {}
    for record in taxonomy_records:
        key = _l1_function_lookup_key(record.get("function"))
        if key:
            taxonomy_key_counts[key] = taxonomy_key_counts.get(key, 0) + 1

    matches: list[dict[str, Any]] = []
    missing: list[str] = []
    ambiguous: list[str] = []
    used_source_indices: set[int] = set()
    for taxonomy_index, record in enumerate(taxonomy_records):
        function_name = clean_text(record.get("function"))
        key = _l1_function_lookup_key(function_name)
        positions = source_positions.get(key, []) if key else []
        if not positions:
            missing.append(function_name)
            continue
        if len(positions) != 1 or taxonomy_key_counts.get(key, 0) != 1:
            ambiguous.append(function_name)
            continue
        source_index = positions[0]
        used_source_indices.add(source_index)
        matches.append(
            {
                "taxonomy_index": taxonomy_index,
                "taxonomy_row": int(record.get("row") or 0),
                "l0": clean_text(record.get("l0")),
                "taxonomy_function": function_name,
                "source_index": source_index,
                "source_excel_row": fte_data_min_row + source_index,
                "source_function": source_labels[source_index],
                "match_key": key,
            }
        )

    if ambiguous:
        raise ValueError(
            "Multiple FTE source rows match Taxonomy function(s): "
            + ", ".join(ambiguous)
            + ". Function names must be unique in both Taxonomy and the Aura source table."
        )
    if missing:
        raise ValueError(
            "No FTE source row was found for Taxonomy function(s): "
            + ", ".join(missing)
            + f". Source labels were read from {fte_sheet}!{label_range}."
        )
    if len(matches) != len(taxonomy_records):
        raise ValueError(
            f"Matched {len(matches)} of {len(taxonomy_records)} Taxonomy functions "
            "to FTE source rows."
        )

    extras = [
        source_labels[index]
        for index in range(len(source_labels))
        if index not in used_source_indices and clean_text(source_labels[index])
    ]
    value_range = (
        f"{col_letter(value_start_column)}{fte_data_min_row}:"
        f"{col_letter(fte_max_col)}{fte_max_row}"
    )
    logger.info(
        "Matched %s Taxonomy L1 function(s) to source FTE rows using %s!%s. "
        "Numeric FTE values are read from %s!%s; %s extra source row(s) will be ignored.",
        len(matches),
        fte_sheet,
        label_range,
        fte_sheet,
        value_range,
        len(extras),
    )
    logger.debug("Aura L1 function-row crosswalk: %s", matches)
    return {
        "function_label_sheet": fte_sheet,
        "function_label_range": label_range,
        "function_label_column": label_column,
        "fte_header_row": fte_header_row,
        "fte_data_start_row": fte_data_min_row,
        "fte_value_range": value_range,
        "value_start_column": value_start_column,
        "value_end_column": fte_max_col,
        "value_column_count": fte_max_col - value_start_column + 1,
        "descriptor_column_count": descriptor_column_count,
        "source_row_count": len(source_labels),
        "matched_count": len(matches),
        "matches": matches,
        "extra_source": extras,
        "detection": detection_details,
    }


def _taxonomy_sheet_sort_key(sheet_name: str, workbook_order: int) -> tuple[int, int, int]:
    raw = clean_text(sheet_name)
    lowered = raw.casefold()
    compact = compact_text(raw)
    if compact == "taxonomy":
        tier = 0
    elif lowered.startswith("taxonomy") or lowered.endswith("taxonomy"):
        tier = 1
    else:
        tier = 2
    return tier, len(raw), workbook_order


def _resolve_taxonomy_sheet_name(
    workbook,
    requested_sheet: str | None,
    logger: logging.Logger,
) -> str:
    requested = clean_text(requested_sheet)
    if requested:
        matched = find_matching_sheet_name(list(workbook.sheetnames), requested)
        if matched is None:
            raise ValueError(
                f"Taxonomy worksheet {requested!r} was not found. "
                f"Available sheets: {workbook.sheetnames}"
            )
        logger.info("Using requested Taxonomy worksheet: %s", matched)
        return matched

    candidates = [
        (index, name)
        for index, name in enumerate(workbook.sheetnames)
        if "taxonomy" in str(name).casefold()
    ]
    if not candidates:
        raise ValueError(
            "Could not auto-detect a Taxonomy worksheet. The tab name must contain "
            "'Taxonomy' (case-insensitive), or choose it explicitly in the UI. "
            f"Available sheets: {workbook.sheetnames}"
        )

    candidates.sort(key=lambda item: _taxonomy_sheet_sort_key(item[1], item[0]))
    selected = candidates[0][1]
    if len(candidates) > 1:
        logger.warning(
            "Multiple Taxonomy-named worksheets were found: %s. Auto-selected %s; "
            "use the Taxonomy dropdown to choose another.",
            [name for _, name in candidates],
            selected,
        )
    else:
        logger.info("Auto-detected Taxonomy worksheet by tab name: %s", selected)
    return selected



def _taxonomy_header_descriptor(value: Any) -> tuple[int, str, str] | None:
    """Return ``(level, normalized_prefix, display_prefix)`` for a taxonomy header."""
    raw = re.sub(r"\s+", " ", clean_text(value))
    key = compact_text(raw)
    if not key:
        return None
    if key in _TAXONOMY_L0_HEADER_KEYS:
        return 0, "", ""
    if key in _TAXONOMY_L1_HEADER_KEYS:
        return 1, "", ""
    for suffix, level in (
        ("levelzero", 0),
        ("levelone", 1),
        ("level0", 0),
        ("level1", 1),
        ("l0", 0),
        ("l1", 1),
    ):
        if key.endswith(suffix) and len(key) > len(suffix):
            prefix = key[: -len(suffix)]
            display = re.sub(
                r"(?i)(?:\s|[-_/])*?(?:level\s*(?:zero|one|0|1)|l\s*[01])\s*$",
                "",
                raw,
            ).strip(" -_/|")
            return level, prefix, display
    return None


def _detect_taxonomy_columns_in_bounds(
    worksheet,
    *,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int,
) -> tuple[int, int, int, str, str, str]:
    candidates: list[tuple[int, int, int, int, str, str, str]] = []
    for row in range(max(1, min_row), max(1, max_row) + 1):
        descriptors: list[tuple[int, int, str, str, str]] = []
        for column in range(max(1, min_column), max(1, max_column) + 1):
            raw = clean_text(worksheet.cell(row=row, column=column).value)
            descriptor = _taxonomy_header_descriptor(raw)
            if descriptor is None:
                continue
            level, prefix, display_prefix = descriptor
            descriptors.append((column, level, prefix, display_prefix, raw))

        for l0_column, l0_level, l0_prefix, l0_display, l0_raw in descriptors:
            if l0_level != 0:
                continue
            for l1_column, l1_level, l1_prefix, l1_display, l1_raw in descriptors:
                if l1_level != 1 or l0_column == l1_column:
                    continue
                if abs(l1_column - l0_column) != 1 or l0_prefix != l1_prefix:
                    continue
                score = 0
                score += 1000 if not l0_prefix else 820
                score += 150 if l1_column == l0_column + 1 else 120
                score += max(0, 80 - row)
                candidates.append(
                    (
                        score,
                        row,
                        l0_column,
                        l1_column,
                        l0_display or l1_display,
                        l0_raw,
                        l1_raw,
                    )
                )

    if not candidates:
        raise ValueError(
            f"Taxonomy worksheet {worksheet.title!r} must contain an adjacent matching "
            "two-column header pair such as L0/L1, Bain L0/Bain L1, or VC L0/VC L1."
        )
    candidates.sort(key=lambda item: (item[0], -item[1]), reverse=True)
    _score, header_row, l0_column, l1_column, prefix, l0_header, l1_header = candidates[0]
    return header_row, l0_column, l1_column, prefix, l0_header, l1_header

def _detect_taxonomy_columns(worksheet) -> tuple[int, int, int]:
    max_row = max(1, min(int(worksheet.max_row or 1), 100))
    max_column = max(1, min(int(worksheet.max_column or 1), 120))
    header_row, l0_column, l1_column, _prefix, _l0_header, _l1_header = (
        _detect_taxonomy_columns_in_bounds(
            worksheet,
            min_row=1,
            max_row=max_row,
            min_column=1,
            max_column=max_column,
        )
    )
    return header_row, l0_column, l1_column


def _read_l1_taxonomy(
    workbook,
    requested_sheet: str | None,
    logger: logging.Logger,
) -> tuple[str, int, int, int, list[dict[str, Any]]]:
    sheet_name = _resolve_taxonomy_sheet_name(workbook, requested_sheet, logger)
    worksheet = workbook[sheet_name]
    header_row, l0_column, l1_column = _detect_taxonomy_columns(worksheet)

    records: list[dict[str, Any]] = []
    last_l0 = ""
    blank_streak = 0
    scan_end = min(int(worksheet.max_row or header_row), header_row + 5000)
    for row in range(header_row + 1, scan_end + 1):
        raw_l0 = clean_text(worksheet.cell(row=row, column=l0_column).value)
        function_name = clean_text(worksheet.cell(row=row, column=l1_column).value)

        if not raw_l0 and not function_name:
            blank_streak += 1
            if records and blank_streak >= 25:
                break
            continue
        blank_streak = 0

        if compact_text(raw_l0) in _TAXONOMY_L0_HEADER_KEYS and compact_text(function_name) in _TAXONOMY_L1_HEADER_KEYS:
            continue
        if raw_l0:
            last_l0 = l0_display_text(raw_l0)
        if not function_name or function_name.startswith(">>"):
            continue

        l0_name = last_l0 or infer_l0_from_l1(function_name)
        records.append(
            {
                "row": row,
                "l0": l0_name,
                "function": function_name,
            }
        )

    if not records:
        raise ValueError(
            f"Taxonomy worksheet {sheet_name!r} contains no L1 rows beneath its "
            "L0/L1 headers."
        )

    logger.info(
        "Taxonomy source: sheet=%s, header_row=%s, L0_col=%s, L1_col=%s, rows=%s",
        sheet_name,
        header_row,
        col_letter(l0_column),
        col_letter(l1_column),
        len(records),
    )
    return sheet_name, header_row, l0_column, l1_column, records


def _read_l1_taxonomy_from_mapping(
    workbook,
    mapping: dict[str, Any],
    logger: logging.Logger,
) -> tuple[str, int, int, int, list[dict[str, Any]]]:
    """Read the exact two-column Taxonomy range supplied by the UI."""
    sheet_name = clean_text(mapping.get("sheet_name"))
    range_a1 = clean_text(mapping.get("range_a1"))
    if not sheet_name or not range_a1:
        raise ValueError("Taxonomy mapping must contain a worksheet name and A1 range.")
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Taxonomy mapping refers to missing worksheet {sheet_name!r}. "
            f"Available sheets: {workbook.sheetnames}"
        )
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    except Exception as exc:
        raise ValueError(f"Taxonomy has an invalid A1 range: {range_a1!r}.") from exc
    if max_col - min_col + 1 != 2:
        raise ValueError(
            f"Taxonomy must be exactly two columns; received "
            f"{max_col - min_col + 1} columns in {range_a1}."
        )
    if max_row - min_row + 1 < 2:
        raise ValueError("Taxonomy must include a header row and at least one data row.")

    worksheet = workbook[sheet_name]
    (
        header_row,
        l0_column,
        l1_column,
        header_prefix,
        l0_header,
        l1_header,
    ) = _detect_taxonomy_columns_in_bounds(
        worksheet,
        min_row=min_row,
        max_row=min(max_row, min_row + 30),
        min_column=min_col,
        max_column=max_col,
    )

    records: list[dict[str, Any]] = []
    last_l0 = ""
    for row in range(header_row + 1, max_row + 1):
        raw_l0 = clean_text(worksheet.cell(row=row, column=l0_column).value)
        function_name = clean_text(worksheet.cell(row=row, column=l1_column).value)
        if not raw_l0 and not function_name:
            continue
        l0_descriptor = _taxonomy_header_descriptor(raw_l0)
        l1_descriptor = _taxonomy_header_descriptor(function_name)
        if (
            l0_descriptor is not None
            and l1_descriptor is not None
            and l0_descriptor[0] == 0
            and l1_descriptor[0] == 1
        ):
            continue
        if raw_l0:
            last_l0 = l0_display_text(raw_l0)
        if not function_name or function_name.startswith(">>"):
            continue
        records.append(
            {
                "row": row,
                "l0": last_l0 or infer_l0_from_l1(function_name),
                "function": function_name,
            }
        )

    if not records:
        raise ValueError(
            f"Taxonomy range {sheet_name}!{range_a1} contains no L1 rows beneath "
            "its detected header pair."
        )

    logger.info(
        "Taxonomy input: sheet=%s, range=%s, headers=%s/%s, prefix=%s, "
        "header_row=%s, L0_col=%s, L1_col=%s, rows=%s",
        sheet_name,
        range_a1,
        l0_header,
        l1_header,
        header_prefix or "<none>",
        header_row,
        col_letter(l0_column),
        col_letter(l1_column),
        len(records),
    )
    return sheet_name, header_row, l0_column, l1_column, records


def _read_source_data_from_l1_input_map(
    config: BuildConfig,
    logger: logging.Logger,
    manifest_path: Path,
) -> tuple[list[str], list[Any], list[dict[str, Any]]]:
    payload = _load_l1_input_map(manifest_path)
    schema_version = int(payload.get("schema_version", -1))
    mapped_workbook = clean_text(payload.get("workbook_path"))
    if mapped_workbook:
        expected = Path(mapped_workbook).expanduser().resolve()
        actual = Path(config.excel_path).expanduser().resolve()
        if expected != actual:
            raise ValueError(
                "Aura L1 input map belongs to a different workbook: "
                f"mapped={expected}, selected={actual}."
            )

    inputs = payload["inputs"]
    workbook = load_workbook_for_mapped_ranges(config.excel_path, payload, logger)
    try:
        (
            taxonomy_sheet,
            taxonomy_header_row,
            taxonomy_l0_column,
            taxonomy_l1_column,
            taxonomy_records,
        ) = _read_l1_taxonomy_from_mapping(
            workbook,
            inputs["taxonomy"],
            logger,
        )

        ignored_taxonomy_functions = [
            clean_text(value)
            for value in (payload.get("ignored_taxonomy_functions") or [])
            if clean_text(value)
        ]
        if ignored_taxonomy_functions:
            ignored_keys = {
                _l1_function_lookup_key(value)
                for value in ignored_taxonomy_functions
                if _l1_function_lookup_key(value)
            }
            skipped_records = [
                record for record in taxonomy_records
                if _l1_function_lookup_key(record.get("function")) in ignored_keys
            ]
            taxonomy_records = [
                record for record in taxonomy_records
                if _l1_function_lookup_key(record.get("function")) not in ignored_keys
            ]
            if skipped_records:
                logger.warning(
                    "Ignoring %s Taxonomy/FTE unmatched function(s) by user request: %s. "
                    "These rows will not generate Aura function slides or contribute to Aura L0 aggregation.",
                    len(skipped_records),
                    ", ".join(clean_text(record.get("function")) for record in skipped_records),
                )

        company_matrix = _mapped_range_matrix(
            workbook, inputs["company_headers"], "Company names"
        )
        if schema_version >= 7:
            company_master_label = ""
            raw_company_headers = _mapped_plain_row(
                company_matrix, "Company names"
            )
        elif schema_version >= 6:
            company_master_label, raw_company_headers = _mapped_labeled_row(
                company_matrix, "Company header row"
            )
            _validate_company_master_label(company_master_label)
        else:
            company_master_label = ""
            raw_company_headers = _mapped_vector(company_matrix, "Company names")

        company_headers = [clean_text(value) for value in raw_company_headers]
        if len(company_headers) < 2:
            raise ValueError(
                "Company names must contain the client plus at least one peer company."
            )
        blank_companies = [index + 1 for index, value in enumerate(company_headers) if not value]
        if blank_companies:
            raise ValueError(
                "Company names contain blank cell(s) at position(s): "
                + ", ".join(str(value) for value in blank_companies)
            )

        revenue_matrix = _mapped_range_matrix(
            workbook, inputs["revenue_values"], "Revenue row"
        )
        if schema_version >= 6:
            revenue_master_label, raw_revenue = _mapped_labeled_row(
                revenue_matrix, "Revenue row"
            )
            if revenue_label_score(revenue_master_label) <= 0:
                raise ValueError(
                    f"Revenue row begins with {revenue_master_label!r}; the first "
                    "selected cell must be the Revenue master label immediately "
                    "before the company values."
                )
        else:
            revenue_master_label = ""
            raw_revenue = _mapped_vector(revenue_matrix, "Revenue values")

        if len(raw_revenue) != len(company_headers):
            raise ValueError(
                f"Revenue row contains {len(raw_revenue)} company values; "
                f"Company names contains {len(company_headers)} companies."
            )
        revenue_values: list[Any] = []
        for position, value in enumerate(raw_revenue, start=1):
            number = _mapped_numeric_value(value, "Revenue values", position)
            revenue_values.append("-" if number is None else number)

        has_benchmark_inputs = False
        l1_quartile_profile: dict[str, Any] | None = None
        if schema_version >= 9 and isinstance(inputs.get("l1_quartile_table"), dict):
            has_benchmark_inputs = True
            l1_quartile_profile = _mapped_quartile_table(
                workbook,
                inputs["l1_quartile_table"],
                "L1 quartile table detector",
            )
            raw_top = list(l1_quartile_profile["raw_top"])
            raw_median = list(l1_quartile_profile["raw_median"])
            raw_bottom = list(l1_quartile_profile["raw_bottom"])
            top_master_label = str(l1_quartile_profile["headers"]["top"]["label"])
            median_master_label = str(l1_quartile_profile["headers"]["median"]["label"])
            bottom_master_label = str(l1_quartile_profile["headers"]["bottom"]["label"])
        else:
            legacy_benchmark_inputs = all(
                isinstance(inputs.get(key), dict)
                for key in LEGACY_L1_BENCHMARK_INPUT_KEYS
            )
            has_benchmark_inputs = legacy_benchmark_inputs
            if legacy_benchmark_inputs and schema_version >= 6:
                top_master_label, raw_top = _mapped_labeled_column(
                    _mapped_range_matrix(
                        workbook, inputs["top_quartile"], "Top quartile column"
                    ),
                    "Top quartile column",
                )
                median_master_label, raw_median = _mapped_labeled_column(
                    _mapped_range_matrix(workbook, inputs["median"], "Median column"),
                    "Median column",
                )
                bottom_master_label, raw_bottom = _mapped_labeled_column(
                    _mapped_range_matrix(
                        workbook, inputs["bottom_quartile"], "Bottom quartile column"
                    ),
                    "Bottom quartile column",
                )
                for label, observed, expected in (
                    ("Top quartile column", top_master_label, "top"),
                    ("Median column", median_master_label, "median"),
                    ("Bottom quartile column", bottom_master_label, "bottom"),
                ):
                    if quartile_kind(observed) != expected:
                        raise ValueError(
                            f"{label} begins with {observed!r}; its first selected cell "
                            "does not contain the required master label."
                        )
            elif legacy_benchmark_inputs:
                top_master_label = median_master_label = bottom_master_label = ""
                raw_top = _mapped_vector(
                    _mapped_range_matrix(workbook, inputs["top_quartile"], "Top quartile"),
                    "Top quartile",
                )
                raw_median = _mapped_vector(
                    _mapped_range_matrix(workbook, inputs["median"], "Median"),
                    "Median",
                )
                raw_bottom = _mapped_vector(
                    _mapped_range_matrix(workbook, inputs["bottom_quartile"], "Bottom quartile"),
                    "Bottom quartile",
                )
            else:
                top_master_label = median_master_label = bottom_master_label = ""
                raw_top = raw_median = raw_bottom = None


        raw_fte_complete = _mapped_range_matrix(
            workbook,
            inputs["fte_values"],
            "Function + FTE / $B table",
        )
        if schema_version >= 6:
            if len(raw_fte_complete) < 2:
                raise ValueError(
                    "Function + FTE / $B table must include its header row and at "
                    "least one function data row."
                )
            fte_header = [clean_text(value) for value in raw_fte_complete[0]]
            raw_fte_table = raw_fte_complete[1:]
            descriptor_count = len(fte_header) - len(company_headers)
            if schema_version >= 7:
                if descriptor_count != 1:
                    raise ValueError(
                        "Function + FTE / $B must start with exactly one L1/function "
                        "column followed by the company columns. Do not include L0."
                    )
                function_header = fte_header[0]
                descriptor = _taxonomy_header_descriptor(function_header)
                if not (
                    (descriptor is not None and descriptor[0] == 1)
                    or function_header_score(function_header) > 0
                ):
                    raise ValueError(
                        f"Function + FTE / $B begins with {function_header!r}; the "
                        "first selected column must be L1/Function."
                    )
                fte_company_headers = fte_header[1:]
            else:
                if descriptor_count < 1:
                    raise ValueError(
                        "Function + FTE / $B table must include at least one descriptor "
                        "header before the company headers."
                    )
                descriptor_headers = fte_header[:descriptor_count]
                fte_company_headers = fte_header[descriptor_count:]
                has_function_header = any(
                    (
                        (_taxonomy_header_descriptor(value) is not None
                         and _taxonomy_header_descriptor(value)[0] == 1)
                        or function_header_score(value) > 0
                    )
                    for value in descriptor_headers
                    if value
                )
                if not has_function_header:
                    raise ValueError(
                        "Function + FTE / $B table header must include an L1/Function "
                        "master label in its descriptor columns."
                    )
            if [
                _normalized_header_key(value) for value in fte_company_headers
            ] != [
                _normalized_header_key(value) for value in company_headers
            ]:
                raise ValueError(
                    "The company headers inside Function + FTE / $B do not match "
                    "the selected Company names in the same order."
                )
        else:
            fte_header = []
            raw_fte_table = raw_fte_complete

        source_row_count = len(raw_fte_table)
        if raw_top is None:
            raw_top = [None] * source_row_count
            raw_median = [None] * source_row_count
            raw_bottom = [None] * source_row_count
            logger.info(
                "Aura input map contains the shared L0/L1 ranges without L1 "
                "benchmark columns; quartile values will remain blank for this read."
            )
        else:
            for label, values in (
                ("Top quartile", raw_top),
                ("Median", raw_median),
                ("Bottom quartile", raw_bottom),
            ):
                if len(values) != source_row_count:
                    raise ValueError(
                        f"{label} contains {len(values)} benchmark values; Function + "
                        f"FTE / $B contains {source_row_count} source data rows beneath "
                        "its header."
                    )

        function_crosswalk = _build_taxonomy_source_crosswalk(
            workbook,
            taxonomy_records=taxonomy_records,
            fte_mapping=inputs["fte_values"],
            function_lookup=dict(payload.get("function_lookup") or {}),
            logger=logger,
            schema_version=schema_version,
            company_count=len(company_headers),
        )

        _fte_sheet, fte_min_col, _fte_min_row, _fte_max_col, _fte_max_row = (
            _mapped_range_bounds(
                inputs["fte_values"],
                "Function + FTE / $B table",
            )
        )
        value_start_column = int(
            function_crosswalk.get("value_start_column") or fte_min_col
        )
        value_start_offset = value_start_column - fte_min_col
        if value_start_offset < 0:
            raise ValueError(
                "The detected FTE company-value columns fall outside the selected table."
            )

        raw_fte: list[list[Any]] = []
        for row_index, row in enumerate(raw_fte_table, start=1):
            values = list(
                row[value_start_offset : value_start_offset + len(company_headers)]
            )
            if len(values) != len(company_headers):
                raise ValueError(
                    f"Function + FTE / $B table row {row_index} does not contain "
                    f"{len(company_headers)} company values after its function columns."
                )
            raw_fte.append(values)

        if int(function_crosswalk.get("value_column_count") or 0) != len(company_headers):
            raise ValueError(
                "Function + FTE / $B table company-value width does not match "
                f"Company names: values={function_crosswalk.get('value_column_count')}, "
                f"companies={len(company_headers)}."
            )

        records: list[dict[str, Any]] = []
        for match in function_crosswalk["matches"]:
            taxonomy_index = int(match["taxonomy_index"])
            source_index = int(match["source_index"])
            taxonomy_record = taxonomy_records[taxonomy_index]
            function_name = clean_text(taxonomy_record.get("function"))
            l0_name = clean_text(taxonomy_record.get("l0")) or infer_l0_from_l1(function_name)
            fte_values = [to_int_or_none(value) for value in raw_fte[source_index]]
            source_position = source_index + 1
            quartile_values = [
                _mapped_numeric_value(
                    raw_top[source_index],
                    "Top quartile",
                    source_position,
                ),
                _mapped_numeric_value(
                    raw_median[source_index],
                    "Median",
                    source_position,
                ),
                _mapped_numeric_value(
                    raw_bottom[source_index],
                    "Bottom quartile",
                    source_position,
                ),
            ]

            if all(value is None for value in fte_values) and all(
                value is None for value in quartile_values
            ):
                logger.debug(
                    "Skipping taxonomy row %s (%s / %s) because matched source row %s is blank.",
                    taxonomy_record.get("row"),
                    l0_name,
                    function_name,
                    match.get("source_excel_row"),
                )
                continue

            records.append(
                {
                    "l0": l0_name,
                    "level0": l0_name,
                    "function": function_name,
                    "l1": function_name,
                    "level1": function_name,
                    "fte_values": fte_values,
                    "quartile_values": quartile_values,
                    "source_function": clean_text(match.get("source_function")),
                    "source_row": int(match.get("source_excel_row") or 0),
                }
            )

        if not records:
            raise ValueError(
                "The mapped Aura ranges and selected Taxonomy tab contain no "
                "chart-ready function rows."
            )

        config.autodetect_source_layout = False
        config.num_companies = len(company_headers)
        config.chart_cols = len(company_headers) + 1
        config.sheet_name = clean_text(inputs["fte_values"].get("sheet_name"))
        config.taxonomy_sheet_name = taxonomy_sheet
        config.detected_source_layout = {
            "mode": "named_ranges_with_taxonomy",
            "manifest_path": str(manifest_path),
            "taxonomy": {
                "sheet_name": taxonomy_sheet,
                "header_row": taxonomy_header_row,
                "l0_column": taxonomy_l0_column,
                "l1_column": taxonomy_l1_column,
                "row_count": len(taxonomy_records),
            },
            "function_lookup": function_crosswalk,
            "inputs": {
                key: {
                    "sheet_name": clean_text(inputs[key].get("sheet_name")),
                    "range_a1": clean_text(inputs[key].get("range_a1")),
                }
                for key in AURA_INPUT_KEYS
                if isinstance(inputs.get(key), dict)
            },
        }

        logger.info("Using explicit standardized Aura worksheet range mappings from %s", manifest_path)
        logger.info("Using L0/L1 labels from Taxonomy worksheet: %s", taxonomy_sheet)
        logger.info(
            "Taxonomy function order is matched to FTE source rows through %s!%s.",
            function_crosswalk["function_label_sheet"],
            function_crosswalk["function_label_range"],
        )
        logger.info("Mapped company headers: %s", company_headers)
        logger.info("Mapped revenue values: %s", revenue_values)
        logger.info("Mapped %s Aura function row(s).", len(records))
        for record in records:
            logger.debug(
                "%s / %s | FTE=%s | Quartiles=%s",
                record.get("l0", ""),
                record["function"],
                record["fte_values"],
                record["quartile_values"],
            )
        return company_headers, revenue_values, records
    finally:
        try:
            workbook.close()
        except Exception:
            pass

def apply_layout_to_config(config: BuildConfig, layout: SourceLayout) -> None:
    config.sheet_name = layout.sheet_name
    config.header_row_excel = layout.header_row
    config.function_col = layout.function_col
    config.company_start_col = layout.company_start_col
    config.company_end_col = layout.company_end_col
    config.quartile_start_col = layout.quartile_start_col
    config.quartile_end_col = layout.quartile_end_col
    config.num_companies = layout.company_end_col - layout.company_start_col + 1
    config.chart_cols = int(config.num_companies) + 1
    config.l0_col = layout.l0_col
    config.detected_source_layout = asdict(layout)


def read_source_data(config: BuildConfig, logger: logging.Logger) -> tuple[list[str], list[Any], list[dict[str, Any]]]:
    require_file(config.excel_path, "Excel workbook")

    with stage(logger, "Read source Excel"):
        mapped_input_path = (
            getattr(config, "aura_input_map_path", None)
            or getattr(config, "l1_input_map_path", None)
        )
        if mapped_input_path is not None:
            return _read_source_data_from_l1_input_map(
                config,
                logger,
                Path(mapped_input_path).expanduser().resolve(),
            )

        logger.info("Opening workbook in read_only values mode and parsing only the Aura Summary tab.")
        wb = load_workbook_fast(config.excel_path)

        try:
            ws = choose_aura_summary_sheet(wb, config, logger)
            layout = detect_fte_table_layout(ws, config, logger)
            apply_layout_to_config(config, layout)

            l0_location = f"{col_letter(layout.l0_col)}{layout.header_row}" if layout.l0_col else "not detected"
            logger.info(
                "Detected FTE/$B table: sheet=%s, header=%s, rows=%s:%s, L0=%s, L1=%s%s, companies=%s:%s, quartiles=%s/%s/%s",
                layout.sheet_name,
                layout.header_row,
                layout.fte_start_row,
                layout.fte_end_row,
                l0_location,
                col_letter(layout.function_col),
                layout.header_row,
                f"{col_letter(layout.company_start_col)}{layout.header_row}",
                f"{col_letter(layout.company_end_col)}{layout.header_row}",
                f"{col_letter(layout.quartile_columns['top'])}{layout.header_row}",
                f"{col_letter(layout.quartile_columns['median'])}{layout.header_row}",
                f"{col_letter(layout.quartile_columns['bottom'])}{layout.header_row}",
            )

            logger.info(
                "Detected revenue row: %s%s (%s)",
                col_letter(layout.revenue_label_col),
                layout.revenue_row,
                clean_text(ws.value(layout.revenue_row, layout.revenue_label_col)),
            )

            company_header_row = layout.company_header_row or layout.header_row
            company_header_start_col = layout.company_header_start_col or layout.company_start_col
            company_header_end_col = layout.company_header_end_col or layout.company_end_col

            company_headers: list[str] = [
                clean_text(ws.value(company_header_row, col))
                for col in range(company_header_start_col, company_header_end_col + 1)
            ]
            fte_company_headers = [
                clean_text(ws.value(layout.header_row, col))
                for col in range(layout.company_start_col, layout.company_end_col + 1)
            ]
            if [compact_text(value) for value in company_headers] != [
                compact_text(value) for value in fte_company_headers
            ]:
                raise ValueError(
                    "The canonical company row in the first Aura table does not match "
                    "the company headers in the detected FTE/$B table."
                )
            logger.info(
                "Canonical company row: %s%s:%s%s",
                col_letter(company_header_start_col),
                company_header_row,
                col_letter(company_header_end_col),
                company_header_row,
            )

            revenue_values: list[Any] = []
            for col in range(company_header_start_col, company_header_end_col + 1):
                raw_value = ws.value(layout.revenue_row, col)

                if raw_value is None or clean_text(raw_value) in {"", "-", "NA", "N/A", "n/a"}:
                    revenue_values.append("-")
                else:
                    revenue_values.append(to_number_or_none(raw_value))

            qcols = [
                layout.quartile_columns["top"],
                layout.quartile_columns["median"],
                layout.quartile_columns["bottom"],
            ]

            records: list[dict[str, Any]] = []
            last_l0 = ""

            for row in range(layout.fte_start_row, layout.fte_end_row + 1):
                function_name = clean_text(ws.value(row, layout.function_col))

                if not function_name or function_name.startswith(">>"):
                    continue

                row_l0 = l0_from_row_or_above(
                    ws=ws,
                    row=row,
                    l0_col=layout.l0_col,
                    header_row=layout.header_row,
                )

                if row_l0:
                    last_l0 = row_l0

                l0_name = row_l0 or last_l0 or infer_l0_from_l1(function_name)

                fte_values = [
                    to_int_or_none(ws.value(row, col))
                    for col in range(layout.company_start_col, layout.company_end_col + 1)
                ]

                quartile_values = [
                    to_number_or_none(ws.value(row, col))
                    for col in qcols
                ]

                if all(value is None for value in fte_values) and all(value is None for value in quartile_values):
                    continue

                records.append(
                    {
                        "l0": l0_name,
                        "level0": l0_name,
                        "function": function_name,
                        "l1": function_name,
                        "level1": function_name,
                        "fte_values": fte_values,
                        "quartile_values": quartile_values,
                    }
                )

            logger.info("Company headers: %s", company_headers)
            logger.info("Revenue values: %s", revenue_values)
            logger.info("Detected %s FTE rows.", len(records))

            for record in records:
                logger.debug(
                    "%s / %s | FTE=%s | Quartiles=%s",
                    record.get("l0", ""),
                    record["function"],
                    record["fte_values"],
                    record["quartile_values"],
                )

            if not records:
                raise ValueError("Detected the FTE/$B table, but no chart-ready function rows were parsed.")

            return company_headers, revenue_values, records

        finally:
            try:
                wb.close()
            except Exception:
                pass


__all__ = [
    "SourceLayout",
    "read_source_data",
    "select_summary_sheet",
    "choose_aura_summary_sheet",
    "detect_fte_table_layout",
    "find_revenue_row",
]