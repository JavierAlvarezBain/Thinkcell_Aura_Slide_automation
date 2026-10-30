from __future__ import annotations

import html
import json
import logging
import re
import warnings
import zipfile
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries

try:
    from shared import build_company_order
except ModuleNotFoundError:  # Direct Geo Cost script execution.
    import sys

    _APP_ROOT = Path(__file__).resolve().parents[2]
    if str(_APP_ROOT) not in sys.path:
        sys.path.insert(0, str(_APP_ROOT))
    from shared import build_company_order

from geocost_common import (
    CountryRow,
    FunctionBlock,
    GeoCostData,
    GeoCostScope,
    GeoCostSheetMatch,
    GeoCostSourceLayout,
    GeoCostTemplateLayout,
    GeoCostWorkbookDetection,
    LabelTarget,
    PLAIN_PERCENT_NUMBER_FORMAT,
    PLAIN_VALUE_NUMBER_FORMAT,
    US_STATE_KEYS,
    canonical_country_key,
    clean_text,
    compact_text,
    display_label_for_country,
    normalize_high_low,
    normalize_segment_label,
    repair_mojibake,
    safe_filename,
    to_number,
)


warnings.filterwarnings(
    "ignore",
    message=r"Cannot parse header or footer.*",
    category=UserWarning,
    module=r"openpyxl\.worksheet\.header_footer",
)


def _get(rows: list[list[Any]], row_num: int, col_num: int) -> Any:
    row_index = row_num - 1
    col_index = col_num - 1
    if row_index < 0 or row_index >= len(rows):
        return None
    row = rows[row_index]
    if col_index < 0 or col_index >= len(row):
        return None
    return row[col_index]


def _max_col_count(rows: list[list[Any]]) -> int:
    return max((len(row) for row in rows), default=0)


GEOCOST_INPUT_MAP_SCHEMA_VERSION = 2


def _geocost_header_kind(value: Any) -> str:
    key = compact_text(value)
    if key == "region":
        return "region"
    if key in {"highlow", "highorlow", "costtier"}:
        return "high_low"
    if key in {"country", "state", "location", "geography"}:
        return "entity"
    if key in {
        "geocostanalysis", "geocostlocation", "displaycountry", "displaystate",
        "countryname", "statename",
    }:
        return "display"
    return ""


def _low_cost_row_label_score(value: Any) -> int:
    raw = clean_text(value)
    text = re.sub(r"[^a-z0-9%]+", " ", raw.lower()).strip()
    compact = compact_text(raw)
    tokens = set(text.replace("%", " % ").split())
    if not text or "low" not in tokens or "cost" not in tokens:
        return 0
    score = 80
    if "%" in raw or "percent" in tokens or "percentage" in tokens or "pct" in tokens:
        score += 45
    if "fte" in compact or "ftes" in compact:
        score += 55
    if "region" in tokens or "regions" in tokens:
        score += 25
    if "location" in tokens or "locations" in tokens:
        score += 25
    if "workforce" in tokens or "employee" in tokens or "employees" in tokens:
        score += 20
    return score


def _optional_ratio(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        raw = value.strip()
        cleaned = raw.replace("%", "").replace(",", "").replace("\u00a0", "").strip()
        if cleaned.lower() in {"", "-", "na", "n/a", "nan", "none"}:
            return None
        try:
            number = float(cleaned)
        except ValueError:
            return None
        return number / 100.0 if "%" in raw or abs(number) > 1.0000001 else number
    try:
        number = float(value)
    except Exception:
        return None
    return number / 100.0 if abs(number) > 1.0000001 else number


def load_geocost_input_manifest(
    path: Path | str | None,
    *,
    excel_path: Path | None = None,
) -> dict[str, Any] | None:
    """Read and validate the standardized Geo Cost table + low-cost-row manifest."""
    if path is None or not str(path).strip():
        return None
    manifest_path = Path(path).expanduser().resolve()
    if not manifest_path.is_file():
        raise FileNotFoundError(f"Geo Cost input map not found: {manifest_path}")
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    if int(payload.get("schema_version", -1)) != GEOCOST_INPUT_MAP_SCHEMA_VERSION:
        raise ValueError(
            "Unsupported Geo Cost input-map schema: "
            f"{payload.get('schema_version')!r}. Reopen the workbook so the Geo Cost "
            "table and Low Cost % row are detected again."
        )
    if str(payload.get("slide_family") or "").strip().lower() != "geocost":
        raise ValueError("The supplied input map is not a Geo Cost manifest.")
    sheet_name = clean_text(payload.get("sheet_name"))
    if not sheet_name:
        raise ValueError("Geo Cost input map has no worksheet name.")
    inputs = payload.get("inputs") or {}
    if not isinstance(inputs, dict):
        raise ValueError("Geo Cost input map inputs must be an object.")
    for key in ("geocost_table", "geocost_data_row"):
        mapping = inputs.get(key)
        if not isinstance(mapping, dict):
            raise ValueError(f"Geo Cost input map is missing {key!r}.")
        if not clean_text(mapping.get("sheet_name")) or not clean_text(mapping.get("range_a1")):
            raise ValueError(f"Geo Cost input map {key!r} has no worksheet/range.")
        if clean_text(mapping.get("sheet_name")).casefold() != sheet_name.casefold():
            raise ValueError(
                "Geo Cost table and low-cost data row must belong to the manifest worksheet."
            )
    if excel_path is not None:
        declared = clean_text(payload.get("workbook_path"))
        if declared:
            try:
                if Path(declared).resolve() != Path(excel_path).resolve():
                    raise ValueError(
                        "Geo Cost input map belongs to a different workbook: "
                        f"manifest={declared!r}, current={str(Path(excel_path).resolve())!r}."
                    )
            except OSError:
                pass
    payload["_manifest_path"] = str(manifest_path)
    return payload


def source_layout_from_geocost_input_manifest(
    rows: list[list[Any]],
    payload: dict[str, Any],
) -> GeoCostSourceLayout:
    """Resolve helper columns and function markers around the two visible ranges."""
    inputs = dict(payload.get("inputs") or {})
    table_mapping = dict(inputs.get("geocost_table") or {})
    low_cost_mapping = dict(inputs.get("geocost_data_row") or {})
    table_range = clean_text(table_mapping.get("range_a1"))
    low_cost_range = clean_text(low_cost_mapping.get("range_a1"))
    table_min_col, table_min_row, table_max_col, table_max_row = range_boundaries(
        table_range
    )
    low_min_col, low_min_row, low_max_col, low_max_row = range_boundaries(
        low_cost_range
    )
    if low_min_row != low_max_row:
        raise ValueError("Geo Cost data row must be exactly one horizontal row.")
    if table_max_row <= table_min_row:
        raise ValueError(
            "Geo Cost table must include its company header and at least one geography row."
        )
    if low_min_row <= table_max_row:
        raise ValueError(
            "Geo Cost data row must be below the final geography row of the Geo Cost table."
        )

    detected = detect_source_layout(rows)
    if detected is None:
        raise ValueError(
            "Could not infer the Geo Cost helper columns and >>> function-marker row around "
            "the selected ranges."
        )
    base_layout, _score = detected
    if int(base_layout.header_row) != int(table_min_row):
        raise ValueError(
            "Geo Cost table must start on the detected company-header row: "
            f"selected row {table_min_row}, detected row {base_layout.header_row}."
        )

    blocks = find_function_blocks(rows, base_layout)
    if not blocks:
        raise ValueError("No >>> / function blocks were found above the selected Geo Cost table.")

    selected_block: FunctionBlock | None = None
    for block in blocks:
        stop_col = block.next_start_col or (_max_col_count(rows) + 1)
        local_country_col = max(1, block.start_col - 1)
        if not (table_max_col < local_country_col or table_min_col >= stop_col):
            selected_block = block
            break
    if selected_block is None:
        raise ValueError("The selected Geo Cost table does not intersect the Overall function block.")

    stop_col = (selected_block.next_start_col - 1) if selected_block.next_start_col else (_max_col_count(rows) + 1)
    local_country_col = max(1, selected_block.start_col - 1)
    if not (table_min_col <= local_country_col <= table_max_col):
        raise ValueError(
            "Geo Cost table must include the local Country/State column immediately before "
            "the Overall company headers."
        )
    company_cols = [
        col_num
        for col_num in range(selected_block.start_col, stop_col)
        if clean_text(_get(rows, table_min_row, col_num))
        and table_min_col <= col_num <= table_max_col
    ]
    if not company_cols:
        raise ValueError("The selected Geo Cost table contains no Overall company headers.")
    company_start_col = min(company_cols)
    company_end_col = max(company_cols)
    if table_max_col >= stop_col:
        raise ValueError(
            "Geo Cost table extends into the next function block; select only the Overall block."
        )

    valid_rows = 0
    for row_num in range(table_min_row + 1, min(table_max_row, len(rows)) + 1):
        bucket = clean_text(_get(rows, row_num, base_layout.segment_bucket_col))
        high_low = clean_text(_get(rows, row_num, base_layout.high_low_col)).lower()
        entity = (
            clean_text(_get(rows, row_num, local_country_col))
            or clean_text(_get(rows, row_num, base_layout.display_country_col))
            or clean_text(_get(rows, row_num, base_layout.raw_country_col))
        )
        if entity and bucket and ("-" in bucket or high_low in {"high", "low"}):
            valid_rows += 1
    if valid_rows < 1:
        raise ValueError(
            "Geo Cost table contains no geography rows aligned to the detected helper columns."
        )

    label_candidates = [
        (col_num, _low_cost_row_label_score(_get(rows, low_min_row, col_num)))
        for col_num in range(low_min_col, low_max_col + 1)
    ]
    label_candidates = [item for item in label_candidates if item[1] > 0]
    if not label_candidates:
        raise ValueError(
            "Geo Cost data row must include a label such as '% of FTEs in Low Cost region', "
            "'FTEs in low-cost locations (%)', or 'Low cost %'."
        )
    low_cost_label_col = max(label_candidates, key=lambda item: item[1])[0]
    missing_company_cols = [
        col_num for col_num in company_cols if not (low_min_col <= col_num <= low_max_col)
    ]
    if missing_company_cols:
        raise ValueError(
            "Geo Cost data row must include the low-cost percentage for every Overall company."
        )
    ratios = [_optional_ratio(_get(rows, low_min_row, col_num)) for col_num in company_cols]
    if not any(value is not None for value in ratios):
        raise ValueError(
            "Geo Cost data row contains no numeric low-cost percentage values under the company headers."
        )

    return GeoCostSourceLayout(
        marker_row=base_layout.marker_row,
        header_row=table_min_row,
        data_start_row=table_min_row + 1,
        segment_bucket_col=base_layout.segment_bucket_col,
        high_low_col=base_layout.high_low_col,
        raw_country_col=base_layout.raw_country_col,
        display_country_col=base_layout.display_country_col,
        entity_label=base_layout.entity_label,
        data_end_row=table_max_row,
        low_cost_row=low_min_row,
        low_cost_label_col=low_cost_label_col,
        company_start_col=company_start_col,
        company_end_col=company_end_col,
        marker_start_col=blocks[0].start_col,
        marker_end_col=_max_col_count(rows),
        table_start_col=table_min_col,
        table_end_col=table_max_col,
    )



# =============================================================================
# PowerPoint template detection


def iter_thinkcell_roots(pptx_path: Path):
    with zipfile.ZipFile(pptx_path, "r") as archive:
        for member in archive.namelist():
            if not member.startswith("ppt/embeddings/") or not member.endswith(".bin"):
                continue
            data = archive.read(member)
            for match in re.finditer(rb"<root\b.*?</root>", data, flags=re.S):
                yield member, match.group(0).decode("latin-1")


def extract_template_chart_names(template_path: Path) -> list[str]:
    names: list[str] = []
    for _, root in iter_thinkcell_roots(template_path):
        for match in re.finditer(r"<m_strName>(.*?)</m_strName>", root, flags=re.S):
            name = html.unescape(match.group(1)).strip()
            if name and name not in names:
                names.append(name)
    return names


def resolve_template_chart_name(
    template_path: Path,
    requested_chart_name: str | None,
    logger: logging.Logger,
) -> str:
    available = extract_template_chart_names(template_path)
    requested = clean_text(requested_chart_name)
    if requested:
        match = next((name for name in available if name.lower() == requested.lower()), None)
        if match is None:
            raise ValueError(
                f"Requested think-cell name {requested!r} is not present. "
                f"Available names: {available or '<none>'}."
            )
        logger.info("Using verified think-cell chart name from template: %s", match)
        return match

    preferred = next((name for name in available if compact_text(name) == "geocostchart"), None)
    if preferred:
        logger.info("Auto-detected think-cell chart name from template: %s", preferred)
        return preferred
    geocost_names = [name for name in available if "geo" in name.lower() and "cost" in name.lower()]
    if len(geocost_names) == 1:
        logger.info("Auto-detected think-cell chart name from template: %s", geocost_names[0])
        return geocost_names[0]
    if len(available) == 1:
        logger.info("Auto-detected think-cell chart name from template: %s", available[0])
        return available[0]
    raise ValueError(
        "Could not choose a unique think-cell AddRangeData name. "
        f"Available names: {available or '<none>'}. Pass --chart-name."
    )


def _extract_table_series_names(root: str) -> list[list[str]]:
    """Extract candidate think-cell series-name tables of any length."""
    vector_starts = [m.start() for m in re.finditer(r"<CSequenceChartDataVector\b", root)]
    candidates: list[list[str]] = []
    table_pattern = re.compile(
        r'<CSequenceChartDataTable id="\d+">(.*?)</CSequenceChartDataTable>',
        flags=re.S,
    )
    for table_match in table_pattern.finditer(root):
        list_matches = list(
            re.finditer(
                r'<m_cscdser length="(\d+)">(.*?)</m_cscdser>',
                table_match.group(1),
                flags=re.S,
            )
        )
        for list_match in list_matches:
            expected_length = int(list_match.group(1))
            series_ids = [
                int(value)
                for value in re.findall(r'<elem idref="(\d+)"/>', list_match.group(2))
            ]
            if expected_length < 2 or len(series_ids) != expected_length:
                continue

            section_start = table_match.start()
            section_end = min(
                (position for position in vector_starts if position > section_start),
                default=len(root),
            )
            section = root[section_start:section_end]

            series_to_varsrc: dict[int, list[tuple[int, int]]] = {}
            for match in re.finditer(
                r'<CSequenceChartDataSeries id="(\d+)">(.*?)</CSequenceChartDataSeries>',
                section,
                flags=re.S,
            ):
                ref = re.search(r'<m_varsrc idref="(\d+)"', match.group(2))
                if ref:
                    series_to_varsrc.setdefault(int(match.group(1)), []).append(
                        (section_start + match.start(), int(ref.group(1)))
                    )

            varsrc_to_name: dict[int, list[tuple[int, str]]] = {}
            for match in re.finditer(
                r'<CVariableSource id="(\d+)">(.*?)</CVariableSource>',
                section,
                flags=re.S,
            ):
                name_match = re.search(
                    r'<m_varval type="5">(.*?)</m_varval>',
                    match.group(2),
                    flags=re.S,
                )
                name = repair_mojibake(html.unescape(name_match.group(1))) if name_match else ""
                varsrc_to_name.setdefault(int(match.group(1)), []).append(
                    (section_start + match.start(), name)
                )

            names: list[str] = []
            valid = True
            for series_id in series_ids:
                series_matches = series_to_varsrc.get(series_id, [])
                if not series_matches:
                    valid = False
                    break
                series_pos, varsrc_id = min(
                    series_matches,
                    key=lambda item: abs(item[0] - section_start),
                )
                name_matches = varsrc_to_name.get(varsrc_id, [])
                if not name_matches:
                    valid = False
                    break
                _, name = min(name_matches, key=lambda item: abs(item[0] - series_pos))
                names.append(clean_text(name))
            if valid and len(names) == expected_length:
                candidates.append(names)
    return candidates


def extract_template_layout(
    template_path: Path,
    requested_chart_name: str | None,
    logger: logging.Logger,
) -> GeoCostTemplateLayout:
    chart_name = resolve_template_chart_name(template_path, requested_chart_name, logger)
    layout_candidates: list[GeoCostTemplateLayout] = []
    for embedding_name, root in iter_thinkcell_roots(template_path):
        root_chart_names = [
            html.unescape(value).strip()
            for value in re.findall(r"<m_strName>(.*?)</m_strName>", root, flags=re.S)
        ]
        if chart_name not in root_chart_names:
            continue
        for names in _extract_table_series_names(root):
            first_blank = next((index for index, value in enumerate(names) if not clean_text(value)), len(names))
            entity_names = [clean_text(value) for value in names[:first_blank]]
            trailing = names[first_blank:]
            if len(entity_names) < 2 or any(not value for value in entity_names):
                continue
            if any(clean_text(value) for value in trailing):
                continue
            layout_candidates.append(
                GeoCostTemplateLayout(
                    chart_name=chart_name,
                    series_row_count=len(names),
                    country_first_row=2,
                    country_last_row=1 + len(entity_names),
                    data_row=2 + len(names),
                    country_names=entity_names,
                )
            )
            logger.debug(
                "Template candidate from %s: %s entities, %s series rows, data row %s.",
                embedding_name,
                len(entity_names),
                len(names),
                2 + len(names),
            )
    if not layout_candidates:
        raise ValueError(
            "Could not extract a Geo Cost entity order from the supplied PowerPoint template. "
            "The template must contain one think-cell chart with a non-empty entity-name prefix "
            "followed by blank spacer series."
        )
    layout = max(layout_candidates, key=lambda item: (item.country_count, item.series_row_count))
    logger.info(
        "Template layout: chart=%s, entities=%s, ChartData rows=1:%s, low-cost row=%s.",
        layout.chart_name,
        layout.country_count,
        layout.data_row,
        layout.data_row,
    )
    logger.debug("First template entities: %s", layout.country_names[:15])
    return layout


# =============================================================================
# Source sheet detection


def read_sheet_values(excel_path: Path, sheet_name: str) -> list[list[Any]]:
    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet {sheet_name!r} not found. Available: {workbook.sheetnames}")
        worksheet = workbook[sheet_name]
        return [list(row) for row in worksheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def detect_source_layout(rows: list[list[Any]]) -> tuple[GeoCostSourceLayout, int] | None:
    max_cols = _max_col_count(rows)
    if not rows or max_cols == 0:
        return None

    marker_candidates: list[tuple[int, int]] = []
    for row_num in range(1, min(len(rows), 20) + 1):
        count = sum(
            1
            for col_num in range(1, max_cols + 1)
            if clean_text(_get(rows, row_num, col_num)) == ">>>"
        )
        if count:
            marker_candidates.append((count, row_num))
    if not marker_candidates:
        return None
    marker_count, marker_row = max(marker_candidates)

    best_header: tuple[int, int, dict[str, int]] | None = None
    for row_num in range(max(1, marker_row), min(len(rows), marker_row + 12) + 1):
        found: dict[str, int] = {}
        for col_num in range(1, max_cols + 1):
            key = compact_text(_get(rows, row_num, col_num))
            if key == "region" and "region" not in found:
                found["region"] = col_num
            elif key in {"highlow", "highorlow", "costtier"} and "high_low" not in found:
                found["high_low"] = col_num
            elif key in {"country", "state", "location", "geography"} and "entity" not in found:
                found["entity"] = col_num
            elif key in {
                "geocostanalysis", "geocostlocation", "displaycountry", "displaystate",
                "countryname", "statename",
            } and "display" not in found:
                found["display"] = col_num
        header_score = sum(1 for key in ("region", "high_low", "entity") if key in found)
        if "display" in found:
            header_score += 1
        candidate = (header_score, -abs(row_num - (marker_row + 3)), found)
        if best_header is None or candidate[:2] > best_header[:2]:
            best_header = (candidate[0], row_num, found)

    if best_header is None or best_header[0] < 3:
        return None
    header_score, header_row, found = best_header
    entity_col = found["entity"]
    display_col = found.get("display", entity_col)
    entity_header = compact_text(_get(rows, header_row, entity_col))
    entity_label = "State" if "state" in entity_header else "Country"
    layout = GeoCostSourceLayout(
        marker_row=marker_row,
        header_row=header_row,
        data_start_row=header_row + 1,
        segment_bucket_col=found["region"],
        high_low_col=found["high_low"],
        raw_country_col=entity_col,
        display_country_col=display_col,
        entity_label=entity_label,
    )
    structure_score = marker_count * 5 + header_score * 5
    return layout, structure_score


def _sheet_name_scores(sheet_name: str) -> tuple[int, int, int]:
    compact = compact_text(sheet_name)
    tokens = set(re.findall(r"[a-z0-9]+", sheet_name.lower()))
    base = 0
    if "geocost" in compact:
        base += 50
    elif "geo" in tokens and "cost" in tokens:
        base += 45
    elif "geo" in compact and "cost" in compact:
        base += 35

    global_score = base
    us_score = base
    if any(token in tokens for token in {"global", "world", "worldwide", "international", "intl"}):
        global_score += 35
        us_score -= 20
    if (
        any(token in tokens for token in {"us", "usa", "domestic"})
        or "unitedstates" in compact
        or "usgeocost" in compact
        or "usageocost" in compact
    ):
        us_score += 40
        global_score -= 20
    return base, global_score, us_score


def _entity_values(rows: list[list[Any]], layout: GeoCostSourceLayout) -> list[str]:
    values: list[str] = []
    blank_streak = 0
    data_end_row = min(
        len(rows),
        int(layout.data_end_row or len(rows)),
    )
    for row_num in range(layout.data_start_row, data_end_row + 1):
        bucket = clean_text(_get(rows, row_num, layout.segment_bucket_col))
        raw = clean_text(_get(rows, row_num, layout.raw_country_col))
        display = clean_text(_get(rows, row_num, layout.display_country_col))
        value = display or raw
        if not value and not bucket:
            blank_streak += 1
            if blank_streak >= 30 and values:
                break
            continue
        blank_streak = 0
        if value and bucket:
            values.append(value)
    return values


def _infer_scope_from_content(rows: list[list[Any]], layout: GeoCostSourceLayout) -> GeoCostScope:
    values = _entity_values(rows, layout)
    keys = [canonical_country_key(value) for value in values if clean_text(value)]
    state_count = sum(1 for key in keys if key in US_STATE_KEYS)
    if len(keys) >= 10 and state_count / max(1, len(keys)) >= 0.6:
        return "us"
    if layout.entity_label.lower() == "state" and state_count >= 5:
        return "us"
    return "global"


def _match_requested_sheet(sheet_names: list[str], requested: str | None) -> str | None:
    requested = clean_text(requested)
    if not requested:
        return None
    exact = next((name for name in sheet_names if name.casefold() == requested.casefold()), None)
    if exact:
        return exact
    key = compact_text(requested)
    matches = [name for name in sheet_names if compact_text(name) == key]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ValueError(f"Geo Cost sheet {requested!r} was not found. Available: {sheet_names}")
    raise ValueError(f"Geo Cost sheet name {requested!r} is ambiguous: {matches}")


def detect_geocost_sheets(
    excel_path: Path,
    global_sheet_override: str | None = None,
    us_sheet_override: str | None = None,
    logger: logging.Logger | None = None,
    input_map_path: Path | str | None = None,
) -> GeoCostWorkbookDetection:
    input_manifest = load_geocost_input_manifest(
        input_map_path,
        excel_path=excel_path,
    )
    manifest_sheet = clean_text((input_manifest or {}).get("sheet_name"))
    manifest_scope = clean_text((input_manifest or {}).get("scope_hint")).lower()

    workbook = load_workbook(excel_path, data_only=True, read_only=True)
    manifest_applied = False
    try:
        sheet_names = list(workbook.sheetnames)
        requested_global = _match_requested_sheet(sheet_names, global_sheet_override)
        requested_us = _match_requested_sheet(sheet_names, us_sheet_override)
        if requested_global and requested_us and requested_global == requested_us:
            raise ValueError("The same worksheet cannot be both Global Geo Cost and US Geo Cost.")
        if manifest_sheet:
            matched_manifest_sheet = _match_requested_sheet(sheet_names, manifest_sheet)
            if matched_manifest_sheet is None:
                raise ValueError(
                    f"Geo Cost input-map worksheet {manifest_sheet!r} was not found."
                )
            manifest_sheet = matched_manifest_sheet

        matches: list[GeoCostSheetMatch] = []
        for worksheet in workbook.worksheets:
            base, global_name_score, us_name_score = _sheet_name_scores(worksheet.title)
            is_manifest_sheet = bool(
                manifest_sheet and worksheet.title.casefold() == manifest_sheet.casefold()
            )
            if (
                base <= 0
                and worksheet.title not in {requested_global, requested_us}
                and not is_manifest_sheet
            ):
                continue

            rows = [list(row) for row in worksheet.iter_rows(values_only=True)]
            if is_manifest_sheet:
                source_layout = source_layout_from_geocost_input_manifest(
                    rows,
                    dict(input_manifest or {}),
                )
                structure_score = 1000
                manifest_applied = True
            else:
                detected = detect_source_layout(rows)
                if detected is None:
                    continue
                source_layout, structure_score = detected

            inferred_scope = _infer_scope_from_content(rows, source_layout)
            if is_manifest_sheet and manifest_scope in {"global", "us"}:
                scope: GeoCostScope = manifest_scope  # type: ignore[assignment]
                name_score = 1000
                inferred = False
            elif worksheet.title == requested_global:
                scope = "global"
                name_score = max(100, global_name_score)
                inferred = False
            elif worksheet.title == requested_us:
                scope = "us"
                name_score = max(100, us_name_score)
                inferred = False
            elif is_manifest_sheet:
                scope = inferred_scope
                name_score = 1000
                inferred = True
            elif global_name_score > us_name_score + 10:
                scope = "global"
                name_score = global_name_score
                inferred = False
            elif us_name_score > global_name_score + 10:
                scope = "us"
                name_score = us_name_score
                inferred = False
            else:
                scope = inferred_scope
                name_score = global_name_score if scope == "global" else us_name_score
                inferred = True

            matches.append(
                GeoCostSheetMatch(
                    scope=scope,
                    sheet_name=worksheet.title,
                    source_layout=source_layout,
                    name_score=name_score,
                    structure_score=structure_score,
                    inferred_from_content=inferred,
                )
            )
    finally:
        workbook.close()

    if input_manifest is not None and not manifest_applied:
        raise ValueError(
            "The Geo Cost input map could not be applied to its selected worksheet."
        )

    def choose(scope: GeoCostScope, explicit_name: str | None) -> GeoCostSheetMatch | None:
        candidates = [item for item in matches if item.scope == scope]
        if explicit_name:
            candidate = next((item for item in candidates if item.sheet_name == explicit_name), None)
            if candidate is None:
                raise ValueError(
                    f"Worksheet {explicit_name!r} does not match the expected Geo Cost structure."
                )
            return candidate
        if not candidates:
            return None
        candidates.sort(key=lambda item: item.total_score, reverse=True)
        if len(candidates) > 1 and candidates[0].total_score == candidates[1].total_score:
            raise ValueError(
                f"Multiple equally likely {scope.upper()} Geo Cost sheets were found: "
                f"{[item.sheet_name for item in candidates if item.total_score == candidates[0].total_score]}. "
                f"Use --{scope}-geocost-sheet to choose one."
            )
        return candidates[0]

    detection = GeoCostWorkbookDetection(
        global_sheet=choose("global", requested_global),
        us_sheet=choose("us", requested_us),
        candidates=matches,
    )
    if logger is not None:
        if input_manifest is not None:
            logger.info(
                "Applied Geo Cost input map %s to worksheet %s.",
                Path(str(input_manifest.get("_manifest_path") or input_map_path)).name,
                manifest_sheet,
            )
        if detection.global_sheet:
            logger.info("Detected Global Geo Cost sheet: %s", detection.global_sheet.sheet_name)
        if detection.us_sheet:
            logger.info("Detected US Geo Cost sheet: %s", detection.us_sheet.sheet_name)
        if not detection.global_sheet and not detection.us_sheet:
            logger.info("No Geo Cost worksheet was detected in %s.", excel_path.name)
    return detection


# =============================================================================
# Function block parsing and ChartData generation


def find_function_blocks(
    rows: list[list[Any]],
    layout: GeoCostSourceLayout,
) -> list[FunctionBlock]:
    blocks: list[FunctionBlock] = []
    marker_start = max(1, int(layout.marker_start_col or 1))
    marker_end = min(
        _max_col_count(rows),
        int(layout.marker_end_col or _max_col_count(rows)),
    )
    for col_num in range(marker_start, marker_end + 1):
        if clean_text(_get(rows, layout.marker_row, col_num)) == ">>>":
            function_name = clean_text(_get(rows, layout.marker_row, col_num + 1))
            if function_name:
                blocks.append(FunctionBlock(function_name, col_num))
    for index, block in enumerate(blocks[:-1]):
        block.next_start_col = blocks[index + 1].start_col
    return blocks


def choose_function_blocks(
    blocks: list[FunctionBlock],
    requested_function: str,
    max_functions: int | None,
) -> list[FunctionBlock]:
    requested_key = compact_text(requested_function)
    if requested_key in {"", "all", "*"}:
        selected = blocks
    else:
        selected = [block for block in blocks if compact_text(block.name) == requested_key]
        if not selected:
            selected = [
                block for block in blocks
                if requested_key and requested_key in compact_text(block.name)
            ][:1]
        if not selected:
            raise ValueError(
                f"Function {requested_function!r} not found. Available: "
                + ", ".join(block.name for block in blocks)
            )
    return selected[:max_functions] if max_functions is not None else selected


def detect_company_columns(
    rows: list[list[Any]],
    block: FunctionBlock,
    layout: GeoCostSourceLayout,
    max_companies: int,
) -> tuple[list[str], list[int]]:
    """Detect every company header in one Geo Cost function block.

    ``max_companies`` is an optional debugging cap. Zero means unlimited. The
    final function block is scanned through the worksheet's actual used columns
    rather than the former arbitrary 30-column window.
    """
    stop_col = (
        block.next_start_col - 1
        if block.next_start_col is not None
        else (_max_col_count(rows) + 1)
    )
    company_limit = max(0, int(max_companies or 0))
    companies: list[str] = []
    company_cols: list[int] = []
    for col_num in range(block.start_col, stop_col):
        name = clean_text(_get(rows, layout.header_row, col_num))
        if not name or name == ">>>":
            continue
        companies.append(name)
        company_cols.append(col_num)
        if company_limit and len(companies) >= company_limit:
            break
    if not companies:
        raise ValueError(f"No company columns found for function {block.name!r}.")
    return companies, company_cols


def country_for_row(
    rows: list[list[Any]],
    row_num: int,
    block: FunctionBlock,
    layout: GeoCostSourceLayout,
) -> str:
    function_country_col = max(1, block.start_col - 1)
    for col_num in (function_country_col, layout.display_country_col, layout.raw_country_col):
        value = clean_text(_get(rows, row_num, col_num))
        if value:
            return value
    return ""


def parse_source_country_rows(
    rows: list[list[Any]],
    block: FunctionBlock,
    company_cols: list[int],
    layout: GeoCostSourceLayout,
    template_layout: GeoCostTemplateLayout,
    logger: logging.Logger,
) -> list[CountryRow]:
    result: list[CountryRow] = []
    blank_streak = 0
    data_end_row = min(
        len(rows),
        int(layout.data_end_row or len(rows)),
    )
    for row_num in range(layout.data_start_row, data_end_row + 1):
        country = country_for_row(rows, row_num, block, layout)
        bucket_raw = clean_text(_get(rows, row_num, layout.segment_bucket_col))
        high_low_raw = clean_text(_get(rows, row_num, layout.high_low_col))
        has_bucket = bool(bucket_raw) and (
            "-" in bucket_raw or high_low_raw.lower() in {"high", "low"}
        )

        if not country and not has_bucket:
            blank_streak += 1
            if blank_streak >= 30 and result:
                break
            continue
        blank_streak = 0
        if not country or not has_bucket:
            continue

        bucket = normalize_segment_label(bucket_raw)
        result.append(
            CountryRow(
                row_num=row_num,
                country=country,
                segment_bucket=bucket,
                high_low=normalize_high_low(high_low_raw, bucket),
                values=[to_number(_get(rows, row_num, col_num)) for col_num in company_cols],
            )
        )

    if len(result) != template_layout.country_count:
        raise ValueError(
            f"Function {block.name!r} on sheet contains {len(result)} {layout.entity_label.lower()} rows; "
            f"the selected PowerPoint template requires exactly {template_layout.country_count}. "
            "No rows were truncated."
        )
    logger.info("Parsed all %s %s rows; none dropped.", len(result), layout.entity_label.lower())
    return result


def reorder_country_rows_to_template(
    source_rows: list[CountryRow],
    template_country_names: list[str],
    logger: logging.Logger,
) -> list[CountryRow]:
    candidates: dict[str, list[CountryRow]] = {}
    for row in source_rows:
        candidates.setdefault(canonical_country_key(row.country), []).append(row)

    ordered: list[CountryRow] = []
    for template_name in template_country_names:
        key = canonical_country_key(template_name)
        matches = candidates.get(key, [])
        if not matches:
            remaining = [row.country for values in candidates.values() for row in values]
            raise ValueError(
                f"Could not map template entity {template_name!r}. Remaining source entities: {remaining}"
            )
        ordered.append(matches.pop(0))

    leftovers = [row.country for values in candidates.values() for row in values]
    if leftovers:
        raise ValueError(f"Unmatched source entities after template mapping: {leftovers}")
    logger.info("Mapped source rows into the PowerPoint template's exact entity order.")
    return ordered


def calculate_totals_and_low_cost(
    country_rows: list[CountryRow],
    company_count: int,
) -> tuple[list[float], list[float | None]]:
    totals = [0.0] * company_count
    low_totals = [0.0] * company_count
    for row in country_rows:
        for company_index, value in enumerate(row.values):
            totals[company_index] += value
            if row.is_low_cost:
                low_totals[company_index] += value
    ratios = [
        low_totals[index] / total if total > 0 else None
        for index, total in enumerate(totals)
    ]
    return totals, ratios


def low_cost_ratios_from_mapped_row(
    rows: list[list[Any]],
    block: FunctionBlock,
    company_cols: list[int],
    layout: GeoCostSourceLayout,
    computed_ratios: list[float | None],
    logger: logging.Logger,
) -> list[float | None]:
    """Use the selected source Low Cost % row when an explicit map is active.

    Legacy Geo Cost runs have no ``layout.low_cost_row`` and continue to use
    ratios calculated from the country rows. For a mapped run, the selected row
    is authoritative for the bottom chart series; calculated values are retained
    only as a per-company fallback when the source cell is blank.
    """
    if not layout.low_cost_row:
        return computed_ratios

    source_values = [
        _optional_ratio(_get(rows, int(layout.low_cost_row), col_num))
        for col_num in company_cols
    ]
    if not any(value is not None for value in source_values):
        logger.warning(
            "Mapped Geo Cost Low Cost %% row %s contains no numeric values for function %s; "
            "using ratios calculated from the geography rows.",
            layout.low_cost_row,
            block.name,
        )
        return computed_ratios

    resolved: list[float | None] = []
    for index, source_value in enumerate(source_values):
        fallback = computed_ratios[index] if index < len(computed_ratios) else None
        resolved.append(source_value if source_value is not None else fallback)
        if source_value is not None and fallback is not None and abs(source_value - fallback) > 0.015:
            logger.warning(
                "Geo Cost Low Cost %% source row differs from the geography-derived ratio for "
                "%s / company column %s: source=%.1f%% calculated=%.1f%%. The mapped source row wins.",
                block.name,
                company_cols[index],
                source_value * 100.0,
                fallback * 100.0,
            )

    logger.info(
        "Using mapped Geo Cost Low Cost %% row %s for function %s: %s",
        layout.low_cost_row,
        block.name,
        [None if value is None else round(value * 100.0, 2) for value in resolved],
    )
    return resolved


def build_label_targets(
    country_rows: list[CountryRow],
    template_country_names: list[str],
    totals: list[float],
    threshold_pct: float,
) -> list[LabelTarget]:
    targets: list[LabelTarget] = []
    for row_index, row in enumerate(country_rows):
        template_country = template_country_names[row_index]
        for company_index, value in enumerate(row.values):
            total = totals[company_index]
            share_pct = value / total * 100.0 if total > 0 else 0.0
            if value > 0 and share_pct + 1e-12 >= threshold_pct:
                targets.append(
                    LabelTarget(
                        company_index=company_index,
                        row_index=row_index,
                        country=row.country,
                        template_country=template_country,
                        display_label=display_label_for_country(template_country),
                        segment_bucket=row.segment_bucket,
                        value=value,
                        total=total,
                        share_pct=share_pct,
                    )
                )
    return targets


def validate_chart_data_workbook(
    path: Path,
    companies: list[str],
    template_layout: GeoCostTemplateLayout,
    totals: list[float],
) -> None:
    workbook = load_workbook(path, data_only=False, read_only=False)
    try:
        chart = workbook["ChartData"]
        actual_names = [
            clean_text(chart.cell(row=row_num, column=1).value)
            for row_num in range(template_layout.country_first_row, template_layout.country_last_row + 1)
        ]
        if actual_names != template_layout.country_names:
            mismatch = next(
                (
                    (index + template_layout.country_first_row, expected, actual)
                    for index, (expected, actual) in enumerate(zip(template_layout.country_names, actual_names))
                    if expected != actual
                ),
                None,
            )
            raise RuntimeError(f"ChartData entity-order validation failed: {mismatch}")
        if any(not name for name in actual_names):
            raise RuntimeError("ChartData does not retain every template entity name in column A.")
        if clean_text(chart.cell(row=template_layout.data_row, column=1).value) != template_layout.low_cost_row_label:
            raise RuntimeError(
                f"ChartData row {template_layout.data_row} is not the expected "
                f"{template_layout.low_cost_row_label!r} row."
            )

        for company_offset, company in enumerate(companies, start=2):
            actual_total = 0.0
            for row_num in range(template_layout.country_first_row, template_layout.country_last_row + 1):
                data_cell = chart.cell(row=row_num, column=company_offset)
                value = data_cell.value
                if data_cell.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
                    raise RuntimeError(
                        f"ChartData contains a formula at {data_cell.coordinate}; plain values are required."
                    )
                if isinstance(value, bool) or not isinstance(value, (int, float)):
                    raise RuntimeError(
                        f"ChartData contains non-numeric data at {data_cell.coordinate}: {value!r}."
                    )
                number_format = clean_text(data_cell.number_format)
                if number_format != PLAIN_VALUE_NUMBER_FORMAT or '"' in number_format or ";" in number_format:
                    raise RuntimeError(
                        f"ChartData has a non-plain number format at {data_cell.coordinate}: {number_format!r}."
                    )
                actual_total += float(value)

            expected_total = totals[company_offset - 2]
            if abs(actual_total - expected_total) > 1e-7:
                raise RuntimeError(
                    f"ChartData total mismatch for {company!r}: expected {expected_total}, found {actual_total}."
                )

            low_cost_cell = chart.cell(row=template_layout.data_row, column=company_offset)
            if low_cost_cell.data_type == "f":
                raise RuntimeError(
                    f"ChartData contains a formula at {low_cost_cell.coordinate}; plain values are required."
                )
            if low_cost_cell.value is not None and not isinstance(low_cost_cell.value, (int, float)):
                raise RuntimeError(
                    f"ChartData contains non-numeric low-cost data at {low_cost_cell.coordinate}: "
                    f"{low_cost_cell.value!r}."
                )
            if clean_text(low_cost_cell.number_format) != PLAIN_PERCENT_NUMBER_FORMAT:
                raise RuntimeError(
                    f"ChartData low-cost format is not plain percent at {low_cost_cell.coordinate}: "
                    f"{low_cost_cell.number_format!r}."
                )
    finally:
        workbook.close()


def write_chart_data_workbook(
    temp_dir: Path,
    scope: GeoCostScope,
    function_name: str,
    companies: list[str],
    country_rows: list[CountryRow],
    template_layout: GeoCostTemplateLayout,
    totals: list[float],
    low_cost_ratio: list[float | None],
    label_targets: list[LabelTarget],
    label_threshold: float,
    logger: logging.Logger,
) -> Path:
    path = temp_dir / f"geocost_{scope}_data_{safe_filename(function_name)}.xlsx"
    workbook = Workbook()
    chart = workbook.active
    chart.title = "ChartData"
    chart.sheet_view.showGridLines = False

    chart.cell(template_layout.header_row, 1).value = ""
    for company_index, company in enumerate(companies, start=2):
        chart.cell(template_layout.header_row, company_index).value = company

    for target_row, (source_row, template_country) in enumerate(
        zip(country_rows, template_layout.country_names),
        start=template_layout.country_first_row,
    ):
        chart.cell(target_row, 1).value = template_country
        for company_index, value in enumerate(source_row.values, start=2):
            chart.cell(target_row, company_index).value = float(value)
            chart.cell(target_row, company_index).number_format = PLAIN_VALUE_NUMBER_FORMAT

    for target_row in range(template_layout.country_last_row + 1, template_layout.data_row):
        chart.cell(target_row, 1).value = None
        for company_index in range(2, len(companies) + 2):
            chart.cell(target_row, company_index).value = None
            chart.cell(target_row, company_index).number_format = "General"

    chart.cell(template_layout.data_row, 1).value = template_layout.low_cost_row_label
    for company_index, ratio in enumerate(low_cost_ratio, start=2):
        chart.cell(template_layout.data_row, company_index).value = None if ratio is None else float(ratio)
        chart.cell(template_layout.data_row, company_index).number_format = PLAIN_PERCENT_NUMBER_FORMAT

    for col_num in range(1, len(companies) + 2):
        chart.column_dimensions[get_column_letter(col_num)].width = 18

    target_keys = {(target.company_index, target.row_index) for target in label_targets}
    audit = workbook.create_sheet("LabelAudit")
    audit.sheet_view.showGridLines = False
    audit.append(
        [
            "Company", "Template row", "Entity in ChartData", "Source entity", "Bucket",
            "FTE", "Company total", "Share", "Display label",
        ]
    )
    for header in audit[1]:
        header.font = Font(bold=True, color="FFFFFF")
        header.fill = PatternFill("solid", fgColor="333333")
        header.alignment = Alignment(horizontal="center")
    for row_index, (source_row, template_country) in enumerate(
        zip(country_rows, template_layout.country_names)
    ):
        for company_index, company in enumerate(companies):
            value = source_row.values[company_index]
            total = totals[company_index]
            share = value / total if total > 0 else 0.0
            audit.append(
                [
                    company,
                    row_index + template_layout.country_first_row,
                    template_country,
                    source_row.country,
                    source_row.segment_bucket,
                    value,
                    total,
                    share,
                    "SHOW" if (company_index, row_index) in target_keys else "HIDE",
                ]
            )
    for row in audit.iter_rows(min_row=2):
        row[5].number_format = "0"
        row[6].number_format = "0"
        row[7].number_format = "0.0%"
    for index, width in enumerate([24, 14, 24, 24, 16, 12, 14, 12, 14], start=1):
        audit.column_dimensions[get_column_letter(index)].width = width
    audit.freeze_panes = "A2"

    reconciliation = workbook.create_sheet("Reconciliation")
    reconciliation.sheet_view.showGridLines = False
    reconciliation.append(["Company", "Source total", "ChartData total", "Difference", "Status"])
    for header in reconciliation[1]:
        header.font = Font(bold=True, color="FFFFFF")
        header.fill = PatternFill("solid", fgColor="333333")
        header.alignment = Alignment(horizontal="center")
    for company_index, company in enumerate(companies):
        chart_total = sum(row.values[company_index] for row in country_rows)
        difference = chart_total - totals[company_index]
        reconciliation.append(
            [company, totals[company_index], chart_total, difference, "PASS" if abs(difference) <= 1e-7 else "FAIL"]
        )
    for row in reconciliation.iter_rows(min_row=2):
        for value_cell in row[1:4]:
            value_cell.number_format = "0.000"
    for index, width in enumerate([28, 16, 16, 14, 12], start=1):
        reconciliation.column_dimensions[get_column_letter(index)].width = width

    workbook.save(path)
    validate_chart_data_workbook(path, companies, template_layout, totals)
    logger.info(
        "Wrote complete %s ChartData: %s entities retained; %s per-bar segments meet the >= %.1f%% rule.",
        scope.upper(),
        len(country_rows),
        len(label_targets),
        label_threshold,
    )
    return path


def parse_geocost_block(
    scope: GeoCostScope,
    sheet_name: str,
    rows: list[list[Any]],
    source_layout: GeoCostSourceLayout,
    block: FunctionBlock,
    max_companies: int,
    label_threshold: float,
    template_layout: GeoCostTemplateLayout,
    temp_dir: Path,
    logger: logging.Logger,
    *,
    company_sort_order: str = "sheet_order",
    company_segment_sort_direction: str = "revenue_descending",
    company_sort_assignments: list[dict[str, Any]] | None = None,
) -> GeoCostData:
    companies, company_cols = detect_company_columns(
        rows, block, source_layout, max_companies
    )
    company_order = build_company_order(
        companies,
        None,
        sort_order=company_sort_order,
        reference_assignments=company_sort_assignments,
        segment_sort_direction=company_segment_sort_direction,
        logger=logger,
        context=f"{scope.upper()} Geo Cost / {block.name}",
    )
    companies = [companies[index] for index in company_order]
    company_cols = [company_cols[index] for index in company_order]
    source_rows = parse_source_country_rows(
        rows, block, company_cols, source_layout, template_layout, logger
    )
    country_rows = reorder_country_rows_to_template(
        source_rows, template_layout.country_names, logger
    )
    totals, computed_low_cost_ratio = calculate_totals_and_low_cost(
        country_rows, len(companies)
    )
    low_cost_ratio = low_cost_ratios_from_mapped_row(
        rows,
        block,
        company_cols,
        source_layout,
        computed_low_cost_ratio,
        logger,
    )
    label_targets = build_label_targets(
        country_rows, template_layout.country_names, totals, label_threshold
    )
    chart_data_path = write_chart_data_workbook(
        temp_dir,
        scope,
        block.name,
        companies,
        country_rows,
        template_layout,
        totals,
        low_cost_ratio,
        label_targets,
        label_threshold,
        logger,
    )

    counts = [
        {
            "company": company,
            "count": sum(1 for target in label_targets if target.company_index == company_index),
        }
        for company_index, company in enumerate(companies)
    ]
    logger.info("Scope: %s | Sheet: %s | Function: %s | Companies: %s", scope, sheet_name, block.name, companies)
    logger.info("Per-bar labels: %s", counts)
    return GeoCostData(
        scope=scope,
        sheet_name=sheet_name,
        function_name=block.name,
        companies=companies,
        country_rows=country_rows,
        template_layout=template_layout,
        totals=totals,
        low_cost_ratio=low_cost_ratio,
        label_targets=label_targets,
        chart_data_path=chart_data_path,
    )
