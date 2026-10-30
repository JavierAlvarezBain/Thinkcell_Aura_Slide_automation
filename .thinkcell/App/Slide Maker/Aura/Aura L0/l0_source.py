from __future__ import annotations

"""Aura L0 source interpretation, quartile calculation, ordering, and color plan.

No L1 quartile or color handler is reused here.  The only shared dependency is
Aura's proven worksheet detector/reader; every L0-specific aggregation,
classification, segment order, palette, and ChartData rule remains in this
module.
"""

from collections import OrderedDict
import math
from pathlib import Path
import re
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from l0_config import (
    BAIN_TOOLBOX_FAMILIES,
    CHART_HEADER_ROW,
    DirectL0QuartileRow,
    FIRST_SEGMENT_ROW,
    GENERIC_L0_ROW_TEXT_KEYS,
    KNOWN_L0_ALIASES,
    L0_COLOR_FAMILY_BY_KEY,
    L0_FIRST_ROW_EXCEPTIONS,
    MAX_FAMILY_COLORED_SEGMENTS,
    OTHER_SEGMENT_PATTERN_BY_L0_KEY,
    PATTERN_SHADE_INDEX_BY_FAMILY,
    REVENUE_CHART_ROW,
    REVENUE_ROW_LABEL,
    SegmentStyle,
    TOTAL_MARKER_PATTERN,
)
from shared import BuildConfig, clean_text, safe_text, to_number_or_none
from fte_quartile_table_detector import (
    choose_aura_summary_sheet,
    compact_text,
    detect_fte_table_layout,
    infer_l0_from_l1,
    load_workbook_fast,
    materialize_values_sheet,
    quartile_columns_for_row,
    quartile_kind,
)

def normalize_l0_display(value: Any) -> str:
    raw = clean_text(value)
    if not raw:
        return ""
    key = compact_text(raw)
    return KNOWN_L0_ALIASES.get(key, raw)


def group_key(value: Any) -> str:
    display = normalize_l0_display(value)
    key = compact_text(display)
    canonical = compact_text(KNOWN_L0_ALIASES.get(key, display))
    # Taxonomy workbooks may use GTM while Aura benchmark rows use S&M.
    # Treat them as the same L0 family without changing the displayed label.
    return {
        "gtm": "sm",
        "gotomarket": "sm",
        "gotomarketcommercial": "sm",
    }.get(canonical, canonical)


def strip_l0_row_descriptors(value: Any) -> str:
    raw = clean_text(value)
    if not raw:
        return ""
    text = re.sub(
        r"\b(total|overall|aggregate|aggregated|benchmark|benchmarks|quartile|quartiles|l0|level\s*0|level\s*zero)\b",
        " ",
        raw,
        flags=re.I,
    )
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def match_text_to_group(value: Any, groups: list[str]) -> str | None:
    raw = clean_text(value)
    if not raw:
        return None

    candidates = {
        compact_text(raw),
        strip_l0_row_descriptors(raw),
        group_key(raw),
    }
    alias_display = KNOWN_L0_ALIASES.get(compact_text(raw))
    if alias_display:
        candidates.add(compact_text(alias_display))

    for group in groups:
        group_candidates = {
            compact_text(group),
            strip_l0_row_descriptors(group),
            group_key(group),
        }
        canonical = KNOWN_L0_ALIASES.get(compact_text(group))
        if canonical:
            group_candidates.add(compact_text(canonical))
        if candidates & group_candidates:
            return group
    return None


def number_or_none(value: Any) -> float | None:
    return to_number_or_none(value)


def percentile_inc(values: Iterable[float], percentile: float) -> float:
    """Excel-compatible PERCENTILE.INC interpolation for 0 <= percentile <= 1."""
    numbers = sorted(float(value) for value in values if value is not None and math.isfinite(float(value)))
    if not numbers:
        raise ValueError("Cannot calculate a percentile from an empty peer set.")
    if len(numbers) == 1:
        return numbers[0]

    p = min(1.0, max(0.0, float(percentile)))
    rank = (len(numbers) - 1) * p
    low = int(math.floor(rank))
    high = int(math.ceil(rank))
    if low == high:
        return numbers[low]
    weight = rank - low
    return numbers[low] + (numbers[high] - numbers[low]) * weight


def sum_segment_values(segments: list[dict[str, Any]], company_count: int) -> list[float | None]:
    totals: list[float | None] = []
    for company_index in range(company_count):
        values: list[float] = []
        for segment in segments:
            segment_values = segment.get("fte_values") or []
            if company_index >= len(segment_values):
                continue
            number = number_or_none(segment_values[company_index])
            if number is not None:
                values.append(number)
        totals.append(sum(values) if values else None)
    return totals


def computed_peer_quartiles(totals: list[float | None]) -> list[float]:
    # The leftmost company is the client. It is intentionally excluded.
    peers = [value for value in totals[1:] if value is not None]
    if not peers:
        raise ValueError("No numeric peer-company L0 totals were available after excluding the client.")
    return [
        percentile_inc(peers, 0.25),
        percentile_inc(peers, 0.50),
        percentile_inc(peers, 0.75),
    ]


def classify_client_quartile(
    client_value: float | None,
    quartiles: list[float],
) -> dict[str, float | str | None]:
    ordered = sorted(float(value) for value in quartiles if value is not None)
    if client_value is None or len(ordered) < 3:
        return {
            "label": "Unknown",
            "client_value": client_value,
            "top_q": None,
            "median": None,
            "bottom_q": None,
        }

    top_q, median, bottom_q = ordered[:3]
    if client_value <= top_q:
        label = "Top"
    elif client_value <= median:
        label = "Second"
    elif client_value <= bottom_q:
        label = "Third"
    else:
        label = "Bottom"

    return {
        "label": label,
        "client_value": client_value,
        "top_q": top_q,
        "median": median,
        "bottom_q": bottom_q,
    }


def record_has_chart_values(record: dict[str, Any]) -> bool:
    return any(number_or_none(value) is not None for value in (record.get("fte_values") or []))


def looks_like_total_segment(function_name: Any) -> bool:
    text = clean_text(function_name)
    if not text:
        return False
    key = compact_text(text)
    return bool(TOTAL_MARKER_PATTERN.search(text)) or key in {"total", "overall", "grandtotal"}


# -----------------------------------------------------------------------------
# L0 grouping and direct L0 quartile-row detection
# -----------------------------------------------------------------------------


def group_l1_records_by_l0(records: list[dict[str, Any]]) -> OrderedDict[str, list[dict[str, Any]]]:
    grouped: OrderedDict[str, list[dict[str, Any]]] = OrderedDict()

    for record in records:
        function_name = clean_text(record.get("function") or record.get("l1"))
        if not function_name or looks_like_total_segment(function_name):
            continue
        if not record_has_chart_values(record):
            continue

        l0_name = normalize_l0_display(record.get("l0") or record.get("level0"))
        if not l0_name:
            l0_name = normalize_l0_display(infer_l0_from_l1(function_name))
        if not l0_name:
            continue

        grouped.setdefault(l0_name, []).append(record)

    return grouped


def row_text_cells(ws, row: int, end_col: int) -> list[tuple[int, str]]:
    result: list[tuple[int, str]] = []
    for col in range(1, max(1, end_col) + 1):
        value = ws.value(row, col)
        text = clean_text(value)
        if not text or number_or_none(value) is not None:
            continue
        result.append((col, text))
    return result


def is_generic_l0_descriptor_text(text: str) -> bool:
    key = compact_text(text)
    if key in GENERIC_L0_ROW_TEXT_KEYS:
        return True
    return bool(TOTAL_MARKER_PATTERN.search(text))


def normalize_direct_quartiles(values: list[float], logger, source: str) -> list[float]:
    ordered = [float(value) for value in values]
    if ordered[0] <= ordered[1] <= ordered[2]:
        return ordered
    sorted_values = sorted(ordered)
    logger.warning(
        "L0 quartile row %s was not ordered Top <= Median <= Bottom. Using sorted values: %s",
        source,
        sorted_values,
    )
    return sorted_values


def _mapped_l0_quartile_rows(
    config: BuildConfig,
    groups: list[str],
    logger,
) -> dict[str, DirectL0QuartileRow] | None:
    """Read the explicit three-column L0 quartile-table mapping when present.

    The selected range may be data-only. In that case the nearest Top/Median/
    Bottom header in the same three columns is resolved above the range. L0
    labels are read from the same rows to the left of the quartile block.
    """
    detected_layout = getattr(config, "detected_source_layout", None)
    inputs = (detected_layout or {}).get("inputs") if isinstance(detected_layout, dict) else None
    mapping = (inputs or {}).get("l0_quartile_table") if isinstance(inputs, dict) else None
    if not isinstance(mapping, dict):
        return None

    sheet_name = clean_text(mapping.get("sheet_name"))
    range_a1 = clean_text(mapping.get("range_a1"))
    if not sheet_name or not range_a1:
        raise ValueError(
            "L0 quartile table detector mapping must contain a worksheet and A1 range."
        )
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    except Exception as exc:
        raise ValueError(
            f"L0 quartile table detector has an invalid range: {range_a1!r}."
        ) from exc
    if max_col - min_col + 1 != 3:
        raise ValueError(
            "L0 quartile table detector must contain exactly three consecutive columns."
        )

    wb = load_workbook_fast(config.excel_path)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f"L0 quartile table detector refers to missing worksheet {sheet_name!r}."
            )
        ws = materialize_values_sheet(wb[sheet_name])

        def header_map(row_number: int) -> dict[str, int]:
            found: dict[str, int] = {}
            for column in range(min_col, max_col + 1):
                kind = quartile_kind(ws.value(row_number, column))
                if kind and kind not in found:
                    found[kind] = column
            return found

        headers = header_map(min_row)
        if {"top", "median", "bottom"}.issubset(headers):
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
                "L0 quartile table detector could not find Top, Median, and Bottom "
                "headers in the selected columns or immediately above them."
            )

        selected: dict[str, DirectL0QuartileRow] = {}
        for row in range(data_start_row, max_row + 1):
            qvalues = [
                number_or_none(ws.value(row, headers[kind]))
                for kind in ("top", "median", "bottom")
            ]
            if any(value is None for value in qvalues):
                continue

            texts = row_text_cells(ws, row, min_col - 1)
            matched_groups: list[str] = []
            for _column, value in texts:
                matched = match_text_to_group(value, groups)
                if matched and matched not in matched_groups:
                    matched_groups.append(matched)
            if len(matched_groups) != 1:
                continue

            group = matched_groups[0]
            source = f"{sheet_name}!{range_a1}, row {row}"
            quartiles = normalize_direct_quartiles(
                [float(value) for value in qvalues if value is not None],
                logger,
                source,
            )
            record = DirectL0QuartileRow(
                l0=group,
                quartiles=quartiles,
                sheet_name=sheet_name,
                row=row,
                header_row=header_row,
                quartile_columns=(
                    int(headers["top"]),
                    int(headers["median"]),
                    int(headers["bottom"]),
                ),
                score=10_000.0,
                inside_main_l1_table=False,
                label_texts=[value for _column, value in texts],
            )
            if group in selected:
                raise ValueError(
                    f"L0 quartile table detector contains more than one benchmark row for {group!r}."
                )
            selected[group] = record

        if not selected:
            raise ValueError(
                "L0 quartile table detector contains no rows that match the detected L0 groups."
            )
        logger.info(
            "Using explicit L0 quartile table mapping %s!%s for groups: %s",
            sheet_name,
            range_a1,
            list(selected),
        )
        return selected
    finally:
        try:
            wb.close()
        except Exception:
            pass


def detect_direct_l0_quartile_rows(
    config: BuildConfig,
    groups: list[str],
    logger,
) -> dict[str, DirectL0QuartileRow]:
    """
    Find dedicated L0 benchmark rows on the detected Aura Summary worksheet.

    Accepted examples include rows labelled G&A, S&M, Ops, R&D, G&A Total,
    Total Ops, and similar. Rows such as "G&A | Finance" are rejected so an L1
    row is not mistaken for an L0 benchmark row.
    """
    if not groups:
        return {}

    mapped_rows = _mapped_l0_quartile_rows(config, groups, logger)
    if mapped_rows is not None:
        return mapped_rows

    wb = load_workbook_fast(config.excel_path)
    try:
        ws = choose_aura_summary_sheet(wb, config, logger)
        layout = detect_fte_table_layout(ws, config, logger)

        header_sets: list[tuple[int, dict[str, int]]] = []
        for row in range(1, ws.max_row + 1):
            found = quartile_columns_for_row(ws, row)
            if {"top", "median", "bottom"}.issubset(found):
                header_sets.append((row, found))

        detected_key = (
            layout.header_row,
            layout.quartile_columns["top"],
            layout.quartile_columns["median"],
            layout.quartile_columns["bottom"],
        )
        if not any(
            (row, found["top"], found["median"], found["bottom"]) == detected_key
            for row, found in header_sets
        ):
            header_sets.append((layout.header_row, dict(layout.quartile_columns)))

        header_sets.sort(key=lambda item: item[0])
        candidates: dict[str, list[DirectL0QuartileRow]] = {group: [] for group in groups}

        for header_index, (header_row, qmap) in enumerate(header_sets):
            qcols = (qmap["top"], qmap["median"], qmap["bottom"])
            q_start = min(qcols)
            next_header_row = (
                header_sets[header_index + 1][0]
                if header_index + 1 < len(header_sets)
                else ws.max_row + 1
            )
            scan_end = min(ws.max_row, next_header_row - 1, header_row + 240)

            for row in range(header_row + 1, scan_end + 1):
                qvalues = [number_or_none(ws.value(row, col)) for col in qcols]
                if any(value is None for value in qvalues):
                    continue

                texts = row_text_cells(ws, row, q_start - 1)
                if not texts:
                    continue

                matched_groups: list[str] = []
                for _, text in texts:
                    matched = match_text_to_group(text, groups)
                    if matched and matched not in matched_groups:
                        matched_groups.append(matched)
                if len(matched_groups) != 1:
                    continue
                group = matched_groups[0]

                function_text = clean_text(ws.value(row, layout.function_col))
                l0_text = clean_text(ws.value(row, layout.l0_col)) if layout.l0_col else ""
                function_group = match_text_to_group(function_text, groups)
                l0_group = match_text_to_group(l0_text, groups)
                total_marker = any(TOTAL_MARKER_PATTERN.search(text) for _, text in texts)

                # Reject a normal L1 row such as A=G&A, B=Finance. Non-group
                # text before the quartile columns must be generic metadata only.
                unrelated_texts = [
                    text
                    for _, text in texts
                    if match_text_to_group(text, groups) is None
                    and not is_generic_l0_descriptor_text(text)
                ]
                if unrelated_texts:
                    continue

                same_qcols_as_main = qcols == (
                    layout.quartile_columns["top"],
                    layout.quartile_columns["median"],
                    layout.quartile_columns["bottom"],
                )
                inside_main = (
                    same_qcols_as_main
                    and layout.fte_start_row <= row <= layout.fte_end_row
                )

                if inside_main and not total_marker:
                    # A plain R&D/R&D row can be a genuine L1 segment. Only a
                    # blank-L0 row labelled R&D in the L1/function column is
                    # accepted as an in-table dedicated L0 row.
                    blank_or_different_l0 = not l0_text or l0_group != group
                    if function_group != group or not blank_or_different_l0:
                        continue

                distance = row - header_row
                score = 0.0
                score += 500.0 if not inside_main else 0.0
                score += 180.0 if total_marker else 0.0
                score += 120.0 if function_group == group else 0.0
                score += 80.0 if l0_group == group else 0.0
                score += 50.0 if len(texts) == 1 else 0.0
                score += max(0.0, 80.0 - distance * 0.5)
                if qvalues[0] <= qvalues[1] <= qvalues[2]:
                    score += 40.0

                source = f"{ws.title}!row {row}"
                quartiles = normalize_direct_quartiles(
                    [float(value) for value in qvalues if value is not None],
                    logger,
                    source,
                )
                candidates[group].append(
                    DirectL0QuartileRow(
                        l0=group,
                        quartiles=quartiles,
                        sheet_name=ws.title,
                        row=row,
                        header_row=header_row,
                        quartile_columns=qcols,
                        score=score,
                        inside_main_l1_table=inside_main,
                        label_texts=[text for _, text in texts],
                    )
                )

        selected: dict[str, DirectL0QuartileRow] = {}
        for group, group_candidates in candidates.items():
            if not group_candidates:
                continue
            group_candidates.sort(key=lambda item: (item.score, -item.row), reverse=True)
            best = group_candidates[0]
            selected[group] = best
            logger.info(
                "Detected dedicated L0 quartile row for %s: %s | values=%s | labels=%s",
                group,
                best.source_description,
                [round(value, 4) for value in best.quartiles],
                best.label_texts,
            )

        return selected
    finally:
        try:
            wb.close()
        except Exception:
            pass


def segment_display_name(segment: dict[str, Any]) -> str:
    return clean_text(segment.get("function") or segment.get("l1"))


def segment_client_sort_value(segment: dict[str, Any]) -> float:
    values = segment.get("fte_values") or []
    if not values:
        return float("-inf")
    value = number_or_none(values[0])
    if value is None:
        return float("-inf")
    return float(value)


def is_first_row_exception_segment(l0_name: str, segment: dict[str, Any]) -> bool:
    l0_key = group_key(l0_name)
    exception_keys = L0_FIRST_ROW_EXCEPTIONS.get(l0_key, set())
    if not exception_keys:
        return False
    return compact_text(segment_display_name(segment)) in exception_keys


def l0_style_key(l0_name: str) -> str:
    """Return the descriptor-insensitive L0 key used for styling.

    Labels such as "Overall Ops", "Ops Overall", and "Overall G&A"
    should retain the same Bain color family as their canonical L0.  Overall
    is metadata about the row, not part of the functional L0 identity.
    """
    raw = clean_text(l0_name)
    if not raw:
        return ""
    style_name = re.sub(r"\boverall\b", " ", raw, flags=re.I)
    style_name = re.sub(r"\s+", " ", style_name).strip()
    return group_key(style_name)


def l0_color_family(l0_name: str) -> str:
    return L0_COLOR_FAMILY_BY_KEY.get(l0_style_key(l0_name), "graphite")


def is_other_exception_for_color(l0_name: str, segment: dict[str, Any]) -> bool:
    l0_key = l0_style_key(l0_name)
    required_key = OTHER_SEGMENT_PATTERN_BY_L0_KEY.get(l0_key)
    if not required_key:
        return False
    return compact_text(segment_display_name(segment)) == required_key


def _normalized_rgb_hex(hex_color: str) -> str:
    color = clean_text(hex_color).lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", color):
        raise ValueError(f"Invalid six-digit RGB hex color: {hex_color!r}")
    return color


def _interpolate_adjacent_hex(
    start_hex: str,
    end_hex: str,
    numerator: int,
    denominator: int,
) -> str:
    """Return the channel-by-channel RGB average between adjacent swatches."""
    if denominator <= 0:
        raise ValueError("Color interpolation denominator must be positive.")
    if numerator < 0 or numerator > denominator:
        raise ValueError("Color interpolation numerator is outside the interval.")

    start = _normalized_rgb_hex(start_hex)
    end = _normalized_rgb_hex(end_hex)
    start_channels = [int(start[index:index + 2], 16) for index in (0, 2, 4)]
    end_channels = [int(end[index:index + 2], 16) for index in (0, 2, 4)]
    result: list[int] = []
    for start_channel, end_channel in zip(start_channels, end_channels):
        weighted = (
            start_channel * (denominator - numerator)
            + end_channel * numerator
        )
        # Round half up.  A 1/2 step is therefore the ordinary RGB average.
        result.append((weighted + denominator // 2) // denominator)
    return "".join(f"{channel:02X}" for channel in result)


def _interpolated_family_scale(
    base_palette: list[str],
    target_count: int,
) -> list[str]:
    """Create a darkest-to-lightest family scale for one through ten rows.

    One through five rows preserve the existing exact Bain Toolbox behavior.
    Six through ten rows are sampled evenly across the five official swatches;
    every non-anchor value is an RGB average of its two adjacent Bain colors.
    """
    if target_count <= 0:
        return []
    palette = [_normalized_rgb_hex(color) for color in base_palette]
    if not palette:
        raise ValueError("Cannot interpolate an empty Bain palette.")
    if target_count <= len(palette):
        return palette[:target_count]
    if target_count > MAX_FAMILY_COLORED_SEGMENTS:
        raise ValueError(
            f"Family scale requested {target_count} colors; maximum is "
            f"{MAX_FAMILY_COLORED_SEGMENTS}."
        )
    if len(palette) == 1:
        return [palette[0]] * target_count

    denominator = target_count - 1
    total_gap_count = len(palette) - 1
    scale: list[str] = []
    for position in range(target_count):
        global_numerator = position * total_gap_count
        left_index, local_numerator = divmod(global_numerator, denominator)
        if left_index >= len(palette) - 1:
            scale.append(palette[-1])
        else:
            scale.append(
                _interpolate_adjacent_hex(
                    palette[left_index],
                    palette[left_index + 1],
                    local_numerator,
                    denominator,
                )
            )
    if len(scale) != target_count:
        raise RuntimeError(
            f"Dynamic Bain scale length mismatch: expected {target_count}, got {len(scale)}."
        )
    return scale


def _solid_palette_with_lightest_family_overflow(
    base_palette: list[str],
    ordinary_segment_count: int,
) -> list[str]:
    """Return darkest-to-lightest colors for all ordinary segments.

    The bottom-most ten ordinary segments receive the dynamic Bain family
    scale.  If there are more than ten, the additional low-value segments at
    the top of the stack repeat the lightest shade of that same ten-color
    family scale.  This preserves the existing darkest-at-bottom visual
    architecture without introducing a different color family.
    """
    if ordinary_segment_count <= 0:
        return []
    family_count = min(ordinary_segment_count, MAX_FAMILY_COLORED_SEGMENTS)
    result = _interpolated_family_scale(base_palette, family_count)
    overflow_count = ordinary_segment_count - family_count
    if overflow_count > 0:
        lightest_family_shade = result[-1]
        result.extend([lightest_family_shade] * overflow_count)
    return result


def segment_style_plan(l0_name: str, segments: list[dict[str, Any]]) -> list[SegmentStyle]:
    family = l0_color_family(l0_name)
    palette = BAIN_TOOLBOX_FAMILIES[family]
    pattern_shade_index = PATTERN_SHADE_INDEX_BY_FAMILY.get(family, 2)

    normal_indices = [
        index for index, segment in enumerate(segments)
        if not is_other_exception_for_color(l0_name, segment)
    ]
    style_by_index: dict[int, SegmentStyle] = {}
    normal_count = len(normal_indices)
    solid_palette = _solid_palette_with_lightest_family_overflow(palette, normal_count)

    for top_to_bottom_index, segment_index in enumerate(normal_indices):
        # ChartData rows are top-to-bottom while the last row renders at the
        # bottom.  Index zero in solid_palette is the darkest family color.
        from_bottom_index = normal_count - 1 - top_to_bottom_index
        hex_color = solid_palette[from_bottom_index]
        style_by_index[segment_index] = SegmentStyle(
            family=family,
            hex_color=hex_color,
            patterned=False,
        )

    for index, segment in enumerate(segments):
        if is_other_exception_for_color(l0_name, segment):
            style_by_index[index] = SegmentStyle(
                family=family,
                hex_color=palette[pattern_shade_index],
                patterned=True,
            )

    return [
        style_by_index.get(
            index,
            SegmentStyle(family=family, hex_color=palette[-1], patterned=False),
        )
        for index in range(len(segments))
    ]

def segment_style_log_items(l0_name: str, segments: list[dict[str, Any]]) -> list[str]:
    styles = segment_style_plan(l0_name, segments)
    result: list[str] = []
    for segment, style in zip(segments, styles):
        suffix = " patterned" if style.patterned else ""
        result.append(
            f"{segment_display_name(segment)}={style.family} #{style.hex_color}{suffix}"
        )
    return result


def order_l0_segments_for_chart(
    l0_name: str,
    segments: list[dict[str, Any]],
    logger,
) -> list[dict[str, Any]]:
    """
    Order L0 stacked-chart rows before writing ChartData.

    Rules:
      1. For G&A, force Other G&A to the first chart row when present.
      2. For Ops, force Other Ops to the first chart row when present.
      3. All remaining rows are sorted ascending by the client value, where
         the client is the first company column detected from Aura Summary.
      4. Ties preserve the source workbook order.
    """
    indexed_segments = list(enumerate(segments))
    exception_items = [
        item for item in indexed_segments if is_first_row_exception_segment(l0_name, item[1])
    ]
    normal_items = [
        item for item in indexed_segments if not is_first_row_exception_segment(l0_name, item[1])
    ]

    exception_items.sort(key=lambda item: item[0])
    normal_items.sort(
        key=lambda item: (
            1 if not math.isfinite(segment_client_sort_value(item[1])) else 0,
            segment_client_sort_value(item[1])
            if math.isfinite(segment_client_sort_value(item[1]))
            else 0.0,
            item[0],
        )
    )

    ordered = [segment for _, segment in exception_items + normal_items]
    original_names = [segment_display_name(segment) for segment in segments]
    ordered_names = [segment_display_name(segment) for segment in ordered]
    ordered_values = [segment_client_sort_value(segment) for segment in ordered]

    if ordered_names != original_names:
        logger.info(
            "L0 segment row order for %s changed from source order %s to chart order %s; "
            "client values=%s; first-row exceptions=%s",
            l0_name,
            original_names,
            ordered_names,
            [None if not math.isfinite(value) else round(value, 4) for value in ordered_values],
            [segment_display_name(segment) for _, segment in exception_items],
        )
    else:
        logger.info(
            "L0 segment row order for %s retained as %s; client values=%s; "
            "first-row exceptions=%s",
            l0_name,
            ordered_names,
            [None if not math.isfinite(value) else round(value, 4) for value in ordered_values],
            [segment_display_name(segment) for _, segment in exception_items],
        )
    return ordered


def build_l0_slide_records(
    company_headers: list[str],
    records: list[dict[str, Any]],
    config: BuildConfig,
    logger,
) -> list[dict[str, Any]]:
    grouped = group_l1_records_by_l0(records)
    if not grouped:
        raise ValueError(
            "No L0 groups could be built. The Aura Summary FTE/$B table must contain an L0 column "
            "or L1 names that can be mapped to L0 groups."
        )

    direct_rows = detect_direct_l0_quartile_rows(config, list(grouped.keys()), logger)
    slide_records: list[dict[str, Any]] = []

    for l0_name, segments in grouped.items():
        direct = direct_rows.get(l0_name)

        # If a dedicated row was located inside the main L1 table, remove the
        # matching all-L0 row from the stack. This does not remove a normal
        # R&D/R&D segment unless that row was explicitly identified as the
        # dedicated L0 benchmark row.
        if direct and direct.inside_main_l1_table:
            filtered: list[dict[str, Any]] = []
            removed = False
            for segment in segments:
                if (
                    not removed
                    and match_text_to_group(segment.get("function"), [l0_name]) == l0_name
                ):
                    removed = True
                    continue
                filtered.append(segment)
            segments = filtered

        if not segments:
            logger.warning("Skipping %s because it has no chart-ready L1 rows.", l0_name)
            continue

        segments = order_l0_segments_for_chart(l0_name, segments, logger)

        totals = sum_segment_values(segments, len(company_headers))
        if direct is not None:
            quartiles = direct.quartiles
            quartile_source = f"dedicated L0 row: {direct.source_description}"
        else:
            quartiles = computed_peer_quartiles(totals)
            quartile_source = "calculated from peer L0 totals (client excluded; PERCENTILE.INC 25/50/75%)"
            logger.info(
                "No dedicated L0 quartile row found for %s. Calculated peer quartiles: %s",
                l0_name,
                [round(value, 4) for value in quartiles],
            )

        client_quartile = classify_client_quartile(totals[0] if totals else None, quartiles)
        slide_records.append(
            {
                "l0": l0_name,
                "function": l0_name,
                "segments": segments,
                "fte_values": totals,
                "quartile_values": quartiles,
                "quartile_source": quartile_source,
                "client_quartile": client_quartile,
            }
        )

    if not slide_records:
        raise ValueError("No chart-ready Aura L0 slide records were created.")
    return slide_records


# -----------------------------------------------------------------------------
# Chart workbook and PowerPoint build
# -----------------------------------------------------------------------------


def write_l0_chart_data_workbook(
    temp_dir: Path,
    company_headers: list[str],
    revenue_values: list[Any],
    slide_record: dict[str, Any],
    config: BuildConfig,
) -> Path:
    l0_name = slide_record["l0"]
    segments = slide_record["segments"]
    temp_dir.mkdir(parents=True, exist_ok=True)
    path = temp_dir / f"l0_chart_data_{safe_text(l0_name)}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "ChartData"

    ws.cell(row=CHART_HEADER_ROW, column=1).value = "Axis"
    for column, company in enumerate(company_headers, start=2):
        ws.cell(row=CHART_HEADER_ROW, column=column).value = company

    ws.cell(row=REVENUE_CHART_ROW, column=1).value = REVENUE_ROW_LABEL
    for column, value in enumerate(revenue_values, start=2):
        ws.cell(row=REVENUE_CHART_ROW, column=column).value = value
        if isinstance(value, (int, float)):
            ws.cell(row=REVENUE_CHART_ROW, column=column).number_format = "$0.00"

    # Rows 3-5 intentionally remain blank. The L0 template expects the first
    # stacked-series row at row 6.
    for row in range(3, FIRST_SEGMENT_ROW):
        for column in range(1, len(company_headers) + 2):
            ws.cell(row=row, column=column).value = None

    styles = segment_style_plan(l0_name, segments)
    for row_offset, segment in enumerate(segments):
        row = FIRST_SEGMENT_ROW + row_offset
        style = styles[row_offset]
        row_fill = style.excel_fill

        label_cell = ws.cell(row=row, column=1)
        label_cell.value = clean_text(segment.get("function") or segment.get("l1"))
        label_cell.fill = row_fill

        values = segment.get("fte_values") or []
        for company_index in range(len(company_headers)):
            cell = ws.cell(row=row, column=company_index + 2)
            value = values[company_index] if company_index < len(values) else None
            cell.value = value
            cell.number_format = "0"
            cell.fill = row_fill

    config.chart_rows = FIRST_SEGMENT_ROW - 1 + len(segments)
    config.chart_cols = len(company_headers) + 1
    for column in range(1, config.chart_cols + 1):
        ws.column_dimensions[get_column_letter(column)].width = 18

    wb.save(path)
    return path




def filter_l0_slide_records(
    slide_records: list[dict[str, Any]],
    only_l0: str | None,
    max_slides: int | None,
) -> list[dict[str, Any]]:
    result = slide_records
    if only_l0:
        wanted = group_key(only_l0)
        result = [record for record in result if group_key(record["l0"]) == wanted]
        if not result:
            available = ", ".join(record["l0"] for record in slide_records)
            raise ValueError(f"No L0 group matched {only_l0!r}. Available groups: {available}")
    if max_slides is not None:
        result = result[: max(0, int(max_slides))]
    if not result:
        raise ValueError("The L0 record filter produced zero slides.")
    return result


