from __future__ import annotations

"""Aura L0 PowerPoint shell, chart-series, and visible-legend helpers."""

from pathlib import Path
import re
import shutil
from typing import Any

from l0_config import (
    L0_LEGEND_BOTTOM_MAX,
    L0_LEGEND_ENTRY_GAP,
    L0_LEGEND_LEFT,
    L0_LEGEND_MARKER_HEIGHT,
    L0_LEGEND_MARKER_TEXT_GAP,
    L0_LEGEND_MARKER_WIDTH,
    L0_LEGEND_NAME_PREFIX,
    L0_LEGEND_RIGHT,
    L0_LEGEND_TOP_MIN,
    MSO_ANCHOR_MIDDLE,
    MSO_PATTERN_WIDE_UPWARD_DIAGONAL,
    MSO_SHAPE_RECTANGLE,
    MSO_TEXT_ORIENTATION_HORIZONTAL,
    PP_ALIGN_LEFT,
    SegmentStyle,
)
from l0_source import segment_display_name, segment_style_plan
from shared import (
    BuildConfig,
    MSO_FALSE,
    MSO_TRUE,
    clean_text,
    pump_office_ui,
    stage,
)
from chart_updater import iter_shapes, update_slide_text
from fte_quartile_table_detector import compact_text

def _shape_text(shape) -> str:
    try:
        if bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText):
            return str(shape.TextFrame.TextRange.Text)
    except Exception:
        pass
    try:
        if bool(shape.TextFrame2.HasText):
            return str(shape.TextFrame2.TextRange.Text)
    except Exception:
        pass
    return ""


def _set_shape_text(shape, text: str) -> bool:
    try:
        shape.TextFrame.TextRange.Text = text
        return True
    except Exception:
        pass
    try:
        shape.TextFrame2.TextRange.Text = text
        return True
    except Exception:
        return False


def _shape_name_key(shape) -> str:
    try:
        raw_name = str(shape.Name or "")
    except Exception:
        raw_name = ""
    return re.sub(r"[^a-z0-9]+", "", raw_name.lower())


def _running_agenda_left_shapes(slide):
    """Return the named Bain L0 agenda text and backing ribbon shapes."""
    text_left = None
    bar_left = None
    for shape in iter_shapes(slide.Shapes):
        name_key = _shape_name_key(shape)
        if "runningagenda" not in name_key:
            continue
        if "textleft" in name_key and text_left is None:
            text_left = shape
        elif "barleft" in name_key and bar_left is None:
            bar_left = shape
    return text_left, bar_left


def _shape_width(shape) -> float | None:
    try:
        width = abs(float(shape.Width))
    except Exception:
        return None
    return width if width > 0 else None


def _text_frame_horizontal_margins(shape) -> float:
    for frame_name in ("TextFrame2", "TextFrame"):
        try:
            frame = getattr(shape, frame_name)
            left = max(0.0, float(frame.MarginLeft))
            right = max(0.0, float(frame.MarginRight))
            return left + right
        except Exception:
            continue
    return 0.0


def _rendered_text_required_width(shape) -> float | None:
    """Measure the current PowerPoint text and include its horizontal margins.

    TextFrame2.TextRange.BoundWidth is PowerPoint's rendered width, so this
    follows the presentation's actual font/size/spacing rather than estimating
    from character count.  The template's existing AutoFit setting is left
    untouched.
    """
    for frame_name in ("TextFrame2", "TextFrame"):
        try:
            frame = getattr(shape, frame_name)
            text_range = frame.TextRange
            bound_width = float(text_range.BoundWidth)
            if bound_width > 0:
                return bound_width + _text_frame_horizontal_margins(shape)
        except Exception:
            continue
    return None


def _expand_l0_running_agenda_ribbon(text_left, bar_left, logger) -> bool:
    """Expand the Bain dark L0 ribbon to fit longer labels, never shrink it."""
    if text_left is None or bar_left is None:
        return False

    original_text_width = _shape_width(text_left)
    original_bar_width = _shape_width(bar_left)
    if original_text_width is None or original_bar_width is None:
        return False

    # Preserve the template's small extra ribbon tail beyond the text box.
    bar_overhang = max(0.0, original_bar_width - original_text_width)

    current_text_width = _shape_width(text_left) or original_text_width
    rendered_required = _rendered_text_required_width(text_left)
    required_text_width = max(
        original_text_width,
        current_text_width,
        rendered_required or 0.0,
    )

    target_bar_width = max(original_bar_width, required_text_width + bar_overhang)
    changed = False

    # Only expand. Short labels therefore retain the exact template geometry.
    if required_text_width > original_text_width + 0.5:
        try:
            text_left.Width = required_text_width
            changed = True
        except Exception as exc:
            logger.debug("Could not expand L0 running-agenda text box: %s", exc)

    if target_bar_width > original_bar_width + 0.5:
        try:
            bar_left.Width = target_bar_width
            changed = True
        except Exception as exc:
            logger.debug("Could not expand L0 running-agenda ribbon bar: %s", exc)

    if changed:
        logger.info(
            "Expanded L0 running-agenda ribbon for %r: text %.1f -> %.1f pt; bar %.1f -> %.1f pt.",
            clean_text(_shape_text(text_left)),
            original_text_width,
            max(original_text_width, _shape_width(text_left) or required_text_width),
            original_bar_width,
            max(original_bar_width, _shape_width(bar_left) or target_bar_width),
        )
    return changed


def update_l0_slide_text_safe(
    pres,
    l0_name: str,
    client_company: str,
    quartile_label: str,
    config: BuildConfig,
    logger,
) -> None:
    """
    Update the L0 title and Bain running-agenda label without touching the
    think-cell legend.

    The Aura L1 helper has a broad fallback that searches for a shape whose
    text is exactly "Finance". In the L0 template that can be the Finance
    legend entry, so using the fallback can rename the legend item to the L0
    name. Disable that fallback here and update only the named Bain agenda
    shape.
    """
    previous_agenda_setting = bool(config.update_agenda_text)
    config.update_agenda_text = False
    try:
        update_slide_text(
            pres,
            l0_name,
            l0_name,
            client_company,
            quartile_label,
            config,
            logger,
        )
    finally:
        config.update_agenda_text = previous_agenda_setting

    updated = 0
    resized = 0
    for slide_index in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(slide_index)
        text_left, bar_left = _running_agenda_left_shapes(slide)
        if text_left is None:
            continue
        if _set_shape_text(text_left, l0_name):
            updated += 1
            if _expand_l0_running_agenda_ribbon(text_left, bar_left, logger):
                resized += 1

    if updated:
        logger.info(
            "Updated %s named Bain running-agenda L0 shape(s) without scanning legend text; "
            "expanded %s ribbon(s) for longer labels.",
            updated,
            resized,
        )
    else:
        logger.warning(
            "No named Bain running-agenda L0 shape was found; title was updated, "
            "and think-cell legend text was deliberately left untouched."
        )


def append_or_create_combined_deck(
    ppt,
    combined_pres,
    combined_working_path: Path,
    slide_path: Path,
    config: BuildConfig,
    logger,
):
    if combined_pres is None:
        with stage(logger, "Create L0 combined deck"):
            shutil.copyfile(slide_path, combined_working_path)
            return ppt.Presentations.Open(
                str(combined_working_path),
                MSO_FALSE,
                MSO_FALSE,
                MSO_TRUE if config.with_window else MSO_FALSE,
            )

    with stage(logger, "Append slide to L0 combined deck"):
        combined_pres.Slides.InsertFromFile(
            str(slide_path),
            combined_pres.Slides.Count,
            1,
            1,
        )
    return combined_pres


# -----------------------------------------------------------------------------
# Live PowerPoint chart-series and visible legend formatting
# -----------------------------------------------------------------------------


def _office_rgb(hex_color: str) -> int:
    """Return the OLE/VBA RGB integer expected by Office COM."""
    color = clean_text(hex_color).lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-F]{6}", color):
        raise ValueError(f"Invalid RGB hex color: {hex_color!r}")
    red = int(color[0:2], 16)
    green = int(color[2:4], 16)
    blue = int(color[4:6], 16)
    return red | (green << 8) | (blue << 16)


def _shape_bounds_com(shape) -> tuple[float, float, float, float]:
    left = float(shape.Left)
    top = float(shape.Top)
    width = abs(float(shape.Width))
    height = abs(float(shape.Height))
    return left, top, width, height


def _find_largest_chart_shape_com(slide):
    candidates: list[tuple[float, Any]] = []
    for index in range(1, int(slide.Shapes.Count) + 1):
        shape = slide.Shapes.Item(index)
        try:
            if not bool(shape.HasChart):
                continue
        except Exception:
            continue
        try:
            _, _, width, height = _shape_bounds_com(shape)
            area = width * height
        except Exception:
            area = 0.0
        candidates.append((area, shape))
    if not candidates:
        raise RuntimeError("No live PowerPoint chart object was found on the Aura L0 slide.")
    return max(candidates, key=lambda item: item[0])[1]


def _series_collection_com(chart):
    try:
        return chart.SeriesCollection()
    except Exception:
        return chart.SeriesCollection


def _points_collection_com(series):
    try:
        return series.Points()
    except Exception:
        return series.Points


def _fill_format_com(target):
    try:
        return target.Format.Fill
    except Exception:
        return target.Fill


def _line_format_com(target):
    try:
        return target.Format.Line
    except Exception:
        return target.Line


def _set_live_fill_com(target, style: SegmentStyle) -> None:
    fill = _fill_format_com(target)
    fill.Visible = MSO_TRUE
    foreground = _office_rgb(style.hex_color)
    if style.patterned:
        patterned = False
        try:
            fill.Patterned(MSO_PATTERN_WIDE_UPWARD_DIAGONAL)
            patterned = True
        except Exception:
            # The top "Other" series is already patterned in the L0 template.
            # On Office builds that reject Patterned() for chart points, keeping
            # that fill type and changing its foreground/background is enough.
            pass
        fill.ForeColor.RGB = foreground
        try:
            fill.BackColor.RGB = _office_rgb("FFFFFF")
        except Exception:
            pass
        if not patterned:
            # Type 2 is msoFillPatterned. Some Office versions do not expose
            # Fill.Type for chart points, so only fail when the property is
            # available and explicitly reports a non-pattern fill.
            try:
                fill_type = int(fill.Type)
            except Exception:
                fill_type = None
            if fill_type is not None and fill_type != 2:
                raise RuntimeError("The requested patterned fill was not available.")
    else:
        fill.Solid()
        fill.ForeColor.RGB = foreground
        try:
            fill.BackColor.RGB = foreground
        except Exception:
            pass
    try:
        fill.Transparency = 0.0
    except Exception:
        pass
    try:
        _line_format_com(target).Visible = MSO_FALSE
    except Exception:
        pass


def _read_live_fill_rgb_com(target) -> int | None:
    try:
        return int(_fill_format_com(target).ForeColor.RGB) & 0xFFFFFF
    except Exception:
        return None


def _series_name_com(series) -> str:
    for getter in (
        lambda: series.Name,
        lambda: series.Formula,
        lambda: series.FormulaFull,
    ):
        try:
            value = clean_text(getter())
        except Exception:
            value = ""
        if value:
            return value
    return ""


def _series_style_mapping_com(
    series_collection,
    segments: list[dict[str, Any]],
    styles: list[SegmentStyle],
) -> list[tuple[int, Any, dict[str, Any], SegmentStyle, str]]:
    series_count = int(series_collection.Count)
    expected_count = len(segments)
    if series_count < expected_count:
        raise RuntimeError(
            f"The live L0 chart exposes {series_count} series; expected at least {expected_count}."
        )

    expected_keys = [compact_text(segment_display_name(segment)) for segment in segments]
    raw_series: list[tuple[int, Any, str, str]] = []
    for series_index in range(1, series_count + 1):
        series = series_collection.Item(series_index)
        name = _series_name_com(series)
        raw_series.append((series_index, series, name, compact_text(name)))

    # First preference: map by the live series names. This avoids relying on
    # chart-stack direction when PowerPoint exposes usable names.
    by_key: dict[str, list[tuple[int, Any, str, str]]] = {}
    for item in raw_series:
        by_key.setdefault(item[3], []).append(item)
    mapping: list[tuple[int, Any, dict[str, Any], SegmentStyle, str]] = []
    used_indices: set[int] = set()
    name_mapping_succeeded = True
    for segment, style, expected_key in zip(segments, styles, expected_keys):
        candidates = [
            item for item in raw_series
            if item[0] not in used_indices
            and expected_key
            and (
                item[3] == expected_key
                or item[3].endswith(expected_key)
                or expected_key in item[3]
            )
        ]
        if not candidates:
            name_mapping_succeeded = False
            break
        chosen = min(candidates, key=lambda item: len(item[3]))
        used_indices.add(chosen[0])
        mapping.append((chosen[0], chosen[1], segment, style, chosen[2]))

    if name_mapping_succeeded and len(mapping) == expected_count:
        return mapping

    # The L0 stacked chart renders the last ChartData row at the bottom. Office
    # exposes stacked series bottom-to-top, so with N rows the first live series
    # corresponds to segment N and the last live series to segment 1.
    if series_count != expected_count:
        names = [item[2] for item in raw_series]
        raise RuntimeError(
            "Could not map L0 series by name and the live series count differs "
            f"from the segment count. series={series_count}, segments={expected_count}, names={names}"
        )
    mapping = []
    for segment_index, (segment, style) in enumerate(zip(segments, styles)):
        series_index = expected_count - segment_index
        item = raw_series[series_index - 1]
        mapping.append((series_index, item[1], segment, style, item[2]))
    return mapping


def _apply_l0_series_colors_to_slide_com(
    slide,
    l0_name: str,
    segments: list[dict[str, Any]],
    logger,
) -> dict[str, Any]:
    chart_shape = _find_largest_chart_shape_com(slide)
    chart = chart_shape.Chart
    series_collection = _series_collection_com(chart)
    styles = segment_style_plan(l0_name, segments)
    mapping = _series_style_mapping_com(series_collection, segments, styles)

    # Let PowerPoint finish any pending refresh before colors are written. No
    # chart refresh is issued after the fill pass, because the final live
    # formatting must be the last chart mutation before the deck is saved.
    try:
        chart.Refresh()
    except Exception:
        pass
    pump_office_ui(0.25)

    applied_points = 0
    mapping_log: list[str] = []
    for series_index, series, segment, style, live_name in mapping:
        _set_live_fill_com(series, style)
        points = _points_collection_com(series)
        point_count = int(points.Count)
        for point_index in range(1, point_count + 1):
            point = points.Item(point_index)
            _set_live_fill_com(point, style)
            applied_points += 1
        mapping_log.append(
            f"series {series_index} ({live_name or '<unnamed>'}) -> "
            f"{segment_display_name(segment)} #{style.hex_color}"
            + (" patterned" if style.patterned else "")
        )

    pump_office_ui(0.35)

    verified_series = 0
    verified_points = 0
    for _, series, segment, style, _ in mapping:
        expected_rgb = _office_rgb(style.hex_color)
        actual_rgb = _read_live_fill_rgb_com(series)
        if actual_rgb != expected_rgb:
            raise RuntimeError(
                f"Live L0 series color verification failed for {segment_display_name(segment)!r}: "
                f"expected #{style.hex_color}, got OLE RGB {actual_rgb!r}."
            )
        verified_series += 1
        points = _points_collection_com(series)
        for point_index in range(1, int(points.Count) + 1):
            point_rgb = _read_live_fill_rgb_com(points.Item(point_index))
            if point_rgb != expected_rgb:
                raise RuntimeError(
                    f"Live L0 point color verification failed for {segment_display_name(segment)!r}, "
                    f"point {point_index}: expected #{style.hex_color}, got {point_rgb!r}."
                )
            verified_points += 1

    logger.info(
        "Applied and verified live L0 series colors: L0=%s; series=%s; points=%s; mapping=%s",
        l0_name,
        verified_series,
        verified_points,
        mapping_log,
    )
    return {
        "verified_series": verified_series,
        "verified_points": verified_points,
        "mapping": mapping_log,
    }


def _legend_candidate_shapes_com(slide) -> tuple[list[Any], list[Any]]:
    markers: list[Any] = []
    texts: list[Any] = []
    for index in range(1, int(slide.Shapes.Count) + 1):
        shape = slide.Shapes.Item(index)
        try:
            left, top, width, height = _shape_bounds_com(shape)
        except Exception:
            continue
        center_x = left + width / 2.0
        center_y = top + height / 2.0
        if not (
            L0_LEGEND_LEFT - 20.0 <= center_x <= L0_LEGEND_RIGHT + 5.0
            and L0_LEGEND_TOP_MIN <= center_y <= L0_LEGEND_BOTTOM_MAX
        ):
            continue
        text = clean_text(_shape_text(shape))
        if text and height <= 26.0 and width <= 230.0:
            texts.append(shape)
        elif not text and 4.0 <= width <= 32.0 and 3.0 <= height <= 22.0:
            markers.append(shape)
    return markers, texts


def _pair_existing_legend_shapes_com(slide) -> list[tuple[Any, Any]]:
    markers, texts = _legend_candidate_shapes_com(slide)
    unused_text_ids = {id(shape) for shape in texts}
    pairs: list[tuple[Any, Any]] = []
    for marker in sorted(markers, key=lambda shape: (_shape_bounds_com(shape)[1], _shape_bounds_com(shape)[0])):
        m_left, m_top, m_width, m_height = _shape_bounds_com(marker)
        m_right = m_left + m_width
        m_center_y = m_top + m_height / 2.0
        candidates: list[tuple[float, Any]] = []
        for text_shape in texts:
            if id(text_shape) not in unused_text_ids:
                continue
            t_left, t_top, _, t_height = _shape_bounds_com(text_shape)
            t_center_y = t_top + t_height / 2.0
            gap = t_left - m_right
            if abs(t_center_y - m_center_y) <= 8.0 and -2.0 <= gap <= 28.0:
                candidates.append((abs(t_center_y - m_center_y) * 10.0 + abs(gap), text_shape))
        if not candidates:
            continue
        text_shape = min(candidates, key=lambda item: item[0])[1]
        unused_text_ids.remove(id(text_shape))
        pairs.append((marker, text_shape))
    return sorted(
        pairs,
        key=lambda pair: (
            round(_shape_bounds_com(pair[0])[1] / 4.0),
            _shape_bounds_com(pair[0])[0],
        ),
    )


def _delete_segment_legend_candidates_com(slide) -> int:
    markers, texts = _legend_candidate_shapes_com(slide)
    candidate_shape_ids: set[int] = set()
    candidate_names: set[str] = set()
    for shape in markers + texts:
        try:
            candidate_shape_ids.add(int(shape.Id))
        except Exception:
            try:
                candidate_names.add(str(shape.Name))
            except Exception:
                pass

    removed = 0
    for index in range(int(slide.Shapes.Count), 0, -1):
        shape = slide.Shapes.Item(index)
        try:
            is_candidate = int(shape.Id) in candidate_shape_ids
        except Exception:
            try:
                is_candidate = str(shape.Name) in candidate_names
            except Exception:
                is_candidate = False
        if not is_candidate:
            continue
        try:
            shape.Delete()
            removed += 1
        except Exception:
            pass
    return removed


def _estimate_legend_text_width(text: str, font_size: float) -> float:
    # Arial's average glyph is about 0.50-0.55 em for these labels.
    return max(10.0, min(205.0, len(text) * font_size * 0.53 + 4.0))


def _legend_layout(
    labels: list[str],
) -> tuple[float, list[tuple[float, float, float]]]:
    max_width = L0_LEGEND_RIGHT - L0_LEGEND_LEFT
    max_rows = 1 if len(labels) <= 6 else (2 if len(labels) <= 9 else 3)
    starting_font = 10.0 if max_rows == 1 else (8.4 if max_rows == 2 else 7.2)

    selected: tuple[float, list[list[tuple[int, float]]]] | None = None
    font_size = starting_font
    while font_size >= 6.2:
        rows: list[list[tuple[int, float]]] = [[]]
        current_width = 0.0
        for index, label in enumerate(labels):
            text_width = _estimate_legend_text_width(label, font_size)
            entry_width = (
                L0_LEGEND_MARKER_WIDTH
                + L0_LEGEND_MARKER_TEXT_GAP
                + text_width
                + L0_LEGEND_ENTRY_GAP
            )
            if rows[-1] and current_width + entry_width > max_width:
                rows.append([])
                current_width = 0.0
            rows[-1].append((index, text_width))
            current_width += entry_width
        if len(rows) <= max_rows:
            selected = (font_size, rows)
            break
        font_size -= 0.4
    if selected is None:
        raise RuntimeError(f"Could not fit {len(labels)} L0 legend entries into the legend region.")

    font_size, rows = selected
    row_count = len(rows)
    if row_count == 1:
        first_top = 118.8
        row_step = 14.0
    elif row_count == 2:
        first_top = 109.0
        row_step = 16.0
    else:
        first_top = 101.5
        row_step = 16.0

    placements: list[tuple[float, float, float] | None] = [None] * len(labels)
    for row_index, row in enumerate(rows):
        row_width = sum(
            L0_LEGEND_MARKER_WIDTH
            + L0_LEGEND_MARKER_TEXT_GAP
            + text_width
            + L0_LEGEND_ENTRY_GAP
            for _, text_width in row
        )
        if row:
            row_width -= L0_LEGEND_ENTRY_GAP
        x = L0_LEGEND_LEFT + max(0.0, (max_width - row_width) / 2.0)
        top = first_top + row_index * row_step
        for item_index, text_width in row:
            placements[item_index] = (x, top, text_width)
            x += (
                L0_LEGEND_MARKER_WIDTH
                + L0_LEGEND_MARKER_TEXT_GAP
                + text_width
                + L0_LEGEND_ENTRY_GAP
            )
    return font_size, [placement for placement in placements if placement is not None]


def _format_legend_text_shape_com(shape, text: str, font_size: float) -> None:
    _set_shape_text(shape, text)
    try:
        shape.Fill.Visible = MSO_FALSE
    except Exception:
        pass
    try:
        shape.Line.Visible = MSO_FALSE
    except Exception:
        pass
    try:
        frame = shape.TextFrame2
        frame.MarginLeft = 0
        frame.MarginRight = 0
        frame.MarginTop = 0
        frame.MarginBottom = 0
        frame.WordWrap = MSO_FALSE
        frame.VerticalAnchor = MSO_ANCHOR_MIDDLE
        text_range = frame.TextRange
        text_range.ParagraphFormat.Alignment = PP_ALIGN_LEFT
        text_range.Font.Name = "Arial"
        text_range.Font.Size = font_size
        text_range.Font.Fill.ForeColor.RGB = _office_rgb("000000")
    except Exception:
        try:
            shape.TextFrame.MarginLeft = 0
            shape.TextFrame.MarginRight = 0
            shape.TextFrame.MarginTop = 0
            shape.TextFrame.MarginBottom = 0
            shape.TextFrame.TextRange.Font.Name = "Arial"
            shape.TextFrame.TextRange.Font.Size = font_size
            shape.TextFrame.TextRange.Font.Color.RGB = _office_rgb("000000")
        except Exception:
            pass


def _create_legend_pair_com(slide, index: int) -> tuple[Any, Any]:
    marker = slide.Shapes.AddShape(
        MSO_SHAPE_RECTANGLE,
        0.0,
        0.0,
        L0_LEGEND_MARKER_WIDTH,
        L0_LEGEND_MARKER_HEIGHT,
    )
    text_shape = slide.Shapes.AddTextbox(
        MSO_TEXT_ORIENTATION_HORIZONTAL,
        0.0,
        0.0,
        40.0,
        12.0,
    )
    try:
        marker.Name = f"{L0_LEGEND_NAME_PREFIX}_MARKER_{index:02d}"
        text_shape.Name = f"{L0_LEGEND_NAME_PREFIX}_TEXT_{index:02d}"
    except Exception:
        pass
    return marker, text_shape


def _apply_visible_l0_legend_com(
    slide,
    l0_name: str,
    segments: list[dict[str, Any]],
    logger,
) -> dict[str, Any]:
    labels = [segment_display_name(segment) for segment in segments]
    styles = segment_style_plan(l0_name, segments)
    pairs = _pair_existing_legend_shapes_com(slide)
    reused_existing = len(pairs) == len(labels)
    removed = 0
    if not reused_existing:
        removed = _delete_segment_legend_candidates_com(slide)
        pairs = [_create_legend_pair_com(slide, index) for index in range(1, len(labels) + 1)]

    font_size, placements = _legend_layout(labels)
    if len(placements) != len(labels):
        raise RuntimeError("Internal L0 legend layout count mismatch.")

    for index, ((marker, text_shape), label, style, placement) in enumerate(
        zip(pairs, labels, styles, placements),
        start=1,
    ):
        x, top, text_width = placement
        marker.Left = x
        marker.Top = top + 1.8
        marker.Width = L0_LEGEND_MARKER_WIDTH
        marker.Height = L0_LEGEND_MARKER_HEIGHT
        _set_live_fill_com(marker, style)
        text_shape.Left = x + L0_LEGEND_MARKER_WIDTH + L0_LEGEND_MARKER_TEXT_GAP
        text_shape.Top = top
        text_shape.Width = text_width
        text_shape.Height = 12.0
        _format_legend_text_shape_com(text_shape, label, font_size)
        try:
            marker.Name = f"{L0_LEGEND_NAME_PREFIX}_MARKER_{index:02d}"
            text_shape.Name = f"{L0_LEGEND_NAME_PREFIX}_TEXT_{index:02d}"
        except Exception:
            pass

    pump_office_ui(0.2)
    verified = 0
    for (marker, text_shape), label, style in zip(pairs, labels, styles):
        actual_text = clean_text(_shape_text(text_shape))
        actual_rgb = _read_live_fill_rgb_com(marker)
        if actual_text != clean_text(label):
            raise RuntimeError(
                f"Visible L0 legend text verification failed: expected {label!r}, got {actual_text!r}."
            )
        if actual_rgb != _office_rgb(style.hex_color):
            raise RuntimeError(
                f"Visible L0 legend color verification failed for {label!r}: "
                f"expected #{style.hex_color}, got {actual_rgb!r}."
            )
        verified += 1

    logger.info(
        "Updated and verified visible L0 legend: L0=%s; entries=%s; reused_existing=%s; "
        "removed_stale_shapes=%s; font_size=%.1f",
        l0_name,
        verified,
        reused_existing,
        removed,
        font_size,
    )
    return {
        "verified_entries": verified,
        "reused_existing": reused_existing,
        "removed_stale_shapes": removed,
        "font_size": font_size,
    }


def apply_live_l0_colors_and_legends_to_presentation_com(
    pres,
    slide_records: list[dict[str, Any]],
    logger,
) -> None:
    if int(pres.Slides.Count) != len(slide_records):
        raise RuntimeError(
            f"Combined L0 deck has {pres.Slides.Count} slides, but {len(slide_records)} records were built."
        )
    total_series = 0
    total_points = 0
    total_legend_entries = 0
    for slide_index, slide_record in enumerate(slide_records, start=1):
        slide = pres.Slides(slide_index)
        l0_name = slide_record["l0"]
        segments = slide_record["segments"]
        color_result = _apply_l0_series_colors_to_slide_com(
            slide,
            l0_name,
            segments,
            logger,
        )
        legend_result = _apply_visible_l0_legend_com(
            slide,
            l0_name,
            segments,
            logger,
        )
        total_series += int(color_result["verified_series"])
        total_points += int(color_result["verified_points"])
        total_legend_entries += int(legend_result["verified_entries"])
        logger.info(
            "Final live L0 color/legend pass: slide=%s; L0=%s; series=%s; points=%s; legend_entries=%s",
            slide_index,
            l0_name,
            color_result["verified_series"],
            color_result["verified_points"],
            legend_result["verified_entries"],
        )
    logger.info(
        "Final live-deck L0 formatting complete: %s series, %s points, %s visible legend entries; "
        "the Geo Cost live PowerPoint chart-object method was used, not OLE color patching.",
        total_series,
        total_points,
        total_legend_entries,
    )


