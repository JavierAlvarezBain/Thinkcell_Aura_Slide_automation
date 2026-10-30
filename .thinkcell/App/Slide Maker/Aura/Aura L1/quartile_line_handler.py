from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
import logging
import math
import re
import sys
import zipfile

from shared import BuildConfig, parse_numeric_text, pump_office_ui
from chart_updater import find_chart_slide_index, iter_shapes

GEOMETRY_CARRYOVER_BUILD_ID = "geometry-active-stream-final-deck-carryover-v2026-09-16-r2-reviewed"

# Template placeholders in the updated think-cell slide.
# The geometry patch maps these source value-line objects to the Excel quartiles.
SOURCE_QUARTILE_PLACEHOLDERS_BY_RANK = {
    "highest": 124,
    "middle": 86,
    "lowest": 68,
}

QUARTILE_AXIS_X_CLUSTER_TOLERANCE = 16.0
QUARTILE_AXIS_R2_MIN = 0.98
QUARTILE_LINE_CLUSTER_TOLERANCE = 4.0


@dataclass
class QuartileTarget:
    rank: str
    value: float
    label: str


@dataclass
class AxisMapping:
    source: str
    a: float
    b: float
    points: list[tuple[float, float]]

    def y_for_value(self, value: float) -> float:
        return self.a + self.b * value


def build_quartile_targets(quartile_values) -> list[QuartileTarget]:
    values = sorted(float(v) for v in quartile_values if v is not None)
    if len(values) < 3:
        return []
    lowest, middle, highest = values[:3]
    return [
        QuartileTarget("highest", highest, str(int(round(highest)))),
        QuartileTarget("middle", middle, str(int(round(middle)))),
        QuartileTarget("lowest", lowest, str(int(round(lowest)))),
    ]


# -----------------------------------------------------------------------------
# Lightweight axis-settle helpers. These are only used after AddRangeData so we
# save/patch the temp PPTX after think-cell has finished resizing the axis.
# -----------------------------------------------------------------------------

def is_group_shape(shape) -> bool:
    try:
        return int(shape.Type) == 6
    except Exception:
        return False


def shape_has_visible_text(shape) -> bool:
    try:
        return bool(shape.HasTextFrame and shape.TextFrame.HasText)
    except Exception:
        return False


def get_shape_bounds(shape) -> dict[str, float]:
    raw_left = float(shape.Left)
    raw_top = float(shape.Top)
    raw_width = float(shape.Width)
    raw_height = float(shape.Height)
    left = min(raw_left, raw_left + raw_width)
    top = min(raw_top, raw_top + raw_height)
    width = abs(raw_width)
    height = abs(raw_height)
    return {
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "right": left + width,
        "bottom": top + height,
        "center_x": left + width / 2,
        "center_y": top + height / 2,
    }


def simple_numeric_text(shape) -> tuple[str, float] | None:
    try:
        raw_text = str(shape.TextFrame.TextRange.Text).strip()
    except Exception:
        return None
    if not re.fullmatch(r"-?\d+(?:\.0+)?", raw_text):
        return None
    value = parse_numeric_text(raw_text)
    if value is None:
        return None
    return raw_text, float(value)


def collect_axis_label_candidates(slide) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for shape in iter_shapes(slide.Shapes):
        try:
            if is_group_shape(shape) or not shape_has_visible_text(shape):
                continue
            parsed = simple_numeric_text(shape)
            if parsed is None:
                continue
            _, value = parsed
            bounds = get_shape_bounds(shape)
            if not (60 <= bounds["center_y"] <= 505):
                continue
            if bounds["width"] > 70 or bounds["height"] > 32:
                continue
            # Axis labels are on the left side of the chart, not the right-side value-line labels.
            if bounds["center_x"] > 260:
                continue
            candidates.append(
                {
                    "value": value,
                    "center_y": bounds["center_y"],
                    "center_x": bounds["center_x"],
                    "right": bounds["right"],
                }
            )
        except Exception:
            pass
    return candidates


def axis_signature_from_candidates(candidates: list[dict[str, Any]]) -> tuple:
    return tuple(
        sorted(
            (
                round(float(c["value"]), 3),
                round(float(c["center_x"]), 1),
                round(float(c["center_y"]), 1),
            )
            for c in candidates
        )
    )


def wait_for_axis_to_settle(slide, config: BuildConfig, logger: logging.Logger, label: str = "axis") -> None:
    timeout = max(float(config.post_chart_update_wait_seconds), 0.3)
    import time

    start = time.perf_counter()
    previous_signature = None
    stable_count = 0
    while time.perf_counter() - start < timeout:
        pump_office_ui(0.18)
        try:
            signature = axis_signature_from_candidates(collect_axis_label_candidates(slide))
        except Exception:
            signature = None
        if signature and signature == previous_signature:
            stable_count += 1
        else:
            stable_count = 0
            previous_signature = signature
        if signature and stable_count >= 2:
            logger.debug("%s settled with axis signature: %s", label, signature)
            return
    logger.debug("%s settle wait completed. Last signature: %s", label, previous_signature)


def settle_after_thinkcell_chart_update(pres, config: BuildConfig, logger: logging.Logger) -> None:
    slide_idx = find_chart_slide_index(pres, config)
    try:
        pres.Windows(1).Activate()
    except Exception:
        pass
    try:
        pres.Application.ActiveWindow.View.GotoSlide(slide_idx)
    except Exception:
        pass
    try:
        pres.Slides(slide_idx).Select()
    except Exception:
        pass
    pump_office_ui(0.2)
    wait_for_axis_to_settle(pres.Slides(slide_idx), config, logger, label="post-think-cell chart update")


# -----------------------------------------------------------------------------
# Saved-PPTX think-cell geometry patch. This is the one path that produced the
# correct line locations after manual think-cell carry-over. It intentionally
# edits the saved temp PPTX and does not touch live PowerPoint shapes.
# -----------------------------------------------------------------------------

def format_thinkcell_float_literal(value: float) -> str:
    return f"{float(value):.20E}"


def as_int_attribute_like(old_text: str, new_value: int) -> str:
    value = int(round(new_value))
    width = len(old_text)
    value_text = str(value)
    if len(value_text) == width:
        return value_text
    if value >= 0 and len(value_text) < width:
        return value_text.rjust(width, "0")
    return old_text


def linear_internal_axis_mapping(points: list[tuple[float, float]]) -> tuple[float, float] | None:
    if len(points) < 2:
        return None
    n = len(points)
    sx = sum(x for x, _ in points)
    sy = sum(y for _, y in points)
    sxx = sum(x * x for x, _ in points)
    sxy = sum(x * y for x, y in points)
    denom = n * sxx - sx * sx
    if abs(denom) < 1e-9:
        return None
    b = (n * sxy - sx * sy) / denom
    a = (sy - b * sx) / n
    if b >= 0:
        return None
    return a, b


def extract_internal_axis_mapping_from_tc_root(root_text: str) -> tuple[float, float, list[tuple[float, float]]] | None:
    label_y_by_id: dict[str, float] = {}
    for match in re.finditer(r'<CDataAxisLabel id="(\d+)">(.*?)</CDataAxisLabel>', root_text, flags=re.S):
        label_id = match.group(1)
        block = match.group(2)
        rect = re.search(
            r'<m_rectPosition[^>]*\btop="([0-9.+\-E]+)"[^>]*\bbottom="([0-9.+\-E]+)"',
            block,
            flags=re.S,
        )
        if not rect:
            continue
        try:
            top = float(rect.group(1))
            bottom = float(rect.group(2))
        except Exception:
            continue
        label_y_by_id[label_id] = (top + bottom) / 2.0

    points: list[tuple[float, float]] = []
    for match in re.finditer(r'<CDataAxisTickmark id="(\d+)">(.*?)</CDataAxisTickmark>', root_text, flags=re.S):
        block = match.group(2)
        label_match = re.search(r'<m_daxislabel idref="(\d+)"/>', block)
        value_match = re.search(r'<m_varval type="1" val="([^"]+)"', block)
        if not label_match or not value_match:
            continue
        label_id = label_match.group(1)
        if label_id not in label_y_by_id:
            continue
        try:
            value = float(value_match.group(1))
        except Exception:
            continue
        if -100000 <= value <= 100000:
            points.append((value, label_y_by_id[label_id]))

    unique: dict[tuple[float, float], tuple[float, float]] = {}
    for value, y in points:
        unique[(round(value, 8), round(y, 4))] = (value, y)
    points = sorted(unique.values(), key=lambda item: item[0])
    mapping = linear_internal_axis_mapping(points)
    if mapping is None:
        return None
    a, b = mapping
    return a, b, points


def get_tc_value_line_infos(root_text: str) -> list[dict[str, Any]]:
    infos: list[dict[str, Any]] = []
    for match in re.finditer(
        r'<CSequenceChartValueIndicatorLine id="(\d+)">(.*?)</CSequenceChartValueIndicatorLine>',
        root_text,
        flags=re.S,
    ):
        line_id = match.group(1)
        block = match.group(2)
        varsrc_match = re.search(r'<m_varsrc idref="(\d+)"/>', block)
        if not varsrc_match:
            continue
        varsrc_id = varsrc_match.group(1)
        varsrc_match_full = re.search(
            rf'(<CVariableSource id="{re.escape(varsrc_id)}">.*?<m_varval type="1" val=")([^"]+)(".*?</CVariableSource>)',
            root_text,
            flags=re.S,
        )
        if not varsrc_match_full:
            continue
        try:
            value = float(varsrc_match_full.group(2))
        except Exception:
            continue

        textvar_id = None
        textvar_match = re.search(
            rf'<CVariableSource id="{re.escape(varsrc_id)}">.*?<m_ctextvar><elem idref="(\d+)"/></m_ctextvar>.*?</CVariableSource>',
            root_text,
            flags=re.S,
        )
        if textvar_match:
            textvar_id = textvar_match.group(1)

        line_shape_ids: list[str] = []
        cppt_block = re.search(r'<m_cpptline[^>]*>(.*?)</m_cpptline>', block, flags=re.S)
        if cppt_block:
            line_shape_ids = re.findall(r'<elem idref="(\d+)"/>', cppt_block.group(1))

        label_rect = re.search(
            r'(<m_scvalindlabel>.*?<m_rectPosition[^>]*\btop=")([0-9]+)("[^>]*\bbottom=")([0-9]+)(".*?</m_scvalindlabel>)',
            block,
            flags=re.S,
        )
        if label_rect:
            top = int(label_rect.group(2))
            bottom = int(label_rect.group(4))
            half_height = max(1, int(round((bottom - top) / 2)))
            label_center_y: float | None = (top + bottom) / 2.0
        else:
            half_height = 48
            label_center_y = None

        infos.append(
            {
                "line_id": line_id,
                "varsrc_id": varsrc_id,
                "textvar_id": textvar_id,
                "var_value": value,
                "var_value_text": varsrc_match_full.group(2),
                "line_shape_ids": line_shape_ids,
                "label_half_height": half_height,
                "label_center_y": label_center_y,
            }
        )
    return infos


def patch_varsrc_value(root_text: str, varsrc_id: str, new_value: float) -> tuple[str, bool]:
    new_value_text = format_thinkcell_float_literal(new_value)
    pattern = re.compile(
        rf'(<CVariableSource id="{re.escape(varsrc_id)}">.*?<m_varval type="1" val=")([^"]+)(".*?</CVariableSource>)',
        flags=re.S,
    )
    match = pattern.search(root_text)
    if not match:
        return root_text, False
    old_value_text = match.group(2)
    if len(old_value_text) != len(new_value_text):
        return root_text, False
    new_block = match.group(1) + new_value_text + match.group(3)
    return root_text[: match.start()] + new_block + root_text[match.end() :], True


def make_same_length_label_format(old_format: str, label: str) -> str:
    if not old_format:
        return old_format
    label = str(label)
    if len(label) > len(old_format):
        return old_format
    return label + ("'" * (len(old_format) - len(label)))


def patch_text_variable_label_cache(root_text: str, textvar_id: str | None, label: str) -> tuple[str, int]:
    if not textvar_id:
        return root_text, 0
    pattern = re.compile(
        rf'(<CTextVariable id="{re.escape(textvar_id)}"><m_bstrFormat>)(.*?)(</m_bstrFormat>)',
        flags=re.S,
    )
    match = pattern.search(root_text)
    if not match:
        return root_text, 0
    old_format = match.group(2)
    new_format = make_same_length_label_format(old_format, label)
    if len(new_format) != len(old_format) or new_format == old_format:
        return root_text, 0
    new_block = match.group(1) + new_format + match.group(3)
    return root_text[: match.start()] + new_block + root_text[match.end() :], 1


def patch_line_shape_y(root_text: str, shape_id: str, new_y: int) -> tuple[str, int]:
    pattern = re.compile(
        rf'(<CPPTLine id="{re.escape(shape_id)}">.*?<m_rectPPTShape[^>]*\btop=")([0-9]+)("[^>]*\bbottom=")([0-9]+)("[^>]*/>.*?</CPPTLine>)',
        flags=re.S,
    )
    match = pattern.search(root_text)
    if not match:
        return root_text, 0
    new_top = as_int_attribute_like(match.group(2), new_y)
    new_bottom = as_int_attribute_like(match.group(4), new_y)
    new_block = match.group(1) + new_top + match.group(3) + new_bottom + match.group(5)
    if len(new_block) != len(match.group(0)):
        return root_text, 0
    return root_text[: match.start()] + new_block + root_text[match.end() :], 1


def patch_value_line_label_block(block: str, new_y: int, half_height: int) -> tuple[str, int]:
    count = 0
    new_top = int(round(new_y - half_height))
    new_bottom = int(round(new_y + half_height))

    def repl_int(match: re.Match) -> str:
        nonlocal count
        top_text = as_int_attribute_like(match.group(2), new_top)
        bottom_text = as_int_attribute_like(match.group(4), new_bottom)
        new_text = match.group(1) + top_text + match.group(3) + bottom_text + match.group(5)
        if len(new_text) == len(match.group(0)):
            count += 1
            return new_text
        return match.group(0)

    block = re.sub(
        r'(<m_rect(?:Position|PPTShape)[^>]*\btop=")([0-9]+)("[^>]*\bbottom=")([0-9]+)("[^>]*/>)',
        repl_int,
        block,
        flags=re.S,
    )

    def repl_float(match: re.Match) -> str:
        nonlocal count
        top_text = format_thinkcell_float_literal(new_top)
        bottom_text = format_thinkcell_float_literal(new_bottom)
        if len(match.group(2)) != len(top_text) or len(match.group(4)) != len(bottom_text):
            return match.group(0)
        new_text = match.group(1) + top_text + match.group(3) + bottom_text + match.group(5)
        if len(new_text) == len(match.group(0)):
            count += 1
            return new_text
        return match.group(0)

    block = re.sub(
        r'(<m_rectPosition[^>]*reqver="36208"[^>]*\btop=")([0-9.+\-E]+)("[^>]*\bbottom=")([0-9.+\-E]+)("[^>]*/>)',
        repl_float,
        block,
        flags=re.S,
    )
    return block, count


def patch_one_value_line_geometry_by_source_label(
    root_text: str,
    source_label: int,
    target_value: float,
    target_label: str,
) -> tuple[str, int, dict[str, Any]]:
    mapping = extract_internal_axis_mapping_from_tc_root(root_text)
    if mapping is None:
        return root_text, 0, {"source_label": source_label, "reason": "axis_mapping_not_found"}
    a, b, axis_points = mapping
    infos = get_tc_value_line_infos(root_text)
    candidates = [info for info in infos if int(round(float(info["var_value"]))) == int(source_label)]
    if not candidates:
        return root_text, 0, {
            "source_label": source_label,
            "reason": "source_value_line_not_found",
            "axis_points": axis_points,
            "values_found": [round(float(i["var_value"]), 6) for i in infos],
        }

    info = candidates[0]
    new_y = int(round(a + b * float(target_value)))
    patches = 0

    root_text, ok = patch_varsrc_value(root_text, str(info["varsrc_id"]), target_value)
    if ok:
        patches += 1

    root_text, count = patch_text_variable_label_cache(root_text, info.get("textvar_id"), target_label)
    patches += count

    for shape_id in info.get("line_shape_ids", []):
        root_text, count = patch_line_shape_y(root_text, str(shape_id), new_y)
        patches += count

    pattern = re.compile(
        rf'<CSequenceChartValueIndicatorLine id="{re.escape(str(info["line_id"]))}">.*?</CSequenceChartValueIndicatorLine>',
        flags=re.S,
    )
    match = pattern.search(root_text)
    if match:
        new_block, count = patch_value_line_label_block(match.group(0), new_y, int(info.get("label_half_height", 48)))
        if len(new_block) == len(match.group(0)):
            root_text = root_text[: match.start()] + new_block + root_text[match.end() :]
            patches += count

    return root_text, patches, {
        "source_label": source_label,
        "target_label": target_label,
        "target_value": target_value,
        "new_y": new_y,
        "axis_points": axis_points,
        "line_id": info["line_id"],
        "varsrc_id": info["varsrc_id"],
        "textvar_id": info.get("textvar_id"),
        "line_shape_ids": info.get("line_shape_ids", []),
        "patches": patches,
    }


def patch_all_quartile_geometry_in_tc_root(root_text: str, quartile_values) -> tuple[str, int, list[dict[str, Any]]]:
    targets = build_quartile_targets(quartile_values)
    if len(targets) < 3:
        return root_text, 0, [{"reason": "not_enough_quartile_targets"}]
    target_by_rank = {target.rank: target for target in targets}
    details: list[dict[str, Any]] = []
    total_patches = 0
    for rank in ["highest", "middle", "lowest"]:
        source_label = SOURCE_QUARTILE_PLACEHOLDERS_BY_RANK[rank]
        target = target_by_rank[rank]
        root_text, patches, info = patch_one_value_line_geometry_by_source_label(
            root_text=root_text,
            source_label=source_label,
            target_value=target.value,
            target_label=target.label,
        )
        total_patches += patches
        info["rank"] = rank
        details.append(info)
    return root_text, total_patches, details



# -----------------------------------------------------------------------------
# Active-stream quartile patching
# -----------------------------------------------------------------------------


def _active_stream_helpers():
    """Load the accepted CFB/slide-part utilities from the Aura L0 module.

    The earlier Aura L1 implementation scanned the complete ``oleObject*.bin``
    byte array. A CFB container can retain stale XML in unused sectors, and a
    live ``think-cellXML`` stream can also be fragmented across non-contiguous
    sectors. The whole-blob regex could therefore patch an obsolete copy while
    manual carry-over continued to read the untouched active stream.

    Aura L0 already contains a verified CFB directory/FAT reader. Reuse only
    those package primitives here; L1 keeps its own template placeholders and
    quartile semantics.
    """
    l1_dir = Path(__file__).resolve().parent
    l0_dir = l1_dir.parent / "Aura L0"
    if l0_dir.is_dir() and str(l0_dir) not in sys.path:
        sys.path.insert(0, str(l0_dir))

    try:
        from l0_native import (  # type: ignore
            _active_thinkcell_root_records_from_ole,
            _ordered_slide_parts,
            _replace_file_with_retry,
            _slide_thinkcell_embedding_parts,
            _wait_for_readable_pptx,
            _zip_info_copy,
        )
    except Exception as exc:  # pragma: no cover - repository-layout failure
        raise RuntimeError(
            "Could not load the active think-cell stream utilities from Aura L0. "
            f"Expected l0_native.py in {l0_dir}."
        ) from exc

    return {
        "active_roots": _active_thinkcell_root_records_from_ole,
        "ordered_slides": _ordered_slide_parts,
        "replace_file": _replace_file_with_retry,
        "slide_embeddings": _slide_thinkcell_embedding_parts,
        "wait_readable": _wait_for_readable_pptx,
        "zip_info_copy": _zip_info_copy,
    }


def _group_value_line_infos(
    infos: list[dict[str, Any]],
) -> dict[tuple[str, str], list[dict[str, Any]]]:
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for info in infos:
        key = (str(info["line_id"]), str(info["varsrc_id"]))
        grouped.setdefault(key, []).append(info)
    return grouped


def _representative_group_value(group: list[dict[str, Any]]) -> float:
    values = sorted(float(item["var_value"]) for item in group)
    return values[len(values) // 2]


def _representative_group_y(group: list[dict[str, Any]]) -> float:
    values = [
        float(item["label_center_y"])
        for item in group
        if item.get("label_center_y") is not None
    ]
    if not values:
        return float("inf")
    return sum(values) / len(values)


def _resolve_quartile_line_groups(
    root_text: str,
    targets: list[QuartileTarget],
) -> tuple[dict[str, list[dict[str, Any]]], dict[str, str]]:
    """Resolve the same three native lines before mutating any of them.

    A normal L1 chart has exactly three distinct value-indicator identities. In
    that case numeric order is definitive and also makes the patch idempotent:
    it works whether the lines still contain 68/86/124 or were already patched
    to the requested values. Fallbacks are retained for roots containing extra
    value-indicator lines.
    """
    infos = get_tc_value_line_infos(root_text)
    grouped = _group_value_line_infos(infos)
    if len(grouped) < 3:
        raise ValueError(
            f"Expected at least three distinct think-cell value lines, found {len(grouped)}."
        )

    target_by_rank = {target.rank: target for target in targets}
    selected: dict[str, list[dict[str, Any]]] = {}
    methods: dict[str, str] = {}

    if len(grouped) == 3:
        ordered = sorted(grouped.values(), key=_representative_group_value)
        selected = {
            "lowest": ordered[0],
            "middle": ordered[1],
            "highest": ordered[2],
        }
        methods = {rank: "three-line-numeric-order" for rank in selected}
        return selected, methods

    unused = set(grouped)
    for rank in ("highest", "middle", "lowest"):
        source_value = float(SOURCE_QUARTILE_PLACEHOLDERS_BY_RANK[rank])
        candidates = [
            key
            for key in unused
            if any(
                math.isclose(
                    float(info["var_value"]),
                    source_value,
                    rel_tol=0.0,
                    abs_tol=0.51,
                )
                for info in grouped[key]
            )
        ]
        method = "template-placeholder"
        if not candidates:
            target_value = float(target_by_rank[rank].value)
            candidates = [
                key
                for key in unused
                if any(
                    math.isclose(
                        float(info["var_value"]),
                        target_value,
                        rel_tol=0.0,
                        abs_tol=1e-7,
                    )
                    for info in grouped[key]
                )
            ]
            method = "already-target-value"
        if candidates:
            key = min(
                candidates,
                key=lambda item: abs(
                    _representative_group_value(grouped[item])
                    - float(SOURCE_QUARTILE_PLACEHOLDERS_BY_RANK[rank])
                ),
            )
            selected[rank] = grouped[key]
            methods[rank] = method
            unused.remove(key)

    unresolved = [rank for rank in ("highest", "middle", "lowest") if rank not in selected]
    if unresolved:
        geometry_groups = sorted(
            ((key, grouped[key]) for key in unused),
            key=lambda item: _representative_group_y(item[1]),
        )
        if len(geometry_groups) >= len(unresolved):
            geometry_by_rank = {
                "highest": geometry_groups[0] if geometry_groups else None,
                "middle": geometry_groups[len(geometry_groups) // 2] if geometry_groups else None,
                "lowest": geometry_groups[-1] if geometry_groups else None,
            }
            for rank in unresolved:
                candidate = geometry_by_rank.get(rank)
                if candidate is None or candidate[0] not in unused:
                    continue
                key, group = candidate
                selected[rank] = group
                methods[rank] = "value-line-geometry-order"
                unused.remove(key)

    missing = [rank for rank in ("highest", "middle", "lowest") if rank not in selected]
    if missing:
        observed = [
            {
                "line_id": key[0],
                "varsrc_id": key[1],
                "values": sorted({float(item["var_value"]) for item in group}),
                "center_y": _representative_group_y(group),
            }
            for key, group in grouped.items()
        ]
        raise ValueError(
            "Could not resolve all Aura L1 quartile lines. "
            f"Missing ranks={missing}; observed={observed}"
        )

    line_ids = [str(selected[rank][0]["line_id"]) for rank in ("highest", "middle", "lowest")]
    if len(set(line_ids)) != 3:
        raise RuntimeError(
            f"Aura L1 quartile resolver selected overlapping line IDs: {line_ids}"
        )
    return selected, methods


def _patch_all_variable_source_copies(
    root_text: str,
    varsrc_id: str,
    target_value: float,
) -> tuple[str, int, int]:
    pattern = re.compile(
        rf'(<CVariableSource id="{re.escape(varsrc_id)}">.*?<m_varval type="1" val=")([^"]+)(".*?</CVariableSource>)',
        flags=re.S,
    )
    target_text = format_thinkcell_float_literal(target_value)
    changed = 0
    matched = 0

    def replace(match: re.Match) -> str:
        nonlocal changed, matched
        matched += 1
        old_text = match.group(2)
        if len(old_text) != len(target_text):
            return match.group(0)
        replacement = match.group(1) + target_text + match.group(3)
        if replacement != match.group(0):
            changed += 1
        return replacement

    return pattern.sub(replace, root_text), changed, matched


def _patch_all_text_variable_copies(
    root_text: str,
    textvar_id: str,
    label: str,
) -> tuple[str, int, int]:
    pattern = re.compile(
        rf'(<CTextVariable id="{re.escape(textvar_id)}"><m_bstrFormat>)(.*?)(</m_bstrFormat>)',
        flags=re.S,
    )
    changed = 0
    matched = 0

    def replace(match: re.Match) -> str:
        nonlocal changed, matched
        matched += 1
        old_format = match.group(2)
        new_format = make_same_length_label_format(old_format, label)
        if len(new_format) != len(old_format):
            return match.group(0)
        replacement = match.group(1) + new_format + match.group(3)
        if replacement != match.group(0):
            changed += 1
        return replacement

    return pattern.sub(replace, root_text), changed, matched


def _patch_all_line_shape_copies(
    root_text: str,
    shape_id: str,
    target_y: int,
) -> tuple[str, int, int]:
    pattern = re.compile(
        rf'(<CPPTLine id="{re.escape(shape_id)}">.*?<m_rectPPTShape[^>]*\btop=")([0-9]+)("[^>]*\bbottom=")([0-9]+)("[^>]*/>.*?</CPPTLine>)',
        flags=re.S,
    )
    changed = 0
    matched = 0

    def replace(match: re.Match) -> str:
        nonlocal changed, matched
        matched += 1
        new_top = as_int_attribute_like(match.group(2), target_y)
        new_bottom = as_int_attribute_like(match.group(4), target_y)
        replacement = (
            match.group(1)
            + new_top
            + match.group(3)
            + new_bottom
            + match.group(5)
        )
        if len(replacement) != len(match.group(0)):
            return match.group(0)
        if replacement != match.group(0):
            changed += 1
        return replacement

    return pattern.sub(replace, root_text), changed, matched


def _patch_all_value_indicator_line_copies(
    root_text: str,
    line_id: str,
    target_y: int,
    half_height: int,
) -> tuple[str, int, int]:
    pattern = re.compile(
        rf'<CSequenceChartValueIndicatorLine id="{re.escape(line_id)}">.*?</CSequenceChartValueIndicatorLine>',
        flags=re.S,
    )
    changed = 0
    matched = 0

    def replace(match: re.Match) -> str:
        nonlocal changed, matched
        matched += 1
        old_block = match.group(0)
        new_block, _patch_count = patch_value_line_label_block(
            old_block,
            target_y,
            half_height,
        )
        if len(new_block) != len(old_block):
            return old_block
        if new_block != old_block:
            changed += 1
        return new_block

    return pattern.sub(replace, root_text), changed, matched


def _collect_line_shape_centers(
    root_text: str,
    shape_ids: Iterable[str],
) -> dict[str, list[float]]:
    wanted = {str(value) for value in shape_ids}
    result: dict[str, list[float]] = {value: [] for value in wanted}
    for match in re.finditer(
        r'<CPPTLine id="(\d+)">.*?<m_rectPPTShape[^>]*\btop="([0-9]+)"[^>]*\bbottom="([0-9]+)"[^>]*/>.*?</CPPTLine>',
        root_text,
        flags=re.S,
    ):
        shape_id = match.group(1)
        if shape_id in wanted:
            result[shape_id].append((int(match.group(2)) + int(match.group(3))) / 2.0)
    return result


def _collect_value_label_centers(
    root_text: str,
    line_ids: Iterable[str],
) -> dict[str, list[float]]:
    wanted = {str(value) for value in line_ids}
    result: dict[str, list[float]] = {value: [] for value in wanted}
    for match in re.finditer(
        r'<CSequenceChartValueIndicatorLine id="(\d+)">(.*?)</CSequenceChartValueIndicatorLine>',
        root_text,
        flags=re.S,
    ):
        line_id = match.group(1)
        if line_id not in wanted:
            continue
        block = re.sub(
            r'<m_pptgenlineConnector>.*?</m_pptgenlineConnector>',
            '',
            match.group(2),
            flags=re.S,
        )
        for rectangle in re.finditer(
            r'<m_rect(?:Position|PPTShape)[^>]*\btop="([0-9.+\-E]+)"[^>]*\bbottom="([0-9.+\-E]+)"',
            block,
            flags=re.S,
        ):
            result[line_id].append(
                (float(rectangle.group(1)) + float(rectangle.group(2))) / 2.0
            )
    return result


def _verify_rank_geometry(
    root_text: str,
    *,
    rank: str,
    target_y: int,
    line_ids: list[str],
    line_shape_ids: list[str],
) -> dict[str, Any]:
    line_centers = _collect_line_shape_centers(root_text, line_shape_ids)
    label_centers = _collect_value_label_centers(root_text, line_ids)

    bad_lines = {
        shape_id: centers
        for shape_id, centers in line_centers.items()
        if not centers or any(abs(center - target_y) > 1.0 for center in centers)
    }
    if bad_lines:
        raise RuntimeError(
            f"Aura L1 {rank} line geometry did not verify at y={target_y}: {bad_lines}"
        )

    bad_labels = {
        line_id: centers
        for line_id, centers in label_centers.items()
        if not centers or any(abs(center - target_y) > 1.0 for center in centers)
    }
    if bad_labels:
        raise RuntimeError(
            f"Aura L1 {rank} label geometry did not verify at y={target_y}: {bad_labels}"
        )

    return {
        "line_shape_centers": line_centers,
        "label_centers": label_centers,
    }


def _patch_quartiles_in_active_root(
    root_text: str,
    quartile_values,
) -> tuple[str, int, list[dict[str, Any]]]:
    targets = build_quartile_targets(quartile_values)
    if len(targets) < 3:
        raise ValueError("Fewer than three usable Aura L1 quartile values were supplied.")

    axis_mapping = extract_internal_axis_mapping_from_tc_root(root_text)
    if axis_mapping is None:
        raise ValueError("Could not derive the active think-cell axis mapping.")
    axis_a, axis_b, axis_points = axis_mapping

    selected, selection_methods = _resolve_quartile_line_groups(root_text, targets)
    target_by_rank = {target.rank: target for target in targets}
    details: list[dict[str, Any]] = []
    total_changed = 0
    target_y_by_rank: dict[str, int] = {}

    for rank in ("highest", "middle", "lowest"):
        group = selected[rank]
        target = target_by_rank[rank]
        varsrc_ids = sorted({str(item["varsrc_id"]) for item in group})
        textvar_ids = sorted(
            {str(item["textvar_id"]) for item in group if item.get("textvar_id")}
        )
        line_ids = sorted({str(item["line_id"]) for item in group})
        line_shape_ids = sorted(
            {
                str(shape_id)
                for item in group
                for shape_id in item.get("line_shape_ids", [])
            }
        )
        half_height = max(int(item.get("label_half_height", 48)) for item in group)
        target_y = int(round(axis_a + axis_b * float(target.value)))
        target_y_by_rank[rank] = target_y

        rank_changed = 0
        matched_components = 0
        for varsrc_id in varsrc_ids:
            root_text, changed, matched = _patch_all_variable_source_copies(
                root_text, varsrc_id, target.value
            )
            rank_changed += changed
            matched_components += matched
        for textvar_id in textvar_ids:
            root_text, changed, matched = _patch_all_text_variable_copies(
                root_text, textvar_id, target.label
            )
            rank_changed += changed
            matched_components += matched
        for shape_id in line_shape_ids:
            root_text, changed, matched = _patch_all_line_shape_copies(
                root_text, shape_id, target_y
            )
            rank_changed += changed
            matched_components += matched
        for line_id in line_ids:
            root_text, changed, matched = _patch_all_value_indicator_line_copies(
                root_text, line_id, target_y, half_height
            )
            rank_changed += changed
            matched_components += matched

        updated_infos = get_tc_value_line_infos(root_text)
        matching_infos = [
            item
            for item in updated_infos
            if str(item["line_id"]) in line_ids
            and str(item["varsrc_id"]) in varsrc_ids
        ]
        if not matching_infos or any(
            not math.isclose(
                float(item["var_value"]),
                float(target.value),
                rel_tol=0.0,
                abs_tol=1e-7,
            )
            for item in matching_infos
        ):
            raise RuntimeError(
                f"Aura L1 {rank} native value did not verify at {target.value}. "
                f"Observed={[item.get('var_value') for item in matching_infos]}"
            )

        geometry = _verify_rank_geometry(
            root_text,
            rank=rank,
            target_y=target_y,
            line_ids=line_ids,
            line_shape_ids=line_shape_ids,
        )
        details.append(
            {
                "rank": rank,
                "selection_method": selection_methods[rank],
                "source_values": sorted({float(item["var_value"]) for item in group}),
                "target_value": float(target.value),
                "target_label": target.label,
                "target_y": target_y,
                "line_ids": line_ids,
                "varsrc_ids": varsrc_ids,
                "line_shape_ids": line_shape_ids,
                "serialized_copies": len(group),
                "verified_copies": len(matching_infos),
                "matched_components": matched_components,
                "changed_components": rank_changed,
                "axis_points": axis_points,
                **geometry,
            }
        )
        total_changed += rank_changed

    distinct_values = {
        round(float(target_by_rank[rank].value), 9)
        for rank in ("highest", "middle", "lowest")
    }
    distinct_y = {target_y_by_rank[rank] for rank in target_y_by_rank}
    if len(distinct_values) == 3 and len(distinct_y) != 3:
        raise RuntimeError(
            "Distinct Aura L1 quartile values mapped to overlapping line coordinates: "
            f"{target_y_by_rank}"
        )

    return root_text, total_changed, details


def _root_verifies_targets(root_text: str, details: list[dict[str, Any]]) -> bool:
    infos = get_tc_value_line_infos(root_text)
    for item in details:
        line_ids = {str(value) for value in item.get("line_ids", [])}
        varsrc_ids = {str(value) for value in item.get("varsrc_ids", [])}
        target = float(item["target_value"])
        matching = [
            info
            for info in infos
            if str(info["line_id"]) in line_ids
            and str(info["varsrc_id"]) in varsrc_ids
        ]
        if not matching or any(
            not math.isclose(float(info["var_value"]), target, rel_tol=0.0, abs_tol=1e-7)
            for info in matching
        ):
            return False
        _verify_rank_geometry(
            root_text,
            rank=str(item["rank"]),
            target_y=int(item["target_y"]),
            line_ids=sorted(line_ids),
            line_shape_ids=[str(value) for value in item.get("line_shape_ids", [])],
        )
    return True


def _find_exact_active_root_record(
    records: list[dict[str, Any]],
    *,
    stream_name: str,
    root_start: int,
) -> dict[str, Any] | None:
    """Return the same active CFB root that was patched before serialization."""
    for record in records:
        entry = record.get("entry") or {}
        if str(entry.get("name") or "") != str(stream_name or ""):
            continue
        if int(record.get("root_start", -1)) != int(root_start):
            continue
        return record
    return None


def patch_quartile_geometry_in_ole_bytes(
    data: bytes,
    quartile_values,
) -> tuple[bytes, int, list[dict[str, Any]]]:
    """Patch only the live CFB ``think-cellXML`` stream and verify it.

    The returned change count may be zero when an already-correct final deck is
    checked again; successful active-stream verification remains authoritative.
    """
    helpers = _active_stream_helpers()
    compound, root_records = helpers["active_roots"](data)
    if not root_records:
        raise ValueError("No active think-cell chart XML stream was found in the OLE object.")

    stream_buffers: dict[int, bytearray] = {}
    entry_by_id: dict[int, dict[str, Any]] = {}
    patched_details: list[dict[str, Any]] = []
    total_changed = 0

    for root_index, record in enumerate(root_records):
        try:
            new_root_text, changed, details = _patch_quartiles_in_active_root(
                str(record["root_text"]),
                quartile_values,
            )
        except (ValueError, RuntimeError) as exc:
            patched_details.append(
                {
                    "root_index": root_index,
                    "stream_name": str(record["entry"].get("name") or ""),
                    "skipped": str(exc),
                }
            )
            continue

        encoding = str(record.get("encoding") or "latin-1")
        new_root_bytes = new_root_text.encode(encoding)
        old_length = int(record["root_end"]) - int(record["root_start"])
        if len(new_root_bytes) != old_length:
            raise RuntimeError(
                "Aura L1 quartile mutation changed the active root byte length: "
                f"old={old_length}, new={len(new_root_bytes)}."
            )

        entry = record["entry"]
        entry_id = int(entry["id"])
        buffer = stream_buffers.setdefault(
            entry_id,
            bytearray(record["stream_bytes"]),
        )
        entry_by_id[entry_id] = entry
        buffer[int(record["root_start"]) : int(record["root_end"])] = new_root_bytes
        total_changed += changed
        patched_details.append(
            {
                "root_index": root_index,
                "stream_name": str(entry.get("name") or ""),
                "stream_entry_id": entry_id,
                "root_start": int(record["root_start"]),
                "changed_components": changed,
                "items": details,
            }
        )

    successful = [item for item in patched_details if "items" in item]
    if not successful:
        raise RuntimeError(
            "No active think-cell root accepted the Aura L1 quartile patch. "
            f"Diagnostics={patched_details}"
        )

    for entry_id, buffer in stream_buffers.items():
        compound.write_stream(entry_by_id[entry_id], bytes(buffer))
    patched_data = compound.to_bytes()

    _verified_compound, verified_records = helpers["active_roots"](patched_data)
    for patched_item in successful:
        verified_record = _find_exact_active_root_record(
            verified_records,
            stream_name=str(patched_item.get("stream_name") or ""),
            root_start=int(patched_item.get("root_start", -1)),
        )
        if verified_record is None or not _root_verifies_targets(
            str(verified_record["root_text"]),
            patched_item["items"],
        ):
            raise RuntimeError(
                "The Aura L1 quartile patch did not survive reopening the exact "
                "active think-cellXML root that was mutated. "
                f"Details={patched_item}"
            )
        patched_item["active_stream_verified"] = True

    return patched_data, total_changed, patched_details


def _normalize_slide_indices(
    total_slide_count: int,
    record_count: int,
    slide_indices: Iterable[int] | None,
) -> list[int]:
    if slide_indices is None:
        if total_slide_count != record_count:
            raise RuntimeError(
                f"Aura L1 deck contains {total_slide_count} slides but {record_count} "
                "quartile records were supplied. Pass explicit final-deck slide indices."
            )
        return list(range(1, record_count + 1))

    result = [int(value) for value in slide_indices]
    if len(result) != record_count:
        raise ValueError(
            f"Received {len(result)} Aura L1 slide indices for {record_count} records."
        )
    if len(set(result)) != len(result):
        raise ValueError(f"Aura L1 slide indices must be unique: {result}")
    invalid = [value for value in result if value < 1 or value > total_slide_count]
    if invalid:
        raise IndexError(
            f"Aura L1 slide indices {invalid} are outside the deck's 1..{total_slide_count} range."
        )
    return result


def patch_l1_quartiles_in_deck(
    output_ppt: Path,
    slide_records: list[dict[str, Any]],
    logger: logging.Logger,
    *,
    slide_indices: Iterable[int] | None = None,
) -> Path:
    """Patch and verify Aura L1 quartile lines in a closed PPTX package.

    This must be the final mutation before manual carry-over. It targets each
    designated slide's own active think-cell OLE relationship, rewrites only the
    active CFB stream, replaces the package atomically, and performs no later
    PowerPoint save.
    """
    if not slide_records:
        raise ValueError("No Aura L1 slide records were supplied for quartile patching.")

    output_ppt = Path(output_ppt).resolve()
    helpers = _active_stream_helpers()
    helpers["wait_readable"](output_ppt, logger)
    temp_path = output_ppt.with_suffix(output_ppt.suffix + ".l1_quartile_patching")
    temp_path.unlink(missing_ok=True)

    patch_plan: dict[str, tuple[int, dict[str, Any]]] = {}

    try:
        with zipfile.ZipFile(output_ppt, "r") as source_archive:
            slide_parts = helpers["ordered_slides"](source_archive)
            target_indices = _normalize_slide_indices(
                len(slide_parts),
                len(slide_records),
                slide_indices,
            )
            for slide_index, record in zip(target_indices, slide_records):
                quartiles = list(record.get("quartile_values") or [])
                if len([value for value in quartiles if value is not None]) < 3:
                    raise ValueError(
                        f"Aura L1 slide {slide_index} ({record.get('function')}) has fewer "
                        "than three quartile values."
                    )
                embeddings = helpers["slide_embeddings"](
                    source_archive,
                    slide_parts[slide_index - 1],
                )
                if not embeddings:
                    raise RuntimeError(
                        f"Aura L1 slide {slide_index} ({record.get('function')}) has no "
                        "active think-cell chart embedding."
                    )
                for embedding in embeddings:
                    existing = patch_plan.get(embedding)
                    if existing is not None and existing[0] != slide_index:
                        raise RuntimeError(
                            f"think-cell embedding {embedding} is shared by Aura L1 slides "
                            f"{existing[0]} and {slide_index}; per-slide quartiles would conflict."
                        )
                    patch_plan[embedding] = (slide_index, record)

            patched_by_slide: dict[int, list[str]] = {index: [] for index in target_indices}
            verification_plan: dict[tuple[int, str], list[dict[str, Any]]] = {}
            logs: list[dict[str, Any]] = []
            with zipfile.ZipFile(temp_path, "w") as target_archive:
                for info in source_archive.infolist():
                    data = source_archive.read(info.filename)
                    plan = patch_plan.get(info.filename)
                    if plan is not None:
                        slide_index, record = plan
                        try:
                            data, changed, details = patch_quartile_geometry_in_ole_bytes(
                                data,
                                record["quartile_values"],
                            )
                        except Exception as exc:
                            logs.append(
                                {
                                    "slide": slide_index,
                                    "function": record.get("function"),
                                    "embedding": info.filename,
                                    "error": repr(exc),
                                }
                            )
                        else:
                            successful_roots = [
                                item for item in details if isinstance(item, dict) and "items" in item
                            ]
                            if not successful_roots:
                                logs.append(
                                    {
                                        "slide": slide_index,
                                        "function": record.get("function"),
                                        "embedding": info.filename,
                                        "error": "No active root reported a verified quartile patch.",
                                        "details": details,
                                    }
                                )
                            else:
                                patched_by_slide[slide_index].append(info.filename)
                                verification_plan[(slide_index, info.filename)] = successful_roots
                                logs.append(
                                    {
                                        "slide": slide_index,
                                        "function": record.get("function"),
                                        "embedding": info.filename,
                                        "changed_components": changed,
                                        "details": details,
                                    }
                                )
                    target_archive.writestr(helpers["zip_info_copy"](info), data)

        missing_slides = [
            slide_index
            for slide_index, embeddings in patched_by_slide.items()
            if not embeddings
        ]
        if missing_slides:
            raise RuntimeError(
                "No active Aura L1 quartile stream was patched for slide(s) "
                f"{missing_slides}. Diagnostics={logs}"
            )

        helpers["replace_file"](temp_path, output_ppt)
        helpers["wait_readable"](output_ppt, logger)

        # Reopen the completed package and verify the exact active CFB root(s)
        # that were patched for every target slide. This checks native values,
        # line geometry, and label geometry; a stale/slack-sector write cannot
        # satisfy this verification.
        with zipfile.ZipFile(output_ppt, "r") as verified_archive:
            verified_slide_parts = helpers["ordered_slides"](verified_archive)
            for slide_index, record in zip(target_indices, slide_records):
                related_embeddings = set(
                    helpers["slide_embeddings"](
                        verified_archive,
                        verified_slide_parts[slide_index - 1],
                    )
                )
                planned_embeddings = [
                    embedding
                    for (planned_slide, embedding) in verification_plan
                    if planned_slide == slide_index
                ]
                verified_any_embedding = False

                for embedding in planned_embeddings:
                    if embedding not in related_embeddings:
                        raise RuntimeError(
                            f"Aura L1 slide {slide_index} ({record.get('function')}) no longer "
                            f"references patched embedding {embedding}."
                        )
                    data = verified_archive.read(embedding)
                    _compound, root_records = helpers["active_roots"](data)
                    root_plans = verification_plan[(slide_index, embedding)]
                    for root_plan in root_plans:
                        verified_record = _find_exact_active_root_record(
                            root_records,
                            stream_name=str(root_plan.get("stream_name") or ""),
                            root_start=int(root_plan.get("root_start", -1)),
                        )
                        if verified_record is None:
                            raise RuntimeError(
                                f"Final Aura L1 verification could not find the patched active root "
                                f"for slide {slide_index} ({record.get('function')}), embedding "
                                f"{embedding}, stream={root_plan.get('stream_name')!r}, "
                                f"root_start={root_plan.get('root_start')!r}."
                            )
                        if not _root_verifies_targets(
                            str(verified_record["root_text"]),
                            root_plan["items"],
                        ):
                            raise RuntimeError(
                                f"Final active-stream value/geometry verification failed for "
                                f"Aura L1 slide {slide_index} ({record.get('function')}), "
                                f"embedding {embedding}."
                            )
                    verified_any_embedding = True

                if not verified_any_embedding:
                    raise RuntimeError(
                        f"Final active-stream verification found no successfully patched "
                        f"embedding for Aura L1 slide {slide_index} ({record.get('function')})."
                    )

        logger.info(
            "Verified active-stream Aura L1 quartile carry-over data on slide(s) %s. "
            "No PowerPoint save occurred after the patch.",
            target_indices,
        )
        logger.debug("Aura L1 active-stream quartile patch details: %s", logs)
        return output_ppt
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def patch_quartile_geometry_in_pptx(
    pptx_path: Path,
    quartile_values,
    logger: logging.Logger,
) -> bool:
    """Backward-compatible one-slide wrapper using the active stream."""
    patch_l1_quartiles_in_deck(
        output_ppt=pptx_path,
        slide_records=[
            {
                "function": pptx_path.stem,
                "quartile_values": list(quartile_values),
            }
        ],
        logger=logger,
        slide_indices=[1],
    )
    return True


def prepare_quartile_lines_for_deferred_carryover(
    temp_ppt: Path,
    quartile_values,
    logger: logging.Logger,
    function_name: str,
) -> bool:
    """Patch one saved slide and require active-stream verification."""
    patch_l1_quartiles_in_deck(
        output_ppt=temp_ppt,
        slide_records=[
            {
                "function": function_name,
                "quartile_values": list(quartile_values),
            }
        ],
        logger=logger,
        slide_indices=[1],
    )
    logger.info(
        "Deferred carry-over for %s: the active think-cellXML stream was patched "
        "and verified. Manual carry-over remains required in the final deck.",
        function_name,
    )
    return True


__all__ = [
    "GEOMETRY_CARRYOVER_BUILD_ID",
    "build_quartile_targets",
    "patch_l1_quartiles_in_deck",
    "patch_quartile_geometry_in_ole_bytes",
    "patch_quartile_geometry_in_pptx",
    "prepare_quartile_lines_for_deferred_carryover",
    "settle_after_thinkcell_chart_update",
]
