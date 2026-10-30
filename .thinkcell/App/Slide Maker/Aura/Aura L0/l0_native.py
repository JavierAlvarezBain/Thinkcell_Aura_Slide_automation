from __future__ import annotations

"""Aura L0 native think-cell state and quartile geometry patching.

This module owns the L0 template's value-line identities, compound-file reader,
active think-cellXML mutations, native segment/series fills, native legend
ordering, and the final closed-deck patch.  It intentionally does not reuse the
Aura L1 quartile handler.
"""

import html
import math
from pathlib import Path
import posixpath
import re
import struct
import time
from typing import Any, Iterable
import xml.etree.ElementTree as ET
import zipfile

from l0_config import (
    L0_QUARTILE_LINE_COLORS_BY_RANK,
    L0_QUARTILE_PLACEHOLDERS_BY_RANK,
    L0QuartileTarget,
    SegmentStyle,
    TC_VECTOR_SEGMENT_START_INDEX,
)
from l0_source import (
    segment_display_name,
    segment_style_log_items,
    segment_style_plan,
)
from shared import BuildConfig, parse_numeric_text, pump_office_ui
from fte_quartile_table_detector import compact_text

def _build_l0_quartile_targets(quartile_values: Iterable[float | None]) -> list[L0QuartileTarget]:
    values = sorted(float(value) for value in quartile_values if value is not None)
    if len(values) < 3:
        return []
    lowest, middle, highest = values[:3]
    return [
        L0QuartileTarget("highest", highest, str(int(round(highest)))),
        L0QuartileTarget("middle", middle, str(int(round(middle)))),
        L0QuartileTarget("lowest", lowest, str(int(round(lowest)))),
    ]


def _is_group_shape(shape) -> bool:
    try:
        return int(shape.Type) == 6
    except Exception:
        return False


def _shape_has_visible_text(shape) -> bool:
    try:
        return bool(shape.HasTextFrame and shape.TextFrame.HasText)
    except Exception:
        return False


def _shape_bounds(shape) -> dict[str, float]:
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


def _simple_numeric_shape_text(shape) -> tuple[str, float] | None:
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


def _collect_axis_label_candidates(slide) -> list[dict[str, Any]]:
    from chart_updater import iter_shapes

    candidates: list[dict[str, Any]] = []
    for shape in iter_shapes(slide.Shapes):
        try:
            if _is_group_shape(shape) or not _shape_has_visible_text(shape):
                continue
            parsed = _simple_numeric_shape_text(shape)
            if parsed is None:
                continue
            _, value = parsed
            bounds = _shape_bounds(shape)
            if not (60 <= bounds["center_y"] <= 505):
                continue
            if bounds["width"] > 70 or bounds["height"] > 32:
                continue
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


def _axis_signature(candidates: list[dict[str, Any]]) -> tuple:
    return tuple(
        sorted(
            (
                round(float(item["value"]), 3),
                round(float(item["center_x"]), 1),
                round(float(item["center_y"]), 1),
            )
            for item in candidates
        )
    )


def _wait_for_axis_to_settle(slide, config: BuildConfig, logger, label: str = "axis") -> None:
    timeout = max(float(config.post_chart_update_wait_seconds), 0.3)
    started = time.perf_counter()
    previous_signature = None
    stable_count = 0
    while time.perf_counter() - started < timeout:
        pump_office_ui(0.18)
        try:
            signature = _axis_signature(_collect_axis_label_candidates(slide))
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


def settle_after_thinkcell_chart_update(pres, config: BuildConfig, logger) -> None:
    from chart_updater import find_chart_slide_index

    slide_index = find_chart_slide_index(pres, config)
    try:
        pres.Windows(1).Activate()
    except Exception:
        pass
    try:
        pres.Application.ActiveWindow.View.GotoSlide(slide_index)
    except Exception:
        pass
    try:
        pres.Slides(slide_index).Select()
    except Exception:
        pass
    pump_office_ui(0.2)
    _wait_for_axis_to_settle(
        pres.Slides(slide_index),
        config,
        logger,
        label="post-think-cell L0 chart update",
    )


def _format_thinkcell_float_literal(value: float) -> str:
    return f"{float(value):.20E}"


def _as_int_attribute_like(old_text: str, new_value: int) -> str:
    value = int(round(new_value))
    width = len(old_text)
    value_text = str(value)
    if len(value_text) == width:
        return value_text
    if value >= 0 and len(value_text) < width:
        return value_text.rjust(width, "0")
    return old_text


def _linear_internal_axis_mapping(
    points: list[tuple[float, float]],
) -> tuple[float, float] | None:
    if len(points) < 2:
        return None
    n = len(points)
    sx = sum(x for x, _ in points)
    sy = sum(y for _, y in points)
    sxx = sum(x * x for x, _ in points)
    sxy = sum(x * y for x, y in points)
    denominator = n * sxx - sx * sx
    if abs(denominator) < 1e-9:
        return None
    b = (n * sxy - sx * sy) / denominator
    a = (sy - b * sx) / n
    if b >= 0:
        return None
    return a, b


def _extract_internal_axis_mapping_from_tc_root(
    root_text: str,
) -> tuple[float, float, list[tuple[float, float]]] | None:
    label_y_by_id: dict[str, float] = {}
    for match in re.finditer(
        r'<CDataAxisLabel id="(\d+)">(.*?)</CDataAxisLabel>',
        root_text,
        flags=re.S,
    ):
        label_id = match.group(1)
        block = match.group(2)
        rectangle = re.search(
            r'<m_rectPosition[^>]*\btop="([0-9.+\-E]+)"[^>]*\bbottom="([0-9.+\-E]+)"',
            block,
            flags=re.S,
        )
        if not rectangle:
            continue
        try:
            top = float(rectangle.group(1))
            bottom = float(rectangle.group(2))
        except Exception:
            continue
        label_y_by_id[label_id] = (top + bottom) / 2.0

    points: list[tuple[float, float]] = []
    for match in re.finditer(
        r'<CDataAxisTickmark id="(\d+)">(.*?)</CDataAxisTickmark>',
        root_text,
        flags=re.S,
    ):
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
    ordered_points = sorted(unique.values(), key=lambda item: item[0])
    mapping = _linear_internal_axis_mapping(ordered_points)
    if mapping is None:
        return None
    a, b = mapping
    return a, b, ordered_points


def _get_tc_value_line_infos(root_text: str) -> list[dict[str, Any]]:
    infos: list[dict[str, Any]] = []
    for match in re.finditer(
        r'<CSequenceChartValueIndicatorLine id="(\d+)">(.*?)</CSequenceChartValueIndicatorLine>',
        root_text,
        flags=re.S,
    ):
        line_id = match.group(1)
        block = match.group(2)
        variable_reference = re.search(r'<m_varsrc idref="(\d+)"/>', block)
        if not variable_reference:
            continue
        variable_id = variable_reference.group(1)
        variable_definition = re.search(
            rf'(<CVariableSource id="{re.escape(variable_id)}">.*?<m_varval type="1" val=")([^"]+)(".*?</CVariableSource>)',
            root_text,
            flags=re.S,
        )
        if not variable_definition:
            continue
        try:
            value = float(variable_definition.group(2))
        except Exception:
            continue

        text_variable_id = None
        text_variable_match = re.search(
            rf'<CVariableSource id="{re.escape(variable_id)}">.*?<m_ctextvar><elem idref="(\d+)"/></m_ctextvar>.*?</CVariableSource>',
            root_text,
            flags=re.S,
        )
        if text_variable_match:
            text_variable_id = text_variable_match.group(1)

        line_shape_ids: list[str] = []
        line_shape_block = re.search(
            r'<m_cpptline[^>]*>(.*?)</m_cpptline>',
            block,
            flags=re.S,
        )
        if line_shape_block:
            line_shape_ids = re.findall(
                r'<elem idref="(\d+)"/>',
                line_shape_block.group(1),
            )

        label_rectangle = re.search(
            r'(<m_scvalindlabel>.*?<m_rectPosition[^>]*\btop=")([0-9]+)("[^>]*\bbottom=")([0-9]+)(".*?</m_scvalindlabel>)',
            block,
            flags=re.S,
        )
        if label_rectangle:
            top = int(label_rectangle.group(2))
            bottom = int(label_rectangle.group(4))
            half_height = max(1, int(round((bottom - top) / 2)))
            label_center_y = (top + bottom) / 2.0
        else:
            half_height = 48
            label_center_y = None

        # Read the value-line color from m_linestyle specifically. The same
        # block also contains black label-font RGB, which must not be mistaken
        # for the line identity.
        line_color_hex = None
        line_style_match = re.search(
            r'<m_linestyle>(.*?)</m_linestyle>',
            block,
            flags=re.S,
        )
        if line_style_match:
            line_color_match = re.search(
                r'<m_col>.*?<m_rgb r="([0-9A-Fa-f]{2})" g="([0-9A-Fa-f]{2})" b="([0-9A-Fa-f]{2})"/>.*?</m_col>',
                line_style_match.group(1),
                flags=re.S,
            )
            if line_color_match:
                line_color_hex = "".join(line_color_match.groups()).upper()

        infos.append(
            {
                "line_id": line_id,
                "varsrc_id": variable_id,
                "textvar_id": text_variable_id,
                "var_value": value,
                "var_value_text": variable_definition.group(2),
                "line_shape_ids": line_shape_ids,
                "label_half_height": half_height,
                "label_center_y": label_center_y,
                "line_color_hex": line_color_hex,
            }
        )
    return infos

def _group_tc_value_line_infos(
    infos: list[dict[str, Any]],
) -> dict[tuple[str, str], list[dict[str, Any]]]:
    """Group duplicate serialized copies of one think-cell value line."""
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for info in infos:
        key = (str(info["line_id"]), str(info["varsrc_id"]))
        grouped.setdefault(key, []).append(info)
    return grouped


def _quartile_group_sort_value(
    group: list[dict[str, Any]],
    source_value: float,
) -> tuple[float, float, int, int]:
    values = [float(info["var_value"]) for info in group]
    centers = [
        float(info["label_center_y"])
        for info in group
        if info.get("label_center_y") is not None
    ]
    distance = min(abs(value - float(source_value)) for value in values)
    center = sum(centers) / len(centers) if centers else float("inf")
    line_id = int(str(group[0]["line_id"]))
    varsrc_id = int(str(group[0]["varsrc_id"]))
    return distance, center, line_id, varsrc_id


def _anchor_from_value_line_group(
    rank: str,
    group: list[dict[str, Any]],
) -> dict[str, Any]:
    if not group:
        raise ValueError(f"Cannot create an L0 quartile anchor for empty rank {rank!r}.")
    line_ids = {str(info["line_id"]) for info in group}
    varsrc_ids = {str(info["varsrc_id"]) for info in group}
    if len(line_ids) != 1 or len(varsrc_ids) != 1:
        raise RuntimeError(
            f"L0 quartile anchor group for {rank} is not a single line identity: "
            f"line_ids={sorted(line_ids)}, varsrc_ids={sorted(varsrc_ids)}"
        )
    return {
        "rank": rank,
        "line_id": next(iter(line_ids)),
        "varsrc_id": next(iter(varsrc_ids)),
        "textvar_ids": sorted(
            {
                str(info["textvar_id"])
                for info in group
                if info.get("textvar_id")
            }
        ),
        "line_shape_ids": sorted(
            {
                str(shape_id)
                for info in group
                for shape_id in info.get("line_shape_ids", [])
            }
        ),
        "source_values": sorted({float(info["var_value"]) for info in group}),
        "line_colors": sorted(
            {
                str(info.get("line_color_hex") or "").upper()
                for info in group
                if info.get("line_color_hex")
            }
        ),
        "serialized_copies": len(group),
        "label_half_height": max(
            int(info.get("label_half_height", 48)) for info in group
        ),
    }


def _capture_l0_quartile_anchor_from_root(
    root_text: str,
    preferred_anchor: dict[str, dict[str, Any]] | None = None,
) -> dict[str, dict[str, Any]]:
    """
    Capture the exact three native quartile-line identities before any live
    chart recoloring.  The final quartile pass can then target those IDs rather
    than rediscovering lines from colors that may also be used by Ops segments.
    """
    infos = _get_tc_value_line_infos(root_text)
    grouped = _group_tc_value_line_infos(infos)
    if len(grouped) < 3:
        raise ValueError(
            f"Expected at least three distinct L0 value lines, found {len(grouped)}."
        )

    chosen: dict[str, list[dict[str, Any]]] = {}
    used_keys: set[tuple[str, str]] = set()

    for rank in ("highest", "middle", "lowest"):
        source_value = float(L0_QUARTILE_PLACEHOLDERS_BY_RANK[rank])
        expected_colors = {
            color.upper()
            for color in L0_QUARTILE_LINE_COLORS_BY_RANK.get(rank, set())
        }

        selected_key: tuple[str, str] | None = None
        preferred = (preferred_anchor or {}).get(rank) or {}
        if preferred.get("line_id") and preferred.get("varsrc_id"):
            preferred_key = (
                str(preferred["line_id"]),
                str(preferred["varsrc_id"]),
            )
            if preferred_key in grouped and preferred_key not in used_keys:
                selected_key = preferred_key

        available = [
            (key, group)
            for key, group in grouped.items()
            if key not in used_keys
        ]

        if selected_key is None:
            placeholder_groups = [
                (key, group)
                for key, group in available
                if any(
                    int(round(float(info["var_value"]))) == int(round(source_value))
                    for info in group
                )
            ]
            color_placeholder_groups = [
                (key, group)
                for key, group in placeholder_groups
                if any(
                    str(info.get("line_color_hex") or "").upper() in expected_colors
                    for info in group
                )
            ]
            candidates = color_placeholder_groups or placeholder_groups
            if candidates:
                selected_key = min(
                    candidates,
                    key=lambda item: _quartile_group_sort_value(item[1], source_value),
                )[0]

        if selected_key is None:
            color_groups = [
                (key, group)
                for key, group in available
                if any(
                    str(info.get("line_color_hex") or "").upper() in expected_colors
                    for info in group
                )
            ]
            if color_groups:
                # Multiple value-indicator lines can share a Bain family color
                # after a live chart save.  The actual template quartile line
                # is the group closest to its original placeholder value.
                selected_key = min(
                    color_groups,
                    key=lambda item: _quartile_group_sort_value(item[1], source_value),
                )[0]

        if selected_key is None:
            remaining_with_geometry = [
                (key, group)
                for key, group in available
                if any(info.get("label_center_y") is not None for info in group)
            ]
            ordered = sorted(
                remaining_with_geometry,
                key=lambda item: sum(
                    float(info["label_center_y"])
                    for info in item[1]
                    if info.get("label_center_y") is not None
                )
                / max(
                    1,
                    sum(
                        1
                        for info in item[1]
                        if info.get("label_center_y") is not None
                    ),
                ),
            )
            if len(ordered) >= 3:
                index = {"highest": 0, "middle": len(ordered) // 2, "lowest": -1}[rank]
                selected_key = ordered[index][0]

        if selected_key is None:
            observed = [
                {
                    "line_id": key[0],
                    "varsrc_id": key[1],
                    "values": sorted({float(info["var_value"]) for info in group}),
                    "colors": sorted(
                        {
                            str(info.get("line_color_hex") or "").upper()
                            for info in group
                            if info.get("line_color_hex")
                        }
                    ),
                }
                for key, group in grouped.items()
            ]
            raise ValueError(
                f"Could not capture the pre-color L0 {rank} quartile line. "
                f"Observed value-line groups: {observed}"
            )

        used_keys.add(selected_key)
        chosen[rank] = grouped[selected_key]

    anchors = {
        rank: _anchor_from_value_line_group(rank, chosen[rank])
        for rank in ("highest", "middle", "lowest")
    }
    line_ids = [anchors[rank]["line_id"] for rank in ("highest", "middle", "lowest")]
    if len(set(line_ids)) != 3:
        raise RuntimeError(
            f"Captured L0 quartile anchors do not have three distinct line IDs: {line_ids}"
        )
    return anchors



# -----------------------------------------------------------------------------
# Compound-file access for the active think-cell XML stream
# -----------------------------------------------------------------------------

_CFB_FREESECT = 0xFFFFFFFF
_CFB_ENDOFCHAIN = 0xFFFFFFFE
_CFB_FATSECT = 0xFFFFFFFD
_CFB_DIFSECT = 0xFFFFFFFC


class _CompoundFileBinary:
    """Minimal in-place CFB reader/writer for embedded think-cell OLE objects.

    The PowerPoint OLE object is a Compound File Binary container.  Scanning the
    complete ``oleObject*.bin`` byte array is unsafe because unused/slack sectors
    can retain older copies of think-cell XML.  This class resolves the active
    directory stream and patches only the live ``think-cellXML`` stream.

    Most mutations remain byte-length preserving.  High-cardinality charts can
    contain native ``<m_fill/>`` placeholders for newly added series; those are
    expanded into explicit fills.  The writer therefore supports growing a
    regular stream by allocating FREESECT sectors already present in the CFB
    container and updating the FAT plus directory size in place.
    """

    def __init__(self, data: bytes):
        self._buffer = bytearray(data)
        if bytes(self._buffer[:8]) != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("The embedded OLE object is not a Compound File Binary container.")

        self.sector_size = 1 << struct.unpack_from("<H", self._buffer, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", self._buffer, 32)[0]
        self.number_of_fat_sectors = struct.unpack_from("<I", self._buffer, 44)[0]
        self.first_directory_sector = struct.unpack_from("<I", self._buffer, 48)[0]
        self.mini_stream_cutoff = struct.unpack_from("<I", self._buffer, 56)[0]
        self.first_mini_fat_sector = struct.unpack_from("<I", self._buffer, 60)[0]
        self.number_of_mini_fat_sectors = struct.unpack_from("<I", self._buffer, 64)[0]
        self.first_difat_sector = struct.unpack_from("<I", self._buffer, 68)[0]
        self.number_of_difat_sectors = struct.unpack_from("<I", self._buffer, 72)[0]

        difat = list(struct.unpack_from("<109I", self._buffer, 76))
        next_difat_sector = self.first_difat_sector
        for _ in range(self.number_of_difat_sectors):
            if next_difat_sector in (_CFB_FREESECT, _CFB_ENDOFCHAIN):
                break
            sector = self._sector_bytes(next_difat_sector)
            values = list(
                struct.unpack(
                    f"<{self.sector_size // 4}I",
                    sector,
                )
            )
            difat.extend(values[:-1])
            next_difat_sector = values[-1]

        self.fat_sector_ids = [
            value
            for value in difat
            if value not in (
                _CFB_FREESECT,
                _CFB_ENDOFCHAIN,
                _CFB_FATSECT,
                _CFB_DIFSECT,
            )
        ][: self.number_of_fat_sectors]
        self.fat: list[int] = []
        for sector_id in self.fat_sector_ids:
            self.fat.extend(
                struct.unpack(
                    f"<{self.sector_size // 4}I",
                    self._sector_bytes(sector_id),
                )
            )
        self.total_sector_count = max(
            0,
            (len(self._buffer) // self.sector_size) - 1,
        )

        directory_bytes = self._read_regular_stream(
            self.first_directory_sector,
            size=None,
        )
        self.entries: list[dict[str, Any]] = []
        for offset in range(0, len(directory_bytes), 128):
            raw_entry = directory_bytes[offset : offset + 128]
            if len(raw_entry) < 128:
                break
            name_length = struct.unpack_from("<H", raw_entry, 64)[0]
            if name_length >= 2:
                name = raw_entry[: name_length - 2].decode("utf-16le", errors="ignore")
            else:
                name = ""
            self.entries.append(
                {
                    "id": offset // 128,
                    "name": name,
                    "type": int(raw_entry[66]),
                    "start_sector": struct.unpack_from("<I", raw_entry, 116)[0],
                    "size": struct.unpack_from("<Q", raw_entry, 120)[0],
                }
            )

        self.root_entry = next(
            (entry for entry in self.entries if entry["type"] == 5),
            None,
        )
        if self.root_entry is None:
            raise ValueError("The embedded OLE object has no root storage entry.")

        self.mini_fat: list[int] = []
        if (
            self.number_of_mini_fat_sectors > 0
            and self.first_mini_fat_sector not in (_CFB_FREESECT, _CFB_ENDOFCHAIN)
        ):
            mini_fat_bytes = self._read_regular_stream(
                self.first_mini_fat_sector,
                size=self.number_of_mini_fat_sectors * self.sector_size,
            )
            whole_dword_length = (len(mini_fat_bytes) // 4) * 4
            self.mini_fat = list(
                struct.unpack(
                    f"<{whole_dword_length // 4}I",
                    mini_fat_bytes[:whole_dword_length],
                )
            )

        root_size = int(self.root_entry["size"])
        if root_size > 0:
            self.mini_stream = bytearray(
                self._read_regular_stream(
                    int(self.root_entry["start_sector"]),
                    size=root_size,
                )
            )
        else:
            self.mini_stream = bytearray()

    def _sector_offset(self, sector_id: int) -> int:
        return (int(sector_id) + 1) * self.sector_size

    def _sector_bytes(self, sector_id: int) -> bytes:
        start = self._sector_offset(sector_id)
        return bytes(self._buffer[start : start + self.sector_size])

    @staticmethod
    def _chain(start_sector: int, allocation_table: list[int]) -> list[int]:
        result: list[int] = []
        seen: set[int] = set()
        sector_id = int(start_sector)
        while (
            sector_id not in (_CFB_FREESECT, _CFB_ENDOFCHAIN)
            and 0 <= sector_id < len(allocation_table)
            and sector_id not in seen
        ):
            seen.add(sector_id)
            result.append(sector_id)
            sector_id = int(allocation_table[sector_id])
        return result

    def _read_regular_stream(self, start_sector: int, size: int | None) -> bytes:
        payload = b"".join(
            self._sector_bytes(sector_id)
            for sector_id in self._chain(start_sector, self.fat)
        )
        return payload if size is None else payload[: int(size)]

    def _write_regular_stream(
        self,
        start_sector: int,
        stream_size: int,
        payload: bytes,
    ) -> None:
        if len(payload) != int(stream_size):
            raise ValueError(
                "A CFB stream mutation must preserve the exact stream byte length."
            )
        cursor = 0
        for sector_id in self._chain(start_sector, self.fat):
            if cursor >= stream_size:
                break
            count = min(self.sector_size, stream_size - cursor)
            destination = self._sector_offset(sector_id)
            self._buffer[destination : destination + count] = payload[
                cursor : cursor + count
            ]
            cursor += count
        if cursor != stream_size:
            raise ValueError(
                f"The CFB FAT chain exposed {cursor} bytes for a {stream_size}-byte stream."
            )

    @property
    def stream_entries(self) -> list[dict[str, Any]]:
        return [entry for entry in self.entries if entry["type"] == 2]

    def read_stream(self, entry: dict[str, Any]) -> bytes:
        stream_size = int(entry["size"])
        if stream_size <= 0:
            return b""
        start_sector = int(entry["start_sector"])
        if stream_size < self.mini_stream_cutoff:
            chunks: list[bytes] = []
            for mini_sector_id in self._chain(start_sector, self.mini_fat):
                offset = mini_sector_id * self.mini_sector_size
                chunks.append(
                    bytes(
                        self.mini_stream[
                            offset : offset + self.mini_sector_size
                        ]
                    )
                )
            return b"".join(chunks)[:stream_size]
        return self._read_regular_stream(start_sector, stream_size)

    def _flush_fat(self) -> None:
        entries_per_sector = self.sector_size // 4
        for fat_index, sector_id in enumerate(self.fat_sector_ids):
            start = fat_index * entries_per_sector
            values = list(self.fat[start : start + entries_per_sector])
            if len(values) < entries_per_sector:
                values.extend([_CFB_FREESECT] * (entries_per_sector - len(values)))
            destination = self._sector_offset(sector_id)
            self._buffer[
                destination : destination + self.sector_size
            ] = struct.pack(f"<{entries_per_sector}I", *values)

    def _update_directory_entry_size(
        self,
        entry: dict[str, Any],
        stream_size: int,
    ) -> None:
        directory_bytes = bytearray(
            self._read_regular_stream(self.first_directory_sector, size=None)
        )
        offset = int(entry["id"]) * 128
        if offset + 128 > len(directory_bytes):
            raise ValueError(
                f"CFB directory entry {entry['id']} is outside the directory stream."
            )
        struct.pack_into("<Q", directory_bytes, offset + 120, int(stream_size))
        self._write_regular_stream(
            self.first_directory_sector,
            len(directory_bytes),
            bytes(directory_bytes),
        )
        entry["size"] = int(stream_size)
        if int(entry["id"]) < len(self.entries):
            self.entries[int(entry["id"])]["size"] = int(stream_size)

    def _allocate_regular_sectors(self, count: int) -> list[int]:
        if count <= 0:
            return []

        free_existing = [
            sector_id
            for sector_id in range(min(self.total_sector_count, len(self.fat)))
            if self.fat[sector_id] == _CFB_FREESECT
        ]
        selected = free_existing[:count]
        remaining = count - len(selected)

        if remaining > 0:
            first_new_id = self.total_sector_count
            available_fat_slots = len(self.fat) - first_new_id
            if available_fat_slots < remaining:
                raise RuntimeError(
                    "The embedded OLE object has insufficient FAT capacity to "
                    f"grow the think-cell XML stream by {remaining} sector(s)."
                )
            selected.extend(range(first_new_id, first_new_id + remaining))
            self._buffer.extend(b"\x00" * (remaining * self.sector_size))
            self.total_sector_count += remaining

        for sector_id in selected:
            self.fat[sector_id] = _CFB_ENDOFCHAIN
        return selected

    def write_stream_resized(
        self,
        entry: dict[str, Any],
        payload: bytes,
    ) -> None:
        """Write a regular stream and grow its FAT chain when required."""
        old_size = int(entry["size"])
        new_size = len(payload)
        if old_size < self.mini_stream_cutoff or new_size < self.mini_stream_cutoff:
            if new_size != old_size:
                raise ValueError(
                    "Resizing a CFB mini-stream is not supported by this L0 patcher."
                )
            self.write_stream(entry, payload)
            return

        start_sector = int(entry["start_sector"])
        chain = self._chain(start_sector, self.fat)
        required_sector_count = max(
            1,
            (new_size + self.sector_size - 1) // self.sector_size,
        )
        if required_sector_count > len(chain):
            additions = self._allocate_regular_sectors(
                required_sector_count - len(chain)
            )
            if not chain:
                if not additions:
                    raise RuntimeError("Could not allocate the first stream sector.")
                entry["start_sector"] = additions[0]
                start_sector = additions[0]
                chain = []
            else:
                self.fat[chain[-1]] = additions[0]
            for index, sector_id in enumerate(additions):
                self.fat[sector_id] = (
                    additions[index + 1]
                    if index + 1 < len(additions)
                    else _CFB_ENDOFCHAIN
                )
            chain.extend(additions)
            self._flush_fat()

        cursor = 0
        for sector_id in chain[:required_sector_count]:
            destination = self._sector_offset(sector_id)
            chunk = payload[cursor : cursor + self.sector_size]
            self._buffer[
                destination : destination + self.sector_size
            ] = chunk + (b"\x00" * (self.sector_size - len(chunk)))
            cursor += len(chunk)
        if cursor != new_size:
            raise ValueError(
                f"The resized CFB stream wrote {cursor} of {new_size} bytes."
            )

        self._update_directory_entry_size(entry, new_size)
        self._flush_fat()

    def write_stream(self, entry: dict[str, Any], payload: bytes) -> None:
        stream_size = int(entry["size"])
        if len(payload) != stream_size:
            self.write_stream_resized(entry, payload)
            return
        start_sector = int(entry["start_sector"])
        if stream_size < self.mini_stream_cutoff:
            cursor = 0
            for mini_sector_id in self._chain(start_sector, self.mini_fat):
                if cursor >= stream_size:
                    break
                count = min(self.mini_sector_size, stream_size - cursor)
                destination = mini_sector_id * self.mini_sector_size
                self.mini_stream[destination : destination + count] = payload[
                    cursor : cursor + count
                ]
                cursor += count
            if cursor != stream_size:
                raise ValueError(
                    f"The CFB mini-FAT chain exposed {cursor} bytes for a "
                    f"{stream_size}-byte stream."
                )
            root_size = int(self.root_entry["size"])
            if root_size:
                self._write_regular_stream(
                    int(self.root_entry["start_sector"]),
                    root_size,
                    bytes(self.mini_stream),
                )
            return
        self._write_regular_stream(start_sector, stream_size, payload)

    def to_bytes(self) -> bytes:
        return bytes(self._buffer)


def _decode_thinkcell_root(root_bytes: bytes) -> tuple[str, str]:
    try:
        return root_bytes.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        return root_bytes.decode("latin-1"), "latin-1"


def _active_thinkcell_root_records_from_ole(
    data: bytes,
) -> tuple[_CompoundFileBinary, list[dict[str, Any]]]:
    compound = _CompoundFileBinary(data)
    records: list[dict[str, Any]] = []
    for entry in compound.stream_entries:
        stream_bytes = compound.read_stream(entry)
        if (
            b"CSequenceChartValueIndicatorLine" not in stream_bytes
            or b"CSequenceChartDataTable" not in stream_bytes
        ):
            continue
        for root_match in re.finditer(rb"<root\b.*?</root>", stream_bytes, flags=re.S):
            root_bytes = root_match.group(0)
            root_text, encoding = _decode_thinkcell_root(root_bytes)
            if (
                "CSequenceChartValueIndicatorLine" not in root_text
                or "CSequenceChartDataTable" not in root_text
            ):
                continue
            records.append(
                {
                    "entry": entry,
                    "stream_bytes": stream_bytes,
                    "root_start": root_match.start(),
                    "root_end": root_match.end(),
                    "root_text": root_text,
                    "encoding": encoding,
                }
            )
    return compound, records


def _active_thinkcell_roots_from_ole_bytes(data: bytes) -> list[str]:
    _compound, records = _active_thinkcell_root_records_from_ole(data)
    return [record["root_text"] for record in records]


def _ole_contains_active_thinkcell_chart(data: bytes) -> bool:
    try:
        return bool(_active_thinkcell_roots_from_ole_bytes(data))
    except Exception:
        return False


def capture_l0_quartile_anchor_from_pptx(
    pptx_path: Path,
    logger,
    context: str,
    preferred_anchor: dict[str, dict[str, Any]] | None = None,
) -> dict[str, dict[str, Any]]:
    """Read one saved PPTX and capture its native quartile IDs before colors."""
    _wait_for_readable_pptx(pptx_path, logger)
    errors: list[str] = []
    with zipfile.ZipFile(pptx_path, "r") as archive:
        for slide_part in _ordered_slide_parts(archive):
            for embedding_part in _slide_thinkcell_embedding_parts(archive, slide_part):
                data = archive.read(embedding_part)
                try:
                    active_roots = _active_thinkcell_roots_from_ole_bytes(data)
                except Exception as exc:
                    errors.append(f"{embedding_part}: could not read active CFB stream: {exc}")
                    continue
                for root_text in active_roots:
                    try:
                        anchor = _capture_l0_quartile_anchor_from_root(
                            root_text,
                            preferred_anchor=preferred_anchor,
                        )
                    except Exception as exc:
                        errors.append(f"{embedding_part}: {exc}")
                        continue
                    logger.info(
                        "Captured pre-color L0 quartile anchors for %s: %s",
                        context,
                        {
                            rank: {
                                "line_id": anchor[rank]["line_id"],
                                "varsrc_id": anchor[rank]["varsrc_id"],
                                "source_values": anchor[rank]["source_values"],
                                "colors": anchor[rank]["line_colors"],
                            }
                            for rank in ("highest", "middle", "lowest")
                        },
                    )
                    return anchor
    raise ValueError(
        f"Could not capture native L0 quartile anchors from {pptx_path}. "
        + (" | ".join(errors) if errors else "No value-line root was found.")
    )


def _root_anchor_match_count(
    root_text: str,
    quartile_anchor: dict[str, dict[str, Any]] | None,
) -> int:
    if not quartile_anchor:
        return 0
    infos = _get_tc_value_line_infos(root_text)
    present = {
        (str(info["line_id"]), str(info["varsrc_id"]))
        for info in infos
    }
    return sum(
        1
        for rank in ("highest", "middle", "lowest")
        if (
            str((quartile_anchor.get(rank) or {}).get("line_id", "")),
            str((quartile_anchor.get(rank) or {}).get("varsrc_id", "")),
        )
        in present
    )


def _make_same_length_label_format(old_format: str, label: str) -> str:
    if not old_format:
        return old_format
    label = str(label)
    if len(label) > len(old_format):
        return old_format
    return label + ("'" * (len(old_format) - len(label)))


def _patch_value_line_label_block(
    block: str,
    new_y: int,
    half_height: int,
) -> tuple[str, int]:
    count = 0
    new_top = int(round(new_y - half_height))
    new_bottom = int(round(new_y + half_height))

    def replace_integer_rectangle(match: re.Match) -> str:
        nonlocal count
        top_text = _as_int_attribute_like(match.group(2), new_top)
        bottom_text = _as_int_attribute_like(match.group(4), new_bottom)
        replacement = (
            match.group(1)
            + top_text
            + match.group(3)
            + bottom_text
            + match.group(5)
        )
        if len(replacement) == len(match.group(0)):
            count += 1
            return replacement
        return match.group(0)

    block = re.sub(
        r'(<m_rect(?:Position|PPTShape)[^>]*\btop=")([0-9]+)("[^>]*\bbottom=")([0-9]+)("[^>]*/>)',
        replace_integer_rectangle,
        block,
        flags=re.S,
    )

    def replace_float_rectangle(match: re.Match) -> str:
        nonlocal count
        top_text = _format_thinkcell_float_literal(new_top)
        bottom_text = _format_thinkcell_float_literal(new_bottom)
        if len(match.group(2)) != len(top_text) or len(match.group(4)) != len(bottom_text):
            return match.group(0)
        replacement = (
            match.group(1)
            + top_text
            + match.group(3)
            + bottom_text
            + match.group(5)
        )
        if len(replacement) == len(match.group(0)):
            count += 1
            return replacement
        return match.group(0)

    block = re.sub(
        r'(<m_rectPosition[^>]*reqver="36208"[^>]*\btop=")([0-9.+\-E]+)("[^>]*\bbottom=")([0-9.+\-E]+)("[^>]*/>)',
        replace_float_rectangle,
        block,
        flags=re.S,
    )
    return block, count



def _zip_info_copy(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    copied = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    copied.compress_type = info.compress_type
    copied.comment = info.comment
    copied.extra = info.extra
    copied.create_system = info.create_system
    copied.create_version = info.create_version
    copied.extract_version = info.extract_version
    copied.flag_bits = info.flag_bits
    copied.internal_attr = info.internal_attr
    copied.external_attr = info.external_attr
    copied.volume = info.volume
    return copied


def _related_part_path(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(source_part), target)
    )


def _rels_part_path(source_part: str) -> str:
    parent, filename = posixpath.split(source_part)
    return posixpath.join(parent, "_rels", filename + ".rels")


def _ordered_slide_parts(archive: zipfile.ZipFile) -> list[str]:
    presentation_part = "ppt/presentation.xml"
    presentation_rels_part = "ppt/_rels/presentation.xml.rels"
    relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    presentation_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    rel_root = ET.fromstring(archive.read(presentation_rels_part))
    slide_by_rel_id: dict[str, str] = {}
    for relationship in rel_root.findall(f"{{{relationship_ns}}}Relationship"):
        rel_type = relationship.attrib.get("Type", "")
        if not rel_type.endswith("/slide"):
            continue
        rel_id = relationship.attrib.get("Id")
        target = relationship.attrib.get("Target")
        if rel_id and target:
            slide_by_rel_id[rel_id] = _related_part_path(
                presentation_part,
                target,
            )

    presentation_root = ET.fromstring(archive.read(presentation_part))
    slide_parts: list[str] = []
    for slide_id in presentation_root.findall(
        f".//{{{presentation_ns}}}sldId"
    ):
        rel_id = slide_id.attrib.get(f"{{{office_rel_ns}}}id")
        if rel_id and rel_id in slide_by_rel_id:
            slide_parts.append(slide_by_rel_id[rel_id])
    return slide_parts


def _slide_thinkcell_embedding_parts(
    archive: zipfile.ZipFile,
    slide_part: str,
) -> list[str]:
    relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    rels_part = _rels_part_path(slide_part)
    if rels_part not in archive.namelist():
        return []

    root = ET.fromstring(archive.read(rels_part))
    result: list[str] = []
    for relationship in root.findall(f"{{{relationship_ns}}}Relationship"):
        rel_type = relationship.attrib.get("Type", "")
        target = relationship.attrib.get("Target")
        if not rel_type.endswith("/oleObject") or not target:
            continue
        target_part = _related_part_path(slide_part, target)
        if not target_part.startswith("ppt/embeddings/"):
            continue
        try:
            data = archive.read(target_part)
        except KeyError:
            continue
        if _ole_contains_active_thinkcell_chart(data):
            result.append(target_part)
    return sorted(set(result))


def _wait_for_readable_pptx(
    path: Path,
    logger,
    timeout_seconds: float = 45.0,
) -> None:
    deadline = time.monotonic() + max(1.0, float(timeout_seconds))
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            with zipfile.ZipFile(path, "r") as archive:
                archive.read("[Content_Types].xml")
                if archive.testzip() is not None:
                    raise zipfile.BadZipFile("A member of the PPTX package failed CRC validation.")
            return
        except (PermissionError, OSError, zipfile.BadZipFile, KeyError) as exc:
            last_error = exc
            time.sleep(0.35)
    raise RuntimeError(
        f"PowerPoint did not release a readable PPTX package within {timeout_seconds:.0f}s: {path}"
    ) from last_error


def _replace_file_with_retry(
    source: Path,
    destination: Path,
    timeout_seconds: float = 60.0,
) -> None:
    deadline = time.monotonic() + max(1.0, float(timeout_seconds))
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            source.replace(destination)
            return
        except (PermissionError, OSError) as exc:
            last_error = exc
            try:
                pump_office_ui(0.20)
            except Exception:
                pass
            time.sleep(0.25)
    raise RuntimeError(
        f"Could not replace the final Aura L0 deck after quartile patching: {destination}"
    ) from last_error


def _patch_fill_on_top_flags_in_ole_bytes(data: bytes) -> tuple[bytes, int, int]:
    """Enable every serialized think-cell datasheet-fill-on-top flag in-place."""
    off_flag = b'<m_bExcelOnTop val="0"/>'
    on_flag = b'<m_bExcelOnTop val="1"/>'
    matched = data.count(off_flag) + data.count(on_flag)
    if matched <= 0:
        return data, 0, 0

    changed = data.count(off_flag)
    patched = data.replace(off_flag, on_flag)
    remaining_off = patched.count(off_flag)
    verified_on = patched.count(on_flag)
    if remaining_off or verified_on != matched or len(patched) != len(data):
        raise RuntimeError(
            "think-cell fill-on-top template patch verification failed: "
            f"matched={matched}, changed={changed}, verified_on={verified_on}, "
            f"remaining_off={remaining_off}"
        )
    return patched, changed, verified_on


def prepare_l0_fill_on_top_template(
    source_pptx: Path,
    destination_pptx: Path,
    logger,
) -> Path:
    """
    Create a runtime copy of the L0 template with ``Use Datasheet Fill on Top``
    enabled *before* think-cell receives the formatted Excel range.

    This is intentionally a temporary copy; the user's source template remains
    untouched. The replacement is byte-length preserving inside each OLE object.
    """
    source_pptx = Path(source_pptx)
    destination_pptx = Path(destination_pptx)
    destination_pptx.parent.mkdir(parents=True, exist_ok=True)
    if destination_pptx.exists():
        destination_pptx.unlink()

    total_matched = 0
    total_changed = 0
    patched_parts: list[str] = []

    with zipfile.ZipFile(source_pptx, "r") as source_archive:
        with zipfile.ZipFile(destination_pptx, "w") as target_archive:
            for info in source_archive.infolist():
                data = source_archive.read(info.filename)
                if info.filename.startswith("ppt/embeddings/") and b"m_bExcelOnTop" in data:
                    data, changed, verified = _patch_fill_on_top_flags_in_ole_bytes(data)
                    if verified:
                        total_matched += verified
                        total_changed += changed
                        patched_parts.append(info.filename)
                target_archive.writestr(_zip_info_copy(info), data)

    if total_matched <= 0:
        try:
            destination_pptx.unlink()
        except Exception:
            pass
        raise RuntimeError(
            "The Aura L0 template contains no serialized think-cell "
            "m_bExcelOnTop setting."
        )

    _wait_for_readable_pptx(destination_pptx, logger)

    # Re-open the copy and verify that no chart active document still has the
    # option disabled.
    verified_on = 0
    remaining_off = 0
    with zipfile.ZipFile(destination_pptx, "r") as archive:
        for part_name in patched_parts:
            data = archive.read(part_name)
            verified_on += data.count(b'<m_bExcelOnTop val="1"/>')
            remaining_off += data.count(b'<m_bExcelOnTop val="0"/>')
    if remaining_off or verified_on != total_matched:
        raise RuntimeError(
            "Runtime L0 template fill-on-top verification failed: "
            f"expected_on={total_matched}, verified_on={verified_on}, "
            f"remaining_off={remaining_off}"
        )

    logger.info(
        "Prepared runtime L0 template before AddRangeData: "
        "Use Datasheet Fill on Top enabled in %s serialized copy/copies "
        "across %s embedding(s); changed=%s; source template left unchanged.",
        verified_on,
        len(set(patched_parts)),
        total_changed,
    )
    return destination_pptx


def _patch_all_variable_source_copies(
    root_text: str,
    varsrc_id: str,
    target_value: float,
) -> tuple[str, int, int]:
    pattern = re.compile(
        rf'(<CVariableSource id="{re.escape(varsrc_id)}">.*?<m_varval type="1" val=")([^\"]+)(".*?</CVariableSource>)',
        flags=re.S,
    )
    changed = 0
    matched = 0
    target_text = _format_thinkcell_float_literal(target_value)

    def replace(match: re.Match) -> str:
        nonlocal changed, matched
        matched += 1
        old_text = match.group(2)
        if len(old_text) != len(target_text):
            return match.group(0)
        new_block = match.group(1) + target_text + match.group(3)
        if new_block != match.group(0):
            changed += 1
        return new_block

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
        new_format = _make_same_length_label_format(
            old_format,
            label,
        )
        if len(new_format) != len(old_format):
            return match.group(0)
        new_block = match.group(1) + new_format + match.group(3)
        if new_block != match.group(0):
            changed += 1
        return new_block

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
        new_top = _as_int_attribute_like(match.group(2), target_y)
        new_bottom = _as_int_attribute_like(match.group(4), target_y)
        new_block = (
            match.group(1)
            + new_top
            + match.group(3)
            + new_bottom
            + match.group(5)
        )
        if len(new_block) != len(match.group(0)):
            return match.group(0)
        if new_block != match.group(0):
            changed += 1
        return new_block

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
        new_block, patch_count = _patch_value_line_label_block(
            old_block,
            target_y,
            half_height,
        )
        if len(new_block) != len(old_block):
            return old_block
        if new_block != old_block:
            changed += patch_count
        return new_block

    return pattern.sub(replace, root_text), changed, matched


def _collect_line_shape_centers_for_ids(
    root_text: str,
    shape_ids: Iterable[str],
) -> dict[str, list[float]]:
    wanted = {str(shape_id) for shape_id in shape_ids}
    result: dict[str, list[float]] = {shape_id: [] for shape_id in wanted}
    for match in re.finditer(
        r'<CPPTLine id="(\d+)">.*?<m_rectPPTShape[^>]*\btop="([0-9]+)"[^>]*\bbottom="([0-9]+)"[^>]*/>.*?</CPPTLine>',
        root_text,
        flags=re.S,
    ):
        shape_id = match.group(1)
        if shape_id not in wanted:
            continue
        top = int(match.group(2))
        bottom = int(match.group(3))
        result.setdefault(shape_id, []).append((top + bottom) / 2.0)
    return result


def _collect_value_indicator_label_centers_for_ids(
    root_text: str,
    line_ids: Iterable[str],
) -> dict[str, list[float]]:
    """
    Collect only the rectangles that position the visible value label itself.

    A think-cell value-indicator line also contains an
    ``m_pptgenlineConnector`` rectangle for the small connector/arrow between
    the label and the line. PowerPoint is allowed to offset that connector by a
    few internal units when it reserializes a slide. It is not the value-label
    position and therefore must not be used as a hard label-geometry check.
    """
    wanted = {str(line_id) for line_id in line_ids}
    result: dict[str, list[float]] = {line_id: [] for line_id in wanted}
    for match in re.finditer(
        r'<CSequenceChartValueIndicatorLine id="(\d+)">(.*?)</CSequenceChartValueIndicatorLine>',
        root_text,
        flags=re.S,
    ):
        line_id = match.group(1)
        if line_id not in wanted:
            continue
        block = match.group(2)
        # Exclude connector geometry from strict label verification. The
        # remaining rectangles are the two think-cell label-position caches
        # plus the PowerPoint text-box rectangle.
        label_block = re.sub(
            r'<m_pptgenlineConnector>.*?</m_pptgenlineConnector>',
            '',
            block,
            flags=re.S,
        )
        for rect in re.finditer(
            r'<m_rect(?:Position|PPTShape)[^>]*\btop="([0-9.+\-E]+)"[^>]*\bbottom="([0-9.+\-E]+)"',
            label_block,
            flags=re.S,
        ):
            try:
                top = float(rect.group(1))
                bottom = float(rect.group(2))
            except Exception:
                continue
            result.setdefault(line_id, []).append((top + bottom) / 2.0)
    return result


def _collect_value_indicator_connector_centers_for_ids(
    root_text: str,
    line_ids: Iterable[str],
) -> dict[str, list[float]]:
    """Collect connector rectangles for diagnostics only, never hard failure."""
    wanted = {str(line_id) for line_id in line_ids}
    result: dict[str, list[float]] = {line_id: [] for line_id in wanted}
    for match in re.finditer(
        r'<CSequenceChartValueIndicatorLine id="(\d+)">(.*?)</CSequenceChartValueIndicatorLine>',
        root_text,
        flags=re.S,
    ):
        line_id = match.group(1)
        if line_id not in wanted:
            continue
        block = match.group(2)
        connector_match = re.search(
            r'<m_pptgenlineConnector>(.*?)</m_pptgenlineConnector>',
            block,
            flags=re.S,
        )
        if not connector_match:
            continue
        for rect in re.finditer(
            r'<m_rectPPTShape[^>]*\btop="([0-9.+\-E]+)"[^>]*\bbottom="([0-9.+\-E]+)"',
            connector_match.group(1),
            flags=re.S,
        ):
            try:
                top = float(rect.group(1))
                bottom = float(rect.group(2))
            except Exception:
                continue
            result.setdefault(line_id, []).append((top + bottom) / 2.0)
    return result


def _verify_l0_quartile_geometry(
    root_text: str,
    rank: str,
    target_y: int,
    line_shape_ids: list[str],
    line_ids: list[str],
) -> dict[str, Any]:
    line_centers = _collect_line_shape_centers_for_ids(root_text, line_shape_ids)
    label_centers = _collect_value_indicator_label_centers_for_ids(root_text, line_ids)
    connector_centers = _collect_value_indicator_connector_centers_for_ids(
        root_text,
        line_ids,
    )

    bad_lines: dict[str, list[float]] = {}
    for shape_id, centers in line_centers.items():
        if not centers:
            bad_lines[shape_id] = []
            continue
        if any(abs(float(center) - float(target_y)) > 1.0 for center in centers):
            bad_lines[shape_id] = centers
    if bad_lines:
        raise RuntimeError(
            f"L0 {rank} line geometry did not verify at y={target_y}: {bad_lines}"
        )

    bad_labels: dict[str, list[float]] = {}
    for line_id, centers in label_centers.items():
        if not centers:
            bad_labels[line_id] = []
            continue
        # Only the actual value-label position caches and PowerPoint text-box
        # rectangles are strict. The connector/arrow rectangle is diagnostic
        # because PowerPoint may legally offset it by a few internal units.
        if any(abs(float(center) - float(target_y)) > 1.0 for center in centers):
            bad_labels[line_id] = centers
    if bad_labels:
        raise RuntimeError(
            f"L0 {rank} visible-label geometry did not verify at y={target_y}: {bad_labels}"
        )

    connector_offsets = {
        line_id: [round(float(center) - float(target_y), 4) for center in centers]
        for line_id, centers in connector_centers.items()
        if centers
    }

    return {
        "line_shape_centers": line_centers,
        "label_centers": label_centers,
        "connector_centers": connector_centers,
        "connector_offsets": connector_offsets,
    }


def _replace_first_rgb_in_fill(fill_block: str, style: SegmentStyle) -> tuple[str, bool]:
    r_value, g_value, b_value = style.rgb
    changed = False

    def replace_fill_type(match: re.Match) -> str:
        nonlocal changed
        # A true think-cell pattern fill needs both m_msofilltype=2 and the
        # existing m_msopatterntype/m_colBackground fields. If a generated
        # scalar block does not contain those fields, keep it as a solid Bain
        # color rather than writing an invalid half-pattern. The Excel source
        # range is still formatted with the requested pattern, so generated
        # charts normally arrive here with a complete pattern block.
        desired = "2" if style.patterned and "m_msopatterntype" in fill_block else "1"
        if match.group(1) == desired:
            return match.group(0)
        changed = True
        return f'<m_msofilltype val="{desired}"/>'

    fill_block = re.sub(
        r'<m_msofilltype val="([0-9])"/>',
        replace_fill_type,
        fill_block,
        count=1,
    )

    def replace_rgb(match: re.Match) -> str:
        nonlocal changed
        replacement = f'<m_rgb r="{r_value}" g="{g_value}" b="{b_value}"/>'
        if replacement != match.group(0):
            changed = True
        return replacement

    fill_block = re.sub(
        r'<m_rgb r="[0-9A-F]{2}" g="[0-9A-F]{2}" b="[0-9A-F]{2}"/>',
        replace_rgb,
        fill_block,
        count=1,
    )
    return fill_block, changed


def _explicit_tc_fill_xml(tag_name: str, style: SegmentStyle) -> str:
    r_value, g_value, b_value = style.rgb
    if style.patterned:
        return (
            f'<{tag_name}><m_bVisible val="1"/><m_msofilltype val="2"/>'
            f'<m_col><m_msothmcolidx val="0"/><m_rgb r="{r_value}" '
            f'g="{g_value}" b="{b_value}"/></m_col>'
            f'<m_msopatterntype val="26"/>'
            f'<m_colBackground><m_msothmcolidx val="0"/>'
            f'<m_rgb r="FF" g="FF" b="FF"/></m_colBackground>'
            f'</{tag_name}>'
        )
    return (
        f'<{tag_name}><m_bVisible val="1"/><m_msofilltype val="1"/>'
        f'<m_col><m_msothmcolidx val="0"/><m_rgb r="{r_value}" '
        f'g="{g_value}" b="{b_value}"/></m_col></{tag_name}>'
    )


def _patch_named_fill_element(
    block: str,
    tag_name: str,
    style: SegmentStyle,
) -> tuple[str, int]:
    """Patch or materialize one native think-cell fill element."""
    full_pattern = re.compile(
        rf'<{re.escape(tag_name)}>.*?</{re.escape(tag_name)}>',
        flags=re.S,
    )
    full_match = full_pattern.search(block)
    desired_fill = _explicit_tc_fill_xml(tag_name, style)

    if full_match:
        old_fill = full_match.group(0)
        old_is_pattern = '<m_msofilltype val="2"/>' in old_fill
        has_complete_pattern = (
            '<m_msopatterntype' in old_fill
            and '<m_colBackground>' in old_fill
        )
        # Preserve existing native structure where possible.  Switching between
        # solid and pattern uses the canonical complete element.
        if (style.patterned and not (old_is_pattern and has_complete_pattern)) or (
            not style.patterned and old_is_pattern
        ):
            new_fill = desired_fill
        else:
            new_fill, changed = _replace_first_rgb_in_fill(old_fill, style)
            if not changed and _scalar_block_fill_rgb(
                f'<CSequenceChartDataScalar>{new_fill}</CSequenceChartDataScalar>'
            ) == style.hex_color.upper():
                return block, 0
            if '<m_rgb ' not in new_fill:
                new_fill = desired_fill
        if new_fill == old_fill:
            return block, 0
        return (
            block[: full_match.start()]
            + new_fill
            + block[full_match.end() :],
            1,
        )

    empty_pattern = re.compile(rf'<{re.escape(tag_name)}\s*/>')
    empty_match = empty_pattern.search(block)
    if not empty_match:
        return block, 0
    return (
        block[: empty_match.start()]
        + desired_fill
        + block[empty_match.end() :],
        1,
    )


def _patch_scalar_fill_block(block: str, style: SegmentStyle) -> tuple[str, int]:
    # Newly generated high-cardinality series frequently arrive with
    # ``<m_fill/>``.  Materialize those placeholders instead of treating them
    # as harmless blanks; otherwise manual think-cell carry-over restores the
    # template's old blue/grey palette.
    return _patch_named_fill_element(block, "m_fill", style)


def _patch_scalar_by_id(root_text: str, scalar_id: str, style: SegmentStyle) -> tuple[str, int, int]:
    pattern = re.compile(
        rf'<CSequenceChartDataScalar id="{re.escape(str(scalar_id))}">.*?</CSequenceChartDataScalar>',
        flags=re.S,
    )
    changed = 0
    matched = 0

    def replace(match: re.Match) -> str:
        nonlocal changed, matched
        matched += 1
        new_block, patch_count = _patch_scalar_fill_block(match.group(0), style)
        changed += patch_count
        return new_block

    return pattern.sub(replace, root_text), changed, matched

def _vector_scalar_ids_from_tc_root(root_text: str) -> list[list[str]]:
    vectors: list[list[str]] = []
    for match in re.finditer(
        r'<CSequenceChartDataVector id="\d+">(.*?)</CSequenceChartDataVector>',
        root_text,
        flags=re.S,
    ):
        block = match.group(1)
        ocol_match = re.search(r'<ocol\b.*?</ocol>', block, flags=re.S)
        if not ocol_match:
            continue
        scalar_ids = re.findall(r'<elem idref="(\d+)"/>', ocol_match.group(0))
        if scalar_ids:
            vectors.append(scalar_ids)
    # The active document contains two serialized copies. Keep all; scalar IDs
    # are repeated by design, and _patch_scalar_by_id patches every copy.
    return vectors



def _tc_variable_source_text_map(root_text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for match in re.finditer(
        r'<CVariableSource id="(\d+)">(.*?)</CVariableSource>',
        root_text,
        flags=re.S,
    ):
        value_match = re.search(
            r'<m_varval type="5">(.*?)</m_varval>',
            match.group(2),
            flags=re.S,
        )
        if value_match:
            result[match.group(1)] = html.unescape(value_match.group(1))
    return result


def _style_by_segment_key(
    l0_name: str,
    segments: list[dict[str, Any]],
) -> dict[str, SegmentStyle]:
    styles = segment_style_plan(l0_name, segments)
    result: dict[str, SegmentStyle] = {}
    for segment, style in zip(segments, styles):
        key = compact_text(segment_display_name(segment))
        if key:
            result[key] = style
    return result


def _patch_l0_series_fills_in_tc_root(
    root_text: str,
    l0_name: str,
    segments: list[dict[str, Any]],
) -> tuple[str, int, dict[str, Any]]:
    """Materialize native series fills used by think-cell and its legend.

    New series created beyond the original six-row template can serialize with
    self-closing ``m_fill`` / ``m_fillExcel`` elements.  Both are made explicit
    here so the one manual carry-over cannot fall back to the template palette.
    """
    styles = segment_style_plan(l0_name, segments)
    desired_keys = [compact_text(segment_display_name(segment)) for segment in segments]
    style_by_key = {key: style for key, style in zip(desired_keys, styles) if key}
    display_by_key = {
        compact_text(segment_display_name(segment)): segment_display_name(segment)
        for segment in segments
        if compact_text(segment_display_name(segment))
    }
    variable_text = _tc_variable_source_text_map(root_text)
    pattern = re.compile(
        r'<CSequenceChartDataSeries id="(\d+)">(.*?)</CSequenceChartDataSeries>',
        flags=re.S,
    )

    entries: list[dict[str, Any]] = []
    unique_series_ids: list[str] = []
    for match in pattern.finditer(root_text):
        block = match.group(0)
        if '<m_fill' not in block:
            continue
        series_id = match.group(1)
        varsrc_match = re.search(r'<m_varsrc idref="(\d+)"/>', block)
        label = variable_text.get(varsrc_match.group(1), "") if varsrc_match else ""
        label_key = compact_text(label)
        # Ignore the two blank helper series retained by the template; only
        # named ChartData series participate in color and legend mapping.
        if not label_key:
            continue
        entries.append(
            {
                "series_id": series_id,
                "label": label,
                "label_key": label_key,
            }
        )
        if series_id not in unique_series_ids:
            unique_series_ids.append(series_id)

    if not entries:
        raise RuntimeError("No think-cell data-series fill blocks were found.")

    style_by_series_id: dict[str, SegmentStyle] = {}
    key_by_series_id: dict[str, str] = {}
    mapping_method = "series-label"
    matched_keys: set[str] = set()
    for entry in entries:
        key = entry["label_key"]
        style = style_by_key.get(key)
        if style is None:
            continue
        style_by_series_id[entry["series_id"]] = style
        key_by_series_id[entry["series_id"]] = key
        matched_keys.add(key)

    if len(matched_keys) != len(style_by_key):
        mapping_method = "reverse-series-order"
        if len(unique_series_ids) < len(styles):
            raise RuntimeError(
                "Not enough serialized think-cell data series for L0 color mapping: "
                f"series={len(unique_series_ids)}, segments={len(styles)}, "
                f"labels={[entry['label'] for entry in entries]}"
            )
        # Serialized series are bottom-to-top.  ChartData rows are top-to-bottom.
        selected_ids = list(reversed(unique_series_ids[-len(styles) :]))
        style_by_series_id = {
            series_id: style
            for series_id, style in zip(selected_ids, styles)
        }
        key_by_series_id = {
            series_id: key
            for series_id, key in zip(selected_ids, desired_keys)
        }

    changed = 0
    matched_by_series: dict[str, int] = {sid: 0 for sid in style_by_series_id}
    verified_native_by_series: dict[str, int] = {sid: 0 for sid in style_by_series_id}
    verified_excel_by_series: dict[str, int] = {sid: 0 for sid in style_by_series_id}

    def replace(match: re.Match) -> str:
        nonlocal changed
        series_id = match.group(1)
        style = style_by_series_id.get(series_id)
        if style is None:
            return match.group(0)
        matched_by_series[series_id] += 1
        block = match.group(0)
        block, native_changed = _patch_named_fill_element(block, "m_fill", style)
        block, excel_changed = _patch_named_fill_element(block, "m_fillExcel", style)
        changed += native_changed + excel_changed
        return block

    root_text = pattern.sub(replace, root_text)

    def tag_state(block: str, tag_name: str) -> tuple[str, bool]:
        match = re.search(
            rf'<{re.escape(tag_name)}>.*?</{re.escape(tag_name)}>',
            block,
            flags=re.S,
        )
        if not match:
            return "", False
        rgb_match = re.search(
            r'<m_rgb r="([0-9A-F]{2})" g="([0-9A-F]{2})" b="([0-9A-F]{2})"/>',
            match.group(0),
        )
        color = "".join(rgb_match.groups()).upper() if rgb_match else ""
        patterned = (
            '<m_msofilltype val="2"/>' in match.group(0)
            and '<m_msopatterntype' in match.group(0)
            and '<m_colBackground>' in match.group(0)
        )
        return color, patterned

    for match in pattern.finditer(root_text):
        series_id = match.group(1)
        style = style_by_series_id.get(series_id)
        if style is None:
            continue
        native_color, native_patterned = tag_state(match.group(0), "m_fill")
        excel_color, excel_patterned = tag_state(match.group(0), "m_fillExcel")
        expected = style.hex_color.upper()
        if native_color == expected and (not style.patterned or native_patterned):
            verified_native_by_series[series_id] += 1
        if excel_color == expected and (not style.patterned or excel_patterned):
            verified_excel_by_series[series_id] += 1

    bad = {
        series_id: {
            "matched": matched_by_series.get(series_id, 0),
            "native_verified": verified_native_by_series.get(series_id, 0),
            "excel_verified": verified_excel_by_series.get(series_id, 0),
            "expected": style_by_series_id[series_id].hex_color.upper(),
        }
        for series_id in style_by_series_id
        if matched_by_series.get(series_id, 0) <= 0
        or verified_native_by_series.get(series_id, 0)
        != matched_by_series.get(series_id, 0)
        or verified_excel_by_series.get(series_id, 0)
        != matched_by_series.get(series_id, 0)
    }
    if bad:
        raise RuntimeError(
            "L0 native series-fill verification failed: "
            f"method={mapping_method}, failures={bad}"
        )

    return root_text, changed, {
        "series_mapping_method": mapping_method,
        "matched_series_copies": matched_by_series,
        "verified_native_series_copies": verified_native_by_series,
        "verified_excel_series_copies": verified_excel_by_series,
        "verified_series_labels": len(style_by_series_id),
        "series_to_segment": {
            series_id: display_by_key.get(key_by_series_id.get(series_id, ""), "")
            for series_id in style_by_series_id
        },
    }

def _legend_entry_to_series_map(root_text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for match in re.finditer(
        r'<CSequenceChartLegendEntry id="(\d+)">(.*?)</CSequenceChartLegendEntry>',
        root_text,
        flags=re.S,
    ):
        parent_match = re.search(r'<m_featParent idref="(\d+)"/>', match.group(2))
        if parent_match:
            result[match.group(1)] = parent_match.group(1)
    return result


def _series_label_key_map(root_text: str) -> dict[str, str]:
    variable_text = _tc_variable_source_text_map(root_text)
    result: dict[str, str] = {}
    for match in re.finditer(
        r'<CSequenceChartDataSeries id="(\d+)">(.*?)</CSequenceChartDataSeries>',
        root_text,
        flags=re.S,
    ):
        varsrc_match = re.search(r'<m_varsrc idref="(\d+)"/>', match.group(2))
        if varsrc_match:
            result[match.group(1)] = compact_text(
                variable_text.get(varsrc_match.group(1), "")
            )
    return result


def _patch_and_verify_l0_legend_order(
    root_text: str,
    segments: list[dict[str, Any]],
) -> tuple[str, int, dict[str, Any]]:
    desired_keys = [compact_text(segment_display_name(segment)) for segment in segments]
    desired_keys = [key for key in desired_keys if key]
    entry_to_series = _legend_entry_to_series_map(root_text)
    series_to_key = _series_label_key_map(root_text)
    changed = 0
    verified_orders = 0
    observed_orders: list[list[str]] = []

    pattern = re.compile(
        r'(<CSequenceChartLegendSE\b.*?<m_clegendentry length="\d+">)(.*?)(</m_clegendentry>.*?</CSequenceChartLegendSE>)',
        flags=re.S,
    )

    def replace(match: re.Match) -> str:
        nonlocal changed, verified_orders
        inner = match.group(2)
        elem_matches = list(re.finditer(r'<elem idref="(\d+)"/>', inner))
        if not elem_matches:
            return match.group(0)
        current_ids = [item.group(1) for item in elem_matches]
        current_key_by_id = {
            entry_id: series_to_key.get(entry_to_series.get(entry_id, ""), "")
            for entry_id in current_ids
        }
        queues: dict[str, list[str]] = {}
        for entry_id in current_ids:
            key = current_key_by_id.get(entry_id, "")
            queues.setdefault(key, []).append(entry_id)

        desired_ids: list[str] = []
        for key in desired_keys:
            if not queues.get(key):
                return match.group(0)
            desired_ids.append(queues[key].pop(0))
        desired_ids.extend(entry_id for entry_id in current_ids if entry_id not in desired_ids)

        rebuilt_elems = "".join(f'<elem idref="{entry_id}"/>' for entry_id in desired_ids)
        first_start = elem_matches[0].start()
        last_end = elem_matches[-1].end()
        new_inner = inner[:first_start] + rebuilt_elems + inner[last_end:]
        if len(new_inner) != len(inner):
            return match.group(0)
        if desired_ids != current_ids:
            changed += 1

        order = [current_key_by_id.get(entry_id, "") for entry_id in desired_ids[: len(desired_keys)]]
        observed_orders.append(order)
        if order == desired_keys:
            verified_orders += 1
        return match.group(1) + new_inner + match.group(3)

    root_text = pattern.sub(replace, root_text)
    # The visible L0 legend is rebuilt and verified through live PowerPoint COM.
    # Some hidden think-cell copies expose stale template labels, so absence of
    # a resolvable native legend list must not invalidate otherwise-correct
    # native series/scalar colors.
    return root_text, changed, {
        "desired_legend_order": desired_keys,
        "observed_legend_orders": observed_orders,
        "legend_orders_verified": verified_orders,
        "native_legend_order_status": (
            "verified" if verified_orders > 0 else "not-resolved; visible legend retained"
        ),
    }


def _tc_scalar_blocks_by_id(root_text: str) -> dict[str, list[str]]:
    blocks: dict[str, list[str]] = {}
    for match in re.finditer(
        r'<CSequenceChartDataScalar id="(\d+)">.*?</CSequenceChartDataScalar>',
        root_text,
        flags=re.S,
    ):
        blocks.setdefault(match.group(1), []).append(match.group(0))
    return blocks


def _scalar_block_fill_rgb(block: str) -> str:
    fill_match = re.search(r'<m_fill>(.*?)</m_fill>', block, flags=re.S)
    if not fill_match:
        return ""
    rgb_match = re.search(
        r'<m_rgb r="([0-9A-F]{2})" g="([0-9A-F]{2})" b="([0-9A-F]{2})"/>',
        fill_match.group(1),
    )
    return "".join(rgb_match.groups()).upper() if rgb_match else ""


def _segment_scalar_ids_for_vector(
    scalar_ids: list[str],
    styles: list[SegmentStyle],
    scalar_blocks: dict[str, list[str]],
) -> tuple[list[str], str]:
    """Return only the real ChartData segment scalars for a company vector.

    The L0 think-cell table has a stable native contract: every company vector
    starts with exactly two non-segment scalar slots, followed by the stacked
    series written from ChartData row 6 onward.  We must never guess a different
    window when one serialized helper/inactive vector is shorter or sparse.

    The previous fallback could slide the window back to position 0 and treat a
    metadata scalar (for example id 463 in the supplied template family) as a
    chart segment.  That produced the false ``no-populated-fill-copy`` failure
    after all live chart colors had already succeeded.  Short/helper vectors are
    now skipped; they are not chart-data company vectors.
    """
    del scalar_blocks  # Presence/fill state is validated after the contract slice.

    segment_count = len(styles)
    if segment_count <= 0:
        return [], "no-segments"

    contract_start = TC_VECTOR_SEGMENT_START_INDEX
    contract_end = contract_start + segment_count
    if len(scalar_ids) < contract_end:
        return [], f"skip-short-vector:{len(scalar_ids)}<{contract_end}"

    contract_ids = [str(value) for value in scalar_ids[contract_start:contract_end]]
    if len(contract_ids) != segment_count:
        return [], "unresolved-contract-slice"
    return contract_ids, f"contract-slice:{contract_start}:{contract_end}"


def _set_tc_datasheet_fill_on_top(
    root_text: str,
    enabled: bool,
) -> tuple[str, int, dict[str, Any]]:
    """Set think-cell's native datasheet-fill-on-top flag deterministically.

    For the final carry-over state we deliberately keep this setting OFF and
    patch the native segment fills themselves.  This avoids depending on
    whether AddRangeData copied Excel cell fills into every serialized
    think-cell scalar.
    """
    desired = "1" if enabled else "0"
    pattern = re.compile(r'<m_bExcelOnTop val="([01])"/>')
    matched = 0
    changed = 0

    def replace(match: re.Match) -> str:
        nonlocal matched, changed
        matched += 1
        if match.group(1) == desired:
            return match.group(0)
        changed += 1
        return f'<m_bExcelOnTop val="{desired}"/>'

    root_text = pattern.sub(replace, root_text)
    if matched <= 0:
        raise RuntimeError(
            "The think-cell active document has no m_bExcelOnTop setting."
        )
    wrong = len(
        re.findall(
            rf'<m_bExcelOnTop val="{("0" if enabled else "1")}"/>',
            root_text,
        )
    )
    verified = len(
        re.findall(rf'<m_bExcelOnTop val="{desired}"/>', root_text)
    )
    if wrong or verified != matched:
        raise RuntimeError(
            "think-cell datasheet-fill-on-top verification failed: "
            f"enabled={enabled}, matched={matched}, changed={changed}, "
            f"verified={verified}, wrong={wrong}"
        )
    return root_text, changed, {
        "requested_fill_on_top": bool(enabled),
        "matched_fill_on_top_flags": matched,
        "changed_fill_on_top_flags": changed,
        "verified_fill_on_top_flags": verified,
    }


def _enable_tc_datasheet_fill_on_top(
    root_text: str,
) -> tuple[str, int, dict[str, Any]]:
    return _set_tc_datasheet_fill_on_top(root_text, enabled=True)


def _disable_tc_datasheet_fill_on_top(
    root_text: str,
) -> tuple[str, int, dict[str, Any]]:
    return _set_tc_datasheet_fill_on_top(root_text, enabled=False)


def _patch_l0_segment_colors_in_tc_root(
    root_text: str,
    l0_name: str,
    segments: list[dict[str, Any]],
) -> tuple[str, int, dict[str, Any]]:
    if not segments:
        return root_text, 0, {"segments": 0, "vectors": 0}

    styles = segment_style_plan(l0_name, segments)
    vectors = _vector_scalar_ids_from_tc_root(root_text)
    if not vectors:
        raise RuntimeError("No think-cell data vectors were found for L0 color patching.")

    scalar_blocks = _tc_scalar_blocks_by_id(root_text)
    scalar_style_by_id: dict[str, SegmentStyle] = {}
    vector_resolution_methods: dict[str, int] = {}
    eligible_vector_count = 0
    skipped_vectors = 0

    for scalar_ids in vectors:
        data_scalar_ids, method = _segment_scalar_ids_for_vector(
            scalar_ids,
            styles,
            scalar_blocks,
        )
        vector_resolution_methods[method] = vector_resolution_methods.get(method, 0) + 1
        if len(data_scalar_ids) != len(styles):
            skipped_vectors += 1
            continue
        eligible_vector_count += 1
        for scalar_id, style in zip(data_scalar_ids, styles):
            scalar_id = str(scalar_id)
            previous = scalar_style_by_id.get(scalar_id)
            if previous is not None and previous != style:
                raise RuntimeError(
                    f"think-cell scalar {scalar_id} resolved to conflicting L0 colors: "
                    f"{previous} versus {style}."
                )
            scalar_style_by_id[scalar_id] = style

    if eligible_vector_count <= 0 or not scalar_style_by_id:
        raise RuntimeError(
            "No chart segment scalar vectors could be resolved for the native L0 color patch. "
            f"methods={vector_resolution_methods}"
        )

    scalar_changed = 0
    total_matched = 0
    for scalar_id, style in sorted(scalar_style_by_id.items(), key=lambda item: int(item[0])):
        root_text, changed, matched = _patch_scalar_by_id(root_text, scalar_id, style)
        scalar_changed += changed
        total_matched += matched

    # Every selected scalar must now have an explicit native fill.  This is the
    # high-cardinality guarantee: no visible point is allowed to fall back to
    # the template palette during manual carry-over.
    verified_scalar_blocks = 0
    ignored_blank_scalar_blocks = 0
    blank_only_scalar_ids: list[str] = []
    missing_scalar_ids: list[str] = []
    mismatched_scalars: dict[str, list[str]] = {}
    pattern_verified_ids: list[str] = []
    pattern_fallback_ids: list[str] = []
    scalar_blocks_after = _tc_scalar_blocks_by_id(root_text)
    for scalar_id, style in scalar_style_by_id.items():
        blocks = scalar_blocks_after.get(scalar_id, [])
        if not blocks:
            missing_scalar_ids.append(scalar_id)
            continue

        observed_colors: list[str] = []
        pattern_ok = True
        for block in blocks:
            color = _scalar_block_fill_rgb(block)
            observed_colors.append(color or "<blank>")
            if style.patterned:
                pattern_ok = pattern_ok and (
                    '<m_msofilltype val="2"/>' in block
                    and '<m_msopatterntype' in block
                    and '<m_colBackground>' in block
                )

        if any(color != style.hex_color.upper() for color in observed_colors):
            mismatched_scalars[scalar_id] = observed_colors
            continue
        if style.patterned and not pattern_ok:
            mismatched_scalars[scalar_id] = observed_colors + ["pattern-missing"]
            continue

        if style.patterned:
            pattern_verified_ids.append(scalar_id)
        verified_scalar_blocks += len(blocks)

    if missing_scalar_ids or mismatched_scalars:
        raise RuntimeError(
            "Native think-cell explicit segment-fill verification failed: "
            f"missing={missing_scalar_ids}, mismatched={mismatched_scalars}"
        )

    # Keep the series-name fallback colors and internal legend order aligned as
    # well. The native scalar fills are authoritative because ExcelOnTop is disabled.
    root_text, series_changed, series_details = _patch_l0_series_fills_in_tc_root(
        root_text,
        l0_name,
        segments,
    )
    root_text, legend_changed, legend_details = _patch_and_verify_l0_legend_order(
        root_text,
        segments,
    )

    details = {
        "segments": len(styles),
        "vectors": eligible_vector_count,
        "skipped_vectors": skipped_vectors,
        "vector_resolution_methods": vector_resolution_methods,
        "unique_segment_scalars": len(scalar_style_by_id),
        "matched_scalar_blocks": total_matched,
        "verified_scalar_blocks": verified_scalar_blocks,
        "ignored_blank_scalar_blocks": ignored_blank_scalar_blocks,
        "blank_only_segment_scalar_ids": blank_only_scalar_ids,
        "missing_segment_scalar_ids": missing_scalar_ids,
        "pattern_verified_scalar_ids": pattern_verified_ids,
        "pattern_fallback_scalar_ids": pattern_fallback_ids,
        "scalar_color_patches": scalar_changed,
        "series_color_patches": series_changed,
        "legend_order_patches": legend_changed,
        "style_plan": segment_style_log_items(l0_name, segments),
        **series_details,
        **legend_details,
    }
    return root_text, scalar_changed + series_changed + legend_changed, details

def _select_l0_quartile_line_candidates(
    original_infos: list[dict[str, Any]],
    rank: str,
    excluded_line_ids: set[str] | None = None,
) -> tuple[list[dict[str, Any]], str]:
    """Fallback resolver used only when a captured pre-color anchor is absent."""
    excluded_line_ids = {str(value) for value in (excluded_line_ids or set())}
    grouped = _group_tc_value_line_infos(original_infos)
    grouped = {
        key: group
        for key, group in grouped.items()
        if key[0] not in excluded_line_ids
    }
    source_value = float(L0_QUARTILE_PLACEHOLDERS_BY_RANK[rank])
    expected_colors = {
        color.upper()
        for color in L0_QUARTILE_LINE_COLORS_BY_RANK.get(rank, set())
    }

    placeholder_groups = [
        (key, group)
        for key, group in grouped.items()
        if any(
            int(round(float(info["var_value"]))) == int(round(source_value))
            for info in group
        )
    ]
    if placeholder_groups:
        color_matches = [
            item
            for item in placeholder_groups
            if any(
                str(info.get("line_color_hex") or "").upper() in expected_colors
                for info in item[1]
            )
        ]
        candidates = color_matches or placeholder_groups
        key, group = min(
            candidates,
            key=lambda item: _quartile_group_sort_value(item[1], source_value),
        )
        return group, f"template-placeholder:{int(round(source_value))}"

    color_groups = [
        (key, group)
        for key, group in grouped.items()
        if any(
            str(info.get("line_color_hex") or "").upper() in expected_colors
            for info in group
        )
    ]
    if color_groups:
        key, group = min(
            color_groups,
            key=lambda item: _quartile_group_sort_value(item[1], source_value),
        )
        colors = sorted(
            {
                str(info.get("line_color_hex") or "").upper()
                for info in group
                if info.get("line_color_hex")
            }
        )
        return group, "line-color-nearest-template:" + "/".join(colors)

    geometry_groups = [
        (key, group)
        for key, group in grouped.items()
        if any(info.get("label_center_y") is not None for info in group)
    ]
    ordered = sorted(
        geometry_groups,
        key=lambda item: sum(
            float(info["label_center_y"])
            for info in item[1]
            if info.get("label_center_y") is not None
        )
        / max(
            1,
            sum(
                1
                for info in item[1]
                if info.get("label_center_y") is not None
            ),
        ),
    )
    if len(ordered) >= 3:
        index = {"highest": 0, "middle": len(ordered) // 2, "lowest": -1}[rank]
        return ordered[index][1], "geometry-order"

    return [], "unresolved"


def _patch_l0_quartiles_in_tc_root_all_copies(
    root_text: str,
    quartile_values: list[float],
    quartile_anchor: dict[str, dict[str, Any]] | None = None,
) -> tuple[str, int, list[dict[str, Any]]]:
    targets = _build_l0_quartile_targets(quartile_values)
    if len(targets) < 3:
        raise ValueError("Fewer than three usable L0 quartile values were supplied.")

    axis_mapping = _extract_internal_axis_mapping_from_tc_root(root_text)
    if axis_mapping is None:
        raise ValueError("Could not derive the think-cell internal axis mapping.")
    axis_a, axis_b, axis_points = axis_mapping

    original_infos = _get_tc_value_line_infos(root_text)
    target_by_rank = {target.rank: target for target in targets}
    selected_by_rank: dict[str, tuple[list[dict[str, Any]], str]] = {}
    used_line_ids: set[str] = set()

    # Resolve all three identities from the same untouched root before changing
    # any values or geometry.  Normal full builds use the exact IDs captured
    # from the per-slide PPTX immediately after AddRangeData and before the live
    # color/legend pass.
    for rank in ("highest", "middle", "lowest"):
        candidates: list[dict[str, Any]] = []
        selection_method = ""
        anchor_entry = (quartile_anchor or {}).get(rank) or {}
        anchor_line_id = str(anchor_entry.get("line_id") or "")
        anchor_varsrc_id = str(anchor_entry.get("varsrc_id") or "")

        if anchor_line_id and anchor_varsrc_id:
            candidates = [
                info
                for info in original_infos
                if str(info["line_id"]) == anchor_line_id
                and str(info["varsrc_id"]) == anchor_varsrc_id
            ]
            if candidates:
                selection_method = "captured-pre-color-anchor"

        if not candidates and anchor_line_id:
            line_only = [
                info
                for info in original_infos
                if str(info["line_id"]) == anchor_line_id
                and str(info["line_id"]) not in used_line_ids
            ]
            line_only_keys = {
                (str(info["line_id"]), str(info["varsrc_id"]))
                for info in line_only
            }
            if len(line_only_keys) == 1:
                candidates = line_only
                selection_method = "captured-line-id-anchor"

        if not candidates:
            candidates, selection_method = _select_l0_quartile_line_candidates(
                original_infos,
                rank,
                excluded_line_ids=used_line_ids,
            )

        if not candidates:
            available = [
                {
                    "line_id": info.get("line_id"),
                    "varsrc_id": info.get("varsrc_id"),
                    "value": info.get("var_value"),
                    "color": info.get("line_color_hex"),
                    "center_y": info.get("label_center_y"),
                }
                for info in original_infos
            ]
            raise ValueError(
                f"Could not identify the L0 {rank} quartile line from the captured "
                f"pre-color anchor or fallback rules. Available value lines: {available}"
            )

        candidate_line_ids = {str(info["line_id"]) for info in candidates}
        overlap = candidate_line_ids & used_line_ids
        if overlap:
            raise RuntimeError(
                f"L0 quartile line identification overlapped across ranks even after "
                f"using captured anchors: {sorted(overlap)}"
            )
        used_line_ids.update(candidate_line_ids)
        selected_by_rank[rank] = (candidates, selection_method)

    details: list[dict[str, Any]] = []
    total_patches = 0
    target_y_by_rank: dict[str, int] = {}

    for rank in ("highest", "middle", "lowest"):
        source_value = L0_QUARTILE_PLACEHOLDERS_BY_RANK[rank]
        target = target_by_rank[rank]
        candidates, selection_method = selected_by_rank[rank]

        varsrc_ids = sorted({str(info["varsrc_id"]) for info in candidates})
        textvar_ids = sorted(
            {
                str(info["textvar_id"])
                for info in candidates
                if info.get("textvar_id")
            }
        )
        line_ids = sorted({str(info["line_id"]) for info in candidates})
        line_shape_ids = sorted(
            {
                str(shape_id)
                for info in candidates
                for shape_id in info.get("line_shape_ids", [])
            }
        )
        half_height = max(
            int(info.get("label_half_height", 48)) for info in candidates
        )
        target_y = int(round(axis_a + axis_b * float(target.value)))
        target_y_by_rank[rank] = target_y

        rank_patches = 0
        matched_components = 0
        for varsrc_id in varsrc_ids:
            root_text, changed, matched = _patch_all_variable_source_copies(
                root_text,
                varsrc_id,
                target.value,
            )
            rank_patches += changed
            matched_components += matched
        for textvar_id in textvar_ids:
            root_text, changed, matched = _patch_all_text_variable_copies(
                root_text,
                textvar_id,
                target.label,
            )
            rank_patches += changed
            matched_components += matched
        for shape_id in line_shape_ids:
            root_text, changed, matched = _patch_all_line_shape_copies(
                root_text,
                shape_id,
                target_y,
            )
            rank_patches += changed
            matched_components += matched
        for line_id in line_ids:
            root_text, changed, matched = _patch_all_value_indicator_line_copies(
                root_text,
                line_id,
                target_y,
                half_height,
            )
            rank_patches += changed
            matched_components += matched

        updated_infos = _get_tc_value_line_infos(root_text)
        updated_copy_count = sum(
            1
            for info in updated_infos
            if str(info["line_id"]) in line_ids
            and str(info["varsrc_id"]) in varsrc_ids
            and math.isclose(
                float(info["var_value"]),
                float(target.value),
                rel_tol=0.0,
                abs_tol=1e-7,
            )
        )
        if updated_copy_count != len(candidates):
            raise RuntimeError(
                f"Only {updated_copy_count} of {len(candidates)} serialized {rank} "
                "L0 value-line copies received the target value."
            )

        geometry_details = _verify_l0_quartile_geometry(
            root_text=root_text,
            rank=rank,
            target_y=target_y,
            line_shape_ids=line_shape_ids,
            line_ids=line_ids,
        )

        details.append(
            {
                "rank": rank,
                "source_value": source_value,
                "selection_method": selection_method,
                "observed_line_colors": sorted(
                    {
                        str(info.get("line_color_hex") or "").upper()
                        for info in candidates
                        if info.get("line_color_hex")
                    }
                ),
                "observed_source_values": sorted(
                    {float(info["var_value"]) for info in candidates}
                ),
                "target_value": float(target.value),
                "target_label": target.label,
                "target_y": target_y,
                "serialized_line_copies": len(candidates),
                "verified_line_copies": updated_copy_count,
                "matched_components": matched_components,
                "patches": rank_patches,
                "axis_points": axis_points,
                "line_ids": line_ids,
                "varsrc_ids": varsrc_ids,
                "verified_line_shape_count": sum(
                    len(values)
                    for values in geometry_details["line_shape_centers"].values()
                ),
                "verified_label_rect_count": sum(
                    len(values)
                    for values in geometry_details["label_centers"].values()
                ),
                "connector_rect_count": sum(
                    len(values)
                    for values in geometry_details["connector_centers"].values()
                ),
                "connector_offsets": geometry_details["connector_offsets"],
            }
        )
        total_patches += rank_patches

    distinct_target_y = {target_y_by_rank[rank] for rank in target_y_by_rank}
    distinct_target_values = {
        round(float(target_by_rank[rank].value), 9)
        for rank in target_y_by_rank
    }
    if len(distinct_target_values) == 3 and len(distinct_target_y) != 3:
        raise RuntimeError(
            f"Distinct L0 quartile values mapped to overlapping line coordinates: "
            f"{target_y_by_rank}"
        )

    return root_text, total_patches, details


def _patch_l0_ole_bytes_quartiles_and_colors(
    data: bytes,
    quartile_values: list[float],
    l0_name: str,
    segments: list[dict[str, Any]],
    quartile_anchor: dict[str, dict[str, Any]] | None = None,
) -> tuple[bytes, int, int, list[dict[str, Any]], list[dict[str, Any]]]:
    """Patch the live think-cell XML stream, never raw OLE/slack sectors.

    A CFB file can retain stale copies of an older XML stream in unallocated
    sectors.  The former whole-blob regex saw those stale bytes as additional
    vectors/scalars, which is why a correctly patched scalar could later appear
    to have its neighbour's color during verification.  This implementation
    resolves the active directory entry named ``think-cellXML`` and mutates only
    that stream.
    """
    compound, root_records = _active_thinkcell_root_records_from_ole(data)
    if not root_records:
        raise ValueError(
            "No active think-cell chart XML stream was found in the OLE object."
        )

    use_anchor = quartile_anchor
    selected_records = list(root_records)
    if quartile_anchor:
        match_counts = [
            _root_anchor_match_count(record["root_text"], quartile_anchor)
            for record in root_records
        ]
        best_count = max(match_counts, default=0)
        if best_count > 0:
            selected_records = [
                record
                for record, count in zip(root_records, match_counts)
                if count == best_count
            ]
        else:
            use_anchor = None

    stream_buffers: dict[int, bytearray] = {}
    stream_entries: dict[int, dict[str, Any]] = {}
    for record in root_records:
        entry = record["entry"]
        entry_id = int(entry["id"])
        stream_entries[entry_id] = entry
        stream_buffers.setdefault(entry_id, bytearray(record["stream_bytes"]))

    total_quartile_patches = 0
    total_color_patches = 0
    quartile_details: list[dict[str, Any]] = []
    color_details: list[dict[str, Any]] = []

    # Process roots from the end of each stream toward the beginning because
    # materializing empty native fills can increase XML length.
    selected_records = sorted(
        selected_records,
        key=lambda item: (int(item["entry"]["id"]), int(item["root_start"])),
        reverse=True,
    )

    for record in selected_records:
        entry = record["entry"]
        entry_id = int(entry["id"])
        stream_buffer = stream_buffers[entry_id]
        root_start = int(record["root_start"])
        root_end = int(record["root_end"])
        old_root_bytes = bytes(stream_buffer[root_start:root_end])
        root_text, encoding = _decode_thinkcell_root(old_root_bytes)

        root_text, fill_mode_changed, fill_mode_details = (
            _disable_tc_datasheet_fill_on_top(root_text)
        )
        root_text, native_color_patches, native_color_details = (
            _patch_l0_segment_colors_in_tc_root(
                root_text,
                l0_name,
                segments,
            )
        )

        # Quartile state is deliberately the final mutation.  PowerPoint must
        # not reopen or resave the deck after this package-level operation.
        root_text, quartile_patches, one_quartile_details = (
            _patch_l0_quartiles_in_tc_root_all_copies(
                root_text,
                quartile_values,
                quartile_anchor=use_anchor,
            )
        )

        new_root_bytes = root_text.encode(encoding)
        stream_buffer[root_start:root_end] = new_root_bytes

        total_color_patches += fill_mode_changed + native_color_patches
        total_quartile_patches += quartile_patches
        quartile_details.extend(one_quartile_details)
        color_details.append(
            {
                "l0": l0_name,
                "compound_stream": {
                    "entry_id": entry_id,
                    "name": entry.get("name", ""),
                    "size": int(entry.get("size", 0)),
                },
                "fill_on_top": fill_mode_details,
                "native_color_mode": "active think-cellXML scalar and series fills",
                **native_color_details,
            }
        )

    for entry_id, stream_buffer in stream_buffers.items():
        compound.write_stream_resized(
            stream_entries[entry_id],
            bytes(stream_buffer),
        )

    patched_data = compound.to_bytes()
    # Reparse the active stream after FAT/directory updates.  This catches any
    # CFB-chain error before the PowerPoint package is published.
    _verified_compound, verified_records = _active_thinkcell_root_records_from_ole(
        patched_data
    )
    if not verified_records:
        raise RuntimeError(
            "The resized active think-cell XML stream could not be reopened."
        )
    if not quartile_details:
        raise ValueError(
            "No quartile value-line root was patched in the active think-cell XML stream."
        )
    if not color_details:
        raise ValueError(
            "No chart-data root was patched in the active think-cell XML stream."
        )
    return (
        patched_data,
        total_quartile_patches,
        total_color_patches,
        quartile_details,
        color_details,
    )


def _normalize_l0_slide_indices(
    total_slide_count: int,
    record_count: int,
    slide_indices: Iterable[int] | None,
) -> list[int]:
    if slide_indices is None:
        if total_slide_count != record_count:
            raise RuntimeError(
                f"L0-only deck contains {total_slide_count} slides but "
                f"{record_count} L0 records were generated. Supply explicit final-deck "
                "slide indices when patching a consolidated presentation."
            )
        return list(range(1, record_count + 1))

    result = [int(value) for value in slide_indices]
    if len(result) != record_count:
        raise ValueError(
            f"Received {len(result)} L0 slide indices for {record_count} records."
        )
    if len(set(result)) != len(result):
        raise ValueError(f"L0 slide indices must be unique: {result}")
    invalid = [value for value in result if value < 1 or value > total_slide_count]
    if invalid:
        raise IndexError(
            f"L0 slide indices {invalid} are outside the final deck's 1..{total_slide_count} range."
        )
    return result


def patch_l0_slides_in_deck(
    output_ppt: Path,
    slide_records: list[dict[str, Any]],
    logger,
    *,
    slide_indices: Iterable[int] | None = None,
) -> None:
    """Patch only the designated L0 slides in a closed PowerPoint package.

    The method is valid for both an L0-only section deck and the final
    consolidated deck.  It resolves each target slide's own OLE relationship,
    patches the active ``think-cellXML`` stream, replaces the package atomically,
    and never opens PowerPoint.  No PowerPoint save is allowed afterward.
    """
    output_ppt = Path(output_ppt).resolve()
    _wait_for_readable_pptx(output_ppt, logger)
    temp_path = output_ppt.with_suffix(output_ppt.suffix + ".l0_native_patching")
    if temp_path.exists():
        temp_path.unlink()

    try:
        with zipfile.ZipFile(output_ppt, "r") as source_archive:
            slide_parts = _ordered_slide_parts(source_archive)
            target_indices = _normalize_l0_slide_indices(
                total_slide_count=len(slide_parts),
                record_count=len(slide_records),
                slide_indices=slide_indices,
            )
            target_slides = [
                (slide_index, slide_parts[slide_index - 1], slide_record)
                for slide_index, slide_record in zip(target_indices, slide_records)
            ]

            embedding_plan: dict[str, tuple[int, dict[str, Any]]] = {}
            for slide_index, slide_part, slide_record in target_slides:
                embedding_parts = _slide_thinkcell_embedding_parts(
                    source_archive,
                    slide_part,
                )
                if not embedding_parts:
                    raise RuntimeError(
                        f"Slide {slide_index} ({slide_record['l0']}) has no "
                        "think-cell active-document OLE relationship."
                    )
                for embedding_part in embedding_parts:
                    existing = embedding_plan.get(embedding_part)
                    if existing and existing[0] != slide_index:
                        raise RuntimeError(
                            f"think-cell embedding {embedding_part} is shared by multiple L0 slides."
                        )
                    embedding_plan[embedding_part] = (slide_index, slide_record)

            patched_embeddings: set[str] = set()
            with zipfile.ZipFile(temp_path, "w") as target_archive:
                for info in source_archive.infolist():
                    data = source_archive.read(info.filename)
                    plan = embedding_plan.get(info.filename)
                    if plan is not None:
                        slide_index, slide_record = plan
                        (
                            data,
                            patch_count,
                            native_color_patch_count,
                            details,
                            native_color_details,
                        ) = _patch_l0_ole_bytes_quartiles_and_colors(
                            data,
                            slide_record["quartile_values"],
                            l0_name=slide_record["l0"],
                            segments=slide_record["segments"],
                            quartile_anchor=slide_record.get("quartile_anchor"),
                        )
                        serialized_copies = {
                            item["rank"]: item["serialized_line_copies"]
                            for item in details
                        }
                        verified_copies = {
                            item["rank"]: item["verified_line_copies"]
                            for item in details
                        }
                        verified_geometry = {
                            item["rank"]: {
                                "line_shapes": item.get("verified_line_shape_count", 0),
                                "visible_label_rects": item.get("verified_label_rect_count", 0),
                                "connector_rects": item.get("connector_rect_count", 0),
                                "connector_offsets": item.get("connector_offsets", {}),
                                "target_y": item.get("target_y"),
                            }
                            for item in details
                        }
                        resolution = {
                            item["rank"]: {
                                "method": item.get("selection_method"),
                                "source_values": item.get("observed_source_values"),
                                "line_colors": item.get("observed_line_colors"),
                            }
                            for item in details
                        }
                        logger.info(
                            "Final-deck L0 native-color + quartile patch: slide %s %s; embedding=%s; "
                            "quartile_patches=%s; serialized copies=%s; verified=%s; "
                            "geometry=%s; resolved_by=%s; native_color_patches=%s; "
                            "native_color_details=%s",
                            slide_index,
                            slide_record["l0"],
                            info.filename,
                            patch_count,
                            serialized_copies,
                            verified_copies,
                            verified_geometry,
                            resolution,
                            native_color_patch_count,
                            native_color_details,
                        )
                        if patch_count <= 0 and native_color_patch_count <= 0:
                            raise RuntimeError(
                                f"No native L0 color or quartile bytes changed for slide {slide_index} "
                                f"({slide_record['l0']})."
                            )
                        patched_embeddings.add(info.filename)
                    target_archive.writestr(_zip_info_copy(info), data)

            missing = set(embedding_plan) - patched_embeddings
            if missing:
                raise RuntimeError(
                    "The following final-deck think-cell embeddings were not patched: "
                    + ", ".join(sorted(missing))
                )

        _replace_file_with_retry(temp_path, output_ppt)
        _wait_for_readable_pptx(output_ppt, logger)
        logger.info(
            "Final deck now contains verified active-stream L0 native segment-color, native legend, "
            "and quartile carry-over data on slide(s) %s; every selected scalar/series fill is "
            "explicit, datasheet fill-on-top is disabled, and the deck was not reopened or resaved "
            "after the final patch.",
            target_indices,
        )
    except Exception:
        try:
            temp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise


def patch_final_l0_deck_quartiles(
    output_ppt: Path,
    slide_records: list[dict[str, Any]],
    logger,
) -> None:
    """Backward-compatible wrapper for an L0-only deck."""
    patch_l0_slides_in_deck(
        output_ppt=output_ppt,
        slide_records=slide_records,
        logger=logger,
    )

