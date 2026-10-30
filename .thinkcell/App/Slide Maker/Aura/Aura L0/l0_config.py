from __future__ import annotations

"""Aura L0-only paths, template contract, constants, and data models.

This module deliberately contains no PowerPoint/Excel COM startup.  It can be
imported by workbook readers, the consolidated workflow, and CLI help without
starting Office.  All Aura L0 behavior remains scoped to the Aura L0 folder.
"""

from dataclasses import dataclass
from pathlib import Path
import re
import sys
from typing import Any

from openpyxl.styles import PatternFill


L0_DIR = Path(__file__).resolve().parent
AURA_ROOT = L0_DIR.parent
APP_ROOT = AURA_ROOT.parent
AURA_L1_DIR = AURA_ROOT / "Aura L1"

# Make the L0 modules independent of the current PowerShell working directory.
for import_dir in (APP_ROOT, AURA_L1_DIR, L0_DIR):
    import_text = str(import_dir)
    if import_text not in sys.path:
        sys.path.insert(0, import_text)


DEFAULT_TEMPLATE_PPT = L0_DIR / "ThinkCell_L0.pptx"
DEFAULT_OUTPUT_DIR = L0_DIR / "output"
DEFAULT_OUTPUT_NAME = "Generated_Aura_L0_Test.pptx"
L0_BUILD_ID = "l0-dynamic-ten-shades-lightest-overflow-active-stream-v2026-07-31"

CHART_NAME = "fte_chart"
CHART_HEADER_ROW = 1
REVENUE_CHART_ROW = 2
FIRST_SEGMENT_ROW = 6
REVENUE_ROW_LABEL = "Revenue ($M)"
TC_VECTOR_SEGMENT_START_INDEX = 2

L0_QUARTILE_PLACEHOLDERS_BY_RANK = {
    "highest": 1460,
    "middle": 1085,
    "lowest": 950,
}

L0_QUARTILE_LINE_COLORS_BY_RANK = {
    "highest": {"CC0000", "990000"},
    "middle": {"E9CD49", "C6AA3D"},
    "lowest": {"507867", "104C3E"},
}

BAIN_TOOLBOX_FAMILIES = {
    "graphite": ["333333", "5C5C5C", "858585", "B4B4B4", "D6D6D6"],
    "sky": ["2D475A", "46647B", "7891AA", "A3BCD3", "DCE5EA"],
    "forest": ["104C3E", "507867", "83AC9A", "BBCABA", "DCE2D6"],
    "berry": ["640A40", "973B74", "BA749F", "D9ABC6", "EED6E5"],
    "sunset": ["AB8933", "C6AA3D", "E9CD49", "F2DE8A", "FAEEC3"],
}

MAX_FAMILY_COLORED_SEGMENTS = 10

L0_COLOR_FAMILY_BY_KEY = {
    "ga": "forest",
    "ganda": "forest",
    "sm": "sky",
    "salesmarketing": "sky",
    "salesandmarketing": "sky",
    "ops": "sunset",
    "operations": "sunset",
    "rd": "berry",
    "randd": "berry",
    "researchdevelopment": "berry",
    "researchanddevelopment": "berry",
    "oos": "graphite",
    "unmapped": "graphite",
}

OTHER_SEGMENT_PATTERN_BY_L0_KEY = {
    "ga": "otherga",
    "ganda": "otherga",
    "ops": "otherops",
    "operations": "otherops",
}

PATTERN_SHADE_INDEX_BY_FAMILY = {
    "graphite": 2,
    "sky": 2,
    "forest": 2,
    "berry": 2,
    "sunset": 2,
}

MSO_SHAPE_RECTANGLE = 1
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_ANCHOR_MIDDLE = 3
PP_ALIGN_LEFT = 1
MSO_PATTERN_WIDE_UPWARD_DIAGONAL = 18

L0_LEGEND_LEFT = 260.0
L0_LEGEND_RIGHT = 710.0
L0_LEGEND_TOP_MIN = 100.0
L0_LEGEND_BOTTOM_MAX = 158.5
L0_LEGEND_MARKER_WIDTH = 10.5
L0_LEGEND_MARKER_HEIGHT = 8.0
L0_LEGEND_MARKER_TEXT_GAP = 4.0
L0_LEGEND_ENTRY_GAP = 9.0
L0_LEGEND_NAME_PREFIX = "AURA_L0_LEGEND"

EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
PPT_SUFFIXES = {".pptx", ".pptm"}

TOTAL_MARKER_PATTERN = re.compile(
    r"\b(total|overall|aggregate|aggregated|l0|level\s*0|level\s*zero|benchmark)\b",
    flags=re.I,
)

GENERIC_L0_ROW_TEXT_KEYS = {
    "total",
    "overall",
    "aggregate",
    "aggregated",
    "l0",
    "level0",
    "levelzero",
    "benchmark",
    "benchmarks",
    "quartile",
    "quartiles",
    "topq",
    "median",
    "bottomq",
    "fteperbrevenue",
    "fteperrevenue",
}

KNOWN_L0_ALIASES = {
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
    "oos": "OOS",
    "outofscope": "OOS",
    "unmapped": "Unmapped",
}

L0_FIRST_ROW_EXCEPTIONS = {
    "ga": {"otherga"},
    "ops": {"otherops"},
}


def column_letter(col_num: int) -> str:
    letters = ""
    while col_num:
        col_num, remainder = divmod(col_num - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


@dataclass(frozen=True)
class SegmentStyle:
    family: str
    hex_color: str
    patterned: bool = False

    @property
    def rgb(self) -> tuple[str, str, str]:
        color = self.hex_color.strip().lstrip("#").upper()
        return color[0:2], color[2:4], color[4:6]

    @property
    def excel_fill(self) -> PatternFill:
        if self.patterned:
            return PatternFill(
                patternType="darkUp",
                fgColor=self.hex_color,
                bgColor="FFFFFF",
            )
        return PatternFill(
            patternType="solid",
            fgColor=self.hex_color,
            bgColor=self.hex_color,
        )


@dataclass
class DirectL0QuartileRow:
    l0: str
    quartiles: list[float]
    sheet_name: str
    row: int
    header_row: int
    quartile_columns: tuple[int, int, int]
    score: float
    inside_main_l1_table: bool
    label_texts: list[str]

    @property
    def source_description(self) -> str:
        cols = "/".join(column_letter(col) for col in self.quartile_columns)
        return f"{self.sheet_name}!row {self.row} ({cols}, header row {self.header_row})"


@dataclass
class L0QuartileTarget:
    rank: str
    value: float
    label: str


@dataclass
class L0BuildResult:
    deck_path: Path
    slide_records: list[dict[str, Any]]
    client_company: str
    native_patch_pending: bool

    @property
    def slide_count(self) -> int:
        return len(self.slide_records)
