from __future__ import annotations

import logging
import re
import sys
from collections import OrderedDict
from copy import deepcopy
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

try:
    from pptx.enum.text import MSO_VERTICAL_ANCHOR
except Exception:
    MSO_VERTICAL_ANCHOR = None

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

SUMMARY_DIR = Path(__file__).parent.resolve()
AURA_ROOT = SUMMARY_DIR.parent.resolve()
AURA_L1_DIR = AURA_ROOT / "Aura L1"

# Import priority:
#   Aura root     -> shared.py / input.py
#   Aura L1       -> fte_quartile_table_detector.py
#   Aura Summary  -> this workflow's files
for path in [str(AURA_ROOT), str(AURA_L1_DIR), str(SUMMARY_DIR)]:
    while path in sys.path:
        sys.path.remove(path)

sys.path.insert(0, str(AURA_ROOT))
sys.path.insert(0, str(AURA_L1_DIR))
sys.path.insert(0, str(SUMMARY_DIR))

from shared import BuildConfig, clean_text, require_file, to_number_or_none  # noqa: E402
from fte_quartile_table_detector import read_source_data  # noqa: E402


L0_SECTION_ORDER = ["G&A", "S&M", "Ops", "R&D"]

L0_DISPLAY_MAP = {
    "ga": "G&A",
    "ganda": "G&A",
    "generalandadministrative": "G&A",
    "generaladministrative": "G&A",
    "generaladmin": "G&A",
    "sm": "S&M",
    "salesmarketing": "S&M",
    "salesandmarketing": "S&M",
    "ops": "Ops",
    "operations": "Ops",
    "rd": "R&D",
    "randd": "R&D",
    "researchdevelopment": "R&D",
    "researchanddevelopment": "R&D",
}

BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
LIGHT_GRID_GRAY = RGBColor(190, 190, 190)

# Legend colors from the Aura Summary slide:
#   Top / Second quartile: blue
#   Third / Bottom quartile: gold
LEGEND_BLUE = RGBColor(0x46, 0x64, 0x7B)
LEGEND_GOLD = RGBColor(0xAB, 0x89, 0x33)


# -----------------------------------------------------------------------------
# Basic text / number helpers
# -----------------------------------------------------------------------------


def compact_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def normalize_l0(value: Any) -> str:
    raw = clean_text(value)
    compact = compact_text(raw)
    return L0_DISPLAY_MAP.get(compact, raw)


def number_or_none(value: Any) -> float | None:
    return to_number_or_none(value)


# -----------------------------------------------------------------------------
# Aura calculations
# -----------------------------------------------------------------------------


def client_fte_value(record: dict[str, Any]) -> float | None:
    fte_values = record.get("fte_values") or []
    if not fte_values:
        return None
    return number_or_none(fte_values[0])


def sorted_quartile_values(record: dict[str, Any]) -> tuple[float, float, float] | None:
    quartile_values = record.get("quartile_values") or []
    quartiles = sorted(value for value in (number_or_none(v) for v in quartile_values) if value is not None)

    if len(quartiles) < 3:
        return None

    # Lower FTE/$B is better:
    # lowest value  = top quartile threshold
    # middle value  = median
    # highest value = bottom quartile threshold
    top_q, median, bottom_q = quartiles[0], quartiles[1], quartiles[2]
    return top_q, median, bottom_q


def classify_client_quartile(record: dict[str, Any]) -> str:
    client_value = client_fte_value(record)
    quartiles = sorted_quartile_values(record)

    if client_value is None or quartiles is None:
        return "Unknown"

    top_q, median, bottom_q = quartiles

    if client_value <= top_q:
        return "Top"
    if client_value <= median:
        return "Second"
    if client_value <= bottom_q:
        return "Third"
    return "Bottom"


def gap_to_quartile_text(record: dict[str, Any], quartile_name: str) -> str:
    """
    Gap formula:
      (Client FTE/$B - Quartile) / Client FTE/$B

    For Aura Summary we show only:
      - Gap to median
      - Gap to top quartile

    Edge cases:
      - Missing client value
      - Missing quartile value
      - Client value <= 0
      - Gap <= 0
      - Gap rounds to 0%
    all display as "-".
    """
    client_value = client_fte_value(record)
    quartiles = sorted_quartile_values(record)

    if client_value is None or quartiles is None:
        return "-"
    if client_value <= 0:
        return "-"

    top_q, median, _bottom_q = quartiles

    if quartile_name == "top":
        target_value = top_q
    elif quartile_name == "median":
        target_value = median
    else:
        return "-"

    gap = (client_value - target_value) / client_value

    if gap <= 0:
        return "-"

    pct = gap * 100.0
    if round(pct) <= 0:
        return "-"

    return f"{pct:.0f}%"


def quartile_cell_text(label: str) -> str:
    label = clean_text(label)

    if not label or label == "Unknown":
        return "-"

    if label == "Top":
        return "Top Quartile"
    if label == "Second":
        return "Second Quartile"
    if label == "Third":
        return "Third Quartile"
    if label == "Bottom":
        return "Bottom Quartile"

    return f"{label} Quartile"


def quartile_text_style(quartile_text: str) -> tuple[RGBColor, bool]:
    """
    Summary-slide quartile formatting:
      Top Quartile     -> legend blue + bold
      Second Quartile  -> legend blue
      Third Quartile   -> legend gold
      Bottom Quartile  -> legend gold + bold
    """
    compact = compact_text(quartile_text)

    if compact in {"topquartile", "topq"}:
        return LEGEND_BLUE, True

    if compact in {"secondquartile", "secondq", "2ndquartile", "q2"}:
        return LEGEND_BLUE, False

    if compact in {"thirdquartile", "thirdq", "3rdquartile", "q3"}:
        return LEGEND_GOLD, False

    if compact in {"bottomquartile", "bottomq"}:
        return LEGEND_GOLD, True

    return BLACK, False


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


def read_aura_l1_records(
    excel_path: Path,
    template_ppt: Path,
    output_ppt: Path,
    logger: logging.Logger,
    sheet_name: str | None = None,
) -> tuple[str, list[dict[str, Any]]]:
    """
    Use the existing Aura L1 detector so Summary and L1 slides use identical
    source parsing.
    """
    config = BuildConfig(
        excel_path=excel_path.resolve(),
        template_ppt=template_ppt.resolve(),
        output_dir=output_ppt.parent.resolve(),
        output_ppt_name=output_ppt.name,
        sheet_name=sheet_name,
    )

    company_headers, _revenue_values, records = read_source_data(config, logger)
    client_name = clean_text(company_headers[0]) if company_headers else "Client"
    return client_name or "Client", records


# -----------------------------------------------------------------------------
# PowerPoint table helpers
# -----------------------------------------------------------------------------


def iter_all_shapes(shapes):
    for shape in shapes:
        yield shape
        try:
            if hasattr(shape, "shapes"):
                yield from iter_all_shapes(shape.shapes)
        except Exception:
            pass


def replace_text_value(text: str, replacements: dict[str, str]) -> str:
    new_text = text
    for old, new in replacements.items():
        new_text = re.sub(re.escape(old), new, new_text, flags=re.I)
    return new_text


def replace_non_table_placeholders(prs: Presentation, replacements: dict[str, str]) -> None:
    """
    Replace slide-level placeholders but intentionally skip tables.

    The table header is formatted separately so the client quartile header does
    not inherit placeholder formatting.
    """
    for slide in prs.slides:
        for shape in iter_all_shapes(slide.shapes):
            try:
                if getattr(shape, "has_table", False):
                    continue
                if getattr(shape, "has_text_frame", False) and shape.text:
                    new_text = replace_text_value(shape.text, replacements)
                    if new_text != shape.text:
                        shape.text = new_text
            except Exception:
                pass


def find_main_summary_table(prs: Presentation):
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_table", False):
                continue

            table = shape.table
            if len(table.rows) < 2 or len(table.columns) < 3:
                continue

            header_text = " ".join(
                clean_text(table.cell(0, col).text).lower()
                for col in range(len(table.columns))
            )

            if "function" in header_text and "quartile" in header_text:
                return slide, shape

    raise RuntimeError(
        "Could not find the Aura Summary Bain table. Expected a table with headers "
        "'Function' and '[Client] Quartile'."
    )


def remove_bullets(paragraph) -> None:
    """Explicitly remove inherited bullet formatting from a python-pptx paragraph."""
    try:
        p_pr = paragraph._p.get_or_add_pPr()
        for child in list(p_pr):
            local_name = child.tag.rsplit("}", 1)[-1]
            if local_name.startswith("bu"):
                p_pr.remove(child)
        p_pr.insert(0, OxmlElement("a:buNone"))
    except Exception:
        pass


def set_cell_text(
    cell,
    text: str,
    *,
    font_size: float,
    bold: bool = False,
    color: RGBColor = BLACK,
    align=PP_ALIGN.LEFT,
) -> None:
    """
    Set cell text without using cell.text.

    cell.text can preserve/inherit bullet paragraph properties from the template.
    This writes a clean first paragraph/run, removes bullets, and applies a
    consistent font while preserving the cell fill/borders from the cloned row.
    """
    try:
        text_frame = cell.text_frame
        text_frame.clear()
    except Exception:
        cell.text = ""
        text_frame = cell.text_frame

    try:
        cell.margin_left = Pt(2)
        cell.margin_right = Pt(2)
        cell.margin_top = Pt(1)
        cell.margin_bottom = Pt(1)
    except Exception:
        pass

    if MSO_VERTICAL_ANCHOR is not None:
        try:
            cell.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        except Exception:
            pass

    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.level = 0
    remove_bullets(paragraph)

    run = paragraph.add_run()
    run.text = clean_text(text)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color


def rgb_to_hex(color: RGBColor) -> str:
    return f"{int(color[0]):02X}{int(color[1]):02X}{int(color[2]):02X}"


def get_cell_tc_pr(cell):
    try:
        return cell._tc.get_or_add_tcPr()
    except Exception:
        tc = cell._tc
        tc_pr = getattr(tc, "tcPr", None)
        if tc_pr is None:
            tc_pr = OxmlElement("a:tcPr")
            tc.insert(0, tc_pr)
        return tc_pr


def set_cell_border(
    cell,
    side: str,
    *,
    color: RGBColor,
    width_pt: float,
) -> None:
    """
    Set one table-cell border.
    """
    side = side.upper()
    if side not in {"T", "B", "L", "R"}:
        return

    tag = f"a:ln{side}"
    tc_pr = get_cell_tc_pr(cell)

    for existing in list(tc_pr.findall(qn(tag))):
        tc_pr.remove(existing)

    ln = OxmlElement(tag)
    ln.set("w", str(int(width_pt * 12700)))  # DrawingML line width units.
    ln.set("cap", "flat")
    ln.set("cmpd", "sng")
    ln.set("algn", "ctr")

    solid_fill = OxmlElement("a:solidFill")
    srgb_clr = OxmlElement("a:srgbClr")
    srgb_clr.set("val", rgb_to_hex(color))
    solid_fill.append(srgb_clr)
    ln.append(solid_fill)

    prst_dash = OxmlElement("a:prstDash")
    prst_dash.set("val", "solid")
    ln.append(prst_dash)

    tc_pr.append(ln)


def clear_cell_border(cell, side: str) -> None:
    """
    Remove a table-cell border.

    Used for the L0 color-band column so G&A / S&M / Ops / R&D blocks do not
    show thin gray internal row dividers. Section changes still get a thick
    black border from apply_section_separator_borders(...).
    """
    side = side.upper()
    if side not in {"T", "B", "L", "R"}:
        return

    tag = f"a:ln{side}"
    tc_pr = get_cell_tc_pr(cell)

    for existing in list(tc_pr.findall(qn(tag))):
        tc_pr.remove(existing)

    ln = OxmlElement(tag)
    ln.set("w", "0")

    no_fill = OxmlElement("a:noFill")
    ln.append(no_fill)

    tc_pr.append(ln)


def set_row_horizontal_borders(
    table,
    row_idx: int,
    *,
    top_color: RGBColor = LIGHT_GRID_GRAY,
    top_width_pt: float = 0.45,
    bottom_color: RGBColor = LIGHT_GRID_GRAY,
    bottom_width_pt: float = 0.45,
) -> None:
    """
    Apply normal thin row borders to data columns only.

    Column 0 is the L0 color-band column, so it should not receive gray
    internal dividers within the same L0 section.
    """
    for col_idx in range(len(table.columns)):
        if col_idx == 0:
            clear_cell_border(table.cell(row_idx, col_idx), "T")
            clear_cell_border(table.cell(row_idx, col_idx), "B")
            continue

        set_cell_border(
            table.cell(row_idx, col_idx),
            "T",
            color=top_color,
            width_pt=top_width_pt,
        )
        set_cell_border(
            table.cell(row_idx, col_idx),
            "B",
            color=bottom_color,
            width_pt=bottom_width_pt,
        )


def apply_section_separator_borders(table, rows: list[dict[str, Any]]) -> None:
    """
    Thin gray borders inside sections for columns 2+ only.

    The first / L0 color-band column has no gray internal dividers.
    A thick black bottom border is added across the full table only on the last
    row before the next L0 section.
    """
    for row_idx in range(1, len(table.rows)):
        set_row_horizontal_borders(
            table,
            row_idx,
            top_color=LIGHT_GRID_GRAY,
            top_width_pt=0.45,
            bottom_color=LIGHT_GRID_GRAY,
            bottom_width_pt=0.45,
        )

    for body_idx, row in enumerate(rows):
        table_row_idx = body_idx + 1
        has_next = body_idx < len(rows) - 1
        next_is_new_l0 = has_next and rows[body_idx + 1]["l0"] != row["l0"]

        if next_is_new_l0:
            for col_idx in range(len(table.columns)):
                set_cell_border(
                    table.cell(table_row_idx, col_idx),
                    "B",
                    color=BLACK,
                    width_pt=1.15,
                )


def row_first_cell_text(table, row_idx: int) -> str:
    try:
        return clean_text(table.cell(row_idx, 0).text)
    except Exception:
        return ""


def collect_template_rows(table):
    """
    Capture the template's own row XML per L0 section.

    This preserves the original section styling/colors from the template:
    G&A / S&M / Ops / R&D each use their own template section row.
    """
    original_rows = list(table._tbl.tr_lst)
    if len(original_rows) < 2:
        raise RuntimeError("Summary table must contain at least a header row and one body row.")

    header_tr = deepcopy(original_rows[0])
    section_templates: dict[str, Any] = {}
    child_templates: dict[str, Any] = {}

    for row_idx in range(1, len(original_rows)):
        l0 = normalize_l0(row_first_cell_text(table, row_idx))
        if l0 not in L0_SECTION_ORDER or l0 in section_templates:
            continue

        section_templates[l0] = deepcopy(original_rows[row_idx])

        child_template = None
        for next_idx in range(row_idx + 1, len(original_rows)):
            next_l0 = normalize_l0(row_first_cell_text(table, next_idx))
            if next_l0 in L0_SECTION_ORDER:
                break
            child_template = deepcopy(original_rows[next_idx])
            break

        child_templates[l0] = child_template if child_template is not None else deepcopy(original_rows[row_idx])

    fallback_section = deepcopy(original_rows[1])
    fallback_child = deepcopy(original_rows[2] if len(original_rows) > 2 else original_rows[1])

    return header_tr, section_templates, child_templates, fallback_section, fallback_child


def make_summary_rows(records: list[dict[str, Any]], logger: logging.Logger) -> list[dict[str, Any]]:
    """
    Build table rows from parsed Aura L1 records.

    Rows are kept only when both L0 and L1/function exist and the L0 belongs to
    one of the supported summary sections. Empty L0 sections are not padded with
    fake '-' rows.
    """
    grouped: OrderedDict[str, list[dict[str, Any]]] = OrderedDict((l0, []) for l0 in L0_SECTION_ORDER)
    skipped: list[str] = []

    for record in records:
        function_name = clean_text(record.get("function") or record.get("l1") or record.get("level1"))

        if not function_name or function_name == "-":
            skipped.append("<blank L1/function>")
            continue

        raw_l0 = record.get("l0") or record.get("level0") or infer_l0_from_l1(function_name)
        l0 = normalize_l0(raw_l0)

        if not l0 or l0 not in grouped:
            skipped.append(f"{l0 or '<blank L0>'} / {function_name}")
            continue

        grouped[l0].append(
            {
                "l0": l0,
                "function": function_name,
                "quartile": quartile_cell_text(classify_client_quartile(record)),
                "gap_median": gap_to_quartile_text(record, "median"),
                "gap_top": gap_to_quartile_text(record, "top"),
            }
        )

    if skipped:
        logger.warning("Skipped Aura Summary rows missing L0/L1 or unsupported L0: %s", skipped)

    rows: list[dict[str, Any]] = []
    for l0 in L0_SECTION_ORDER:
        entries = grouped[l0]
        if not entries:
            logger.info("No Aura Summary rows for %s; removing that section from the summary table.", l0)
            continue

        for idx, entry in enumerate(entries):
            rows.append(
                {
                    "l0": l0,
                    "function": entry["function"],
                    "quartile": entry["quartile"],
                    "gap_median": entry["gap_median"],
                    "gap_top": entry["gap_top"],
                    "is_first_in_section": idx == 0,
                }
            )

    return rows


def rebuild_summary_table(table_shape, rows: list[dict[str, Any]], client_name: str, logger: logging.Logger) -> None:
    if not rows:
        raise ValueError(
            "No valid Aura Summary rows were found. The summary table requires records with both L0 and L1/function."
        )

    table = table_shape.table
    tbl = table._tbl

    header_tr, section_templates, child_templates, fallback_section, fallback_child = collect_template_rows(table)

    original_table_height = int(table_shape.height)
    original_header_height = int(table.rows[0].height or original_table_height // 9)

    for tr in list(tbl.tr_lst):
        tbl.remove(tr)

    tbl.append(header_tr)

    for row in rows:
        l0 = row["l0"]
        if row["is_first_in_section"]:
            tbl.append(deepcopy(section_templates.get(l0, fallback_section)))
        else:
            tbl.append(deepcopy(child_templates.get(l0, fallback_child)))

    body_row_count = max(1, len(rows))
    body_height = max(1, original_table_height - original_header_height)
    body_row_height = max(50_000, int(body_height / body_row_count))

    table.rows[0].height = original_header_height
    for row_idx in range(1, len(table.rows)):
        table.rows[row_idx].height = body_row_height

    # Quick sizing knobs:
    # Increase these if the table looks too small. Decrease if rows overflow.
    header_font_size = 10.0
    body_font_size = 8.4
    if body_row_count > 22:
        body_font_size = 7.8
    if body_row_count > 30:
        body_font_size = 7.2

    # Header row.
    set_cell_text(table.cell(0, 0), "", font_size=header_font_size, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    set_cell_text(table.cell(0, 1), "Function", font_size=header_font_size, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    set_cell_text(
        table.cell(0, 2),
        f"{client_name} Quartile",
        font_size=header_font_size,
        bold=True,
        color=BLACK,
        align=PP_ALIGN.CENTER,
    )
    set_cell_text(table.cell(0, 3), "Gap to median", font_size=header_font_size, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    set_cell_text(table.cell(0, 4), "Gap to top quartile", font_size=header_font_size, bold=True, color=BLACK, align=PP_ALIGN.CENTER)

    # Body rows.
    for row_idx, row in enumerate(rows, start=1):
        is_first = bool(row["is_first_in_section"])

        set_cell_text(
            table.cell(row_idx, 0),
            row["l0"] if is_first else "",
            font_size=body_font_size,
            bold=True,
            color=WHITE,
            align=PP_ALIGN.LEFT,
        )

        set_cell_text(
            table.cell(row_idx, 1),
            row["function"],
            font_size=body_font_size,
            bold=False,
            color=BLACK,
            align=PP_ALIGN.LEFT,
        )

        quartile_text = row["quartile"]
        quartile_color, quartile_bold = quartile_text_style(quartile_text)
        set_cell_text(
            table.cell(row_idx, 2),
            quartile_text,
            font_size=body_font_size,
            bold=quartile_bold,
            color=quartile_color,
            align=PP_ALIGN.LEFT,
        )

        set_cell_text(
            table.cell(row_idx, 3),
            row["gap_median"],
            font_size=body_font_size,
            bold=False,
            color=BLACK,
            align=PP_ALIGN.LEFT,
        )

        set_cell_text(
            table.cell(row_idx, 4),
            row["gap_top"],
            font_size=body_font_size,
            bold=False,
            color=BLACK,
            align=PP_ALIGN.LEFT,
        )

    apply_section_separator_borders(table, rows)

    logger.info("Aura Summary table rebuilt with %s body rows.", len(rows))


def build_aura_summary_slide(
    excel_path: Path,
    template_ppt: Path,
    output_ppt: Path,
    logger: logging.Logger,
    sheet_name: str | None = None,
) -> Path:
    require_file(excel_path, "Excel workbook")
    require_file(template_ppt, "Aura Summary template PPT")

    output_ppt.parent.mkdir(parents=True, exist_ok=True)

    logger.info("Building Aura Summary slide from template: %s", template_ppt)
    client_name, records = read_aura_l1_records(
        excel_path=excel_path,
        template_ppt=template_ppt,
        output_ppt=output_ppt,
        logger=logger,
        sheet_name=sheet_name,
    )

    rows = make_summary_rows(records, logger)

    prs = Presentation(str(template_ppt))
    replace_non_table_placeholders(
        prs,
        {
            "[Client]": client_name,
            "[Company]": client_name,
        },
    )

    _slide, table_shape = find_main_summary_table(prs)
    rebuild_summary_table(table_shape, rows, client_name=client_name, logger=logger)

    prs.save(str(output_ppt))
    logger.info("Created Aura Summary slide deck: %s", output_ppt)
    return output_ppt


__all__ = ["build_aura_summary_slide"]
