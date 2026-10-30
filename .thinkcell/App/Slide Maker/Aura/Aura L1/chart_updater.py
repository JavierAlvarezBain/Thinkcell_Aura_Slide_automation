from __future__ import annotations

import logging
import re
import subprocess
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from shared import (
    BuildConfig,
    MSO_FALSE,
    MSO_TRUE,
    agenda_display_text,
    clean_text,
    normalize_letters,
    pump_office_ui,
    safe_text,
)


def rgb_int(red: int, green: int, blue: int) -> int:
    return int(red) + int(green) * 256 + int(blue) * 65536


QUARTILE_TEXT_COLORS = {
    "Top": rgb_int(0, 176, 80),
    "Second": rgb_int(255, 192, 0),
    "Third": rgb_int(237, 125, 49),
    "Bottom": rgb_int(192, 0, 0),
    "Unknown": rgb_int(89, 89, 89),
}

CLIENT_TAG_FALLBACK_TEXTS = {"DBS"}
L0_FALLBACK_LABELS = {"G&A", "S&M", "Ops", "R&D", "G and A", "Sales & Marketing", "Operations"}


def set_office_visibility(app, prog_id: str, visible: bool) -> None:
    try:
        if prog_id.lower() == "powerpoint.application":
            app.Visible = True
        else:
            app.Visible = visible
    except Exception:
        pass


def launch_powerpoint_externally(logger: logging.Logger) -> None:
    logger.warning("Launching PowerPoint externally...")
    try:
        subprocess.Popen(
            ["cmd", "/c", "start", "", "powerpnt"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            shell=False,
        )
    except Exception as exc:
        logger.warning("Could not launch PowerPoint via cmd/start: %s", exc)


def start_powerpoint_app(config: BuildConfig, logger: logging.Logger):
    prog_id = "PowerPoint.Application"
    logger.info("Trying to attach to existing PowerPoint instance...")
    try:
        ppt = win32.GetActiveObject(prog_id)
        set_office_visibility(ppt, prog_id, True)
        logger.info("Attached to existing PowerPoint instance.")
        return ppt
    except Exception:
        logger.info("No active PowerPoint instance found.")

    launch_powerpoint_externally(logger)
    for attempt in range(1, 61):
        try:
            pump_office_ui(0.2)
            ppt = win32.GetActiveObject(prog_id)
            set_office_visibility(ppt, prog_id, True)
            logger.info("Attached to PowerPoint after external launch.")
            return ppt
        except Exception:
            if attempt % 5 == 0:
                logger.info("Waiting for PowerPoint COM registration... %ss", attempt)
            time.sleep(1)

    raise RuntimeError(
        "PowerPoint could not be attached through COM. Open PowerPoint manually, "
        "clear all prompts, then rerun the script."
    )


EXCEL_MANUAL_DIALOG_WAIT_SECONDS = 300.0
EXCEL_MANUAL_DIALOG_LOG_INTERVAL_SECONDS = 15.0


def _finish_excel_startup_with_manual_dialog_wait(
    excel,
    *,
    requested_visible: bool,
    logger: logging.Logger,
    timeout_seconds: float = EXCEL_MANUAL_DIALOG_WAIT_SECONDS,
):
    """Wait for a blocking Excel startup/security dialog to be dismissed.

    Some managed Office installations display a modal startup dialog after the
    COM server has been created.  While that dialog is open, even a harmless
    call such as ``excel.Visible = True`` can be rejected.  Keep the same Excel
    instance alive, make the dialog visible, and retry until the user clicks
    through it instead of creating multiple Excel processes and failing.
    """
    deadline = time.monotonic() + max(5.0, float(timeout_seconds))
    last_error: BaseException | None = None
    announced = False
    next_status_log = 0.0

    while time.monotonic() < deadline:
        try:
            # Keep Excel visible while waiting so a modal startup/security
            # notice can be dismissed by the user.
            excel.Visible = True
            _ = int(excel.Workbooks.Count)
            excel.DisplayAlerts = False
            if not requested_visible:
                excel.Visible = False
            pump_office_ui(0.5)
            if announced:
                logger.info(
                    "Excel startup dialog was cleared. Automation is resuming automatically."
                )
            return excel
        except Exception as exc:
            last_error = exc
            now = time.monotonic()
            if not announced:
                logger.warning(
                    "ACTION REQUIRED: Excel has started but is blocked by a startup/security "
                    "dialog. Switch to Excel and click OK/Continue on every dialog. "
                    "The Slide Maker will wait up to 5 minutes and resume automatically; "
                    "do not click Generate again."
                )
                announced = True
                next_status_log = now
            if now >= next_status_log:
                remaining = max(0, int(round(deadline - now)))
                logger.warning(
                    "Waiting for the Excel startup dialog to be cleared... %s second(s) remaining.",
                    remaining,
                )
                next_status_log = now + EXCEL_MANUAL_DIALOG_LOG_INTERVAL_SECONDS
            try:
                pump_office_ui(0.25)
            except Exception:
                pass
            time.sleep(0.75)

    raise RuntimeError(
        "Excel remained blocked by a startup/security dialog for 5 minutes. "
        "Dismiss the Excel dialog and run the build again."
    ) from last_error


def start_excel_app(config: BuildConfig, logger: logging.Logger):
    """Start one Excel COM instance and tolerate a manual startup dialog."""
    last_error: BaseException | None = None
    excel = None

    # Retry only COM creation. Once an Excel instance exists, never create a
    # second one merely because a modal startup dialog is blocking automation.
    for attempt in range(1, 4):
        logger.info("Starting Excel.Application, attempt %s...", attempt)
        try:
            try:
                excel = win32.DispatchEx("Excel.Application")
            except Exception:
                excel = win32.Dispatch("Excel.Application")
            break
        except Exception as exc:
            last_error = exc
            logger.warning(
                "Attempt %s could not create Excel.Application: %s",
                attempt,
                exc,
            )
            if attempt < 3:
                time.sleep(2)

    if excel is None:
        raise RuntimeError("Could not create Excel.Application.") from last_error

    try:
        excel = _finish_excel_startup_with_manual_dialog_wait(
            excel,
            requested_visible=bool(config.excel_visible),
            logger=logger,
        )
    except Exception:
        # Keep the original exception, but make a best-effort cleanup so a
        # timed-out run does not deliberately leave an extra Excel instance.
        try:
            excel.Quit()
        except Exception:
            pass
        raise

    logger.info("Excel.Application started.")
    return excel


def list_com_addins(app, app_label: str, logger: logging.Logger) -> list[str]:
    candidates: list[str] = []
    try:
        count = app.COMAddIns.Count
    except Exception as exc:
        logger.warning("%s COMAddIns not accessible: %s", app_label, exc)
        return candidates

    logger.info("Inspecting %s COM add-ins...", app_label)
    for i in range(1, count + 1):
        try:
            addin = app.COMAddIns.Item(i)
        except Exception as exc:
            logger.debug("Could not access %s COM add-in %s: %s", app_label, i, exc)
            continue
        try:
            prog_id = str(addin.ProgId)
        except Exception:
            prog_id = ""
        try:
            desc = str(addin.Description)
        except Exception:
            desc = ""
        try:
            connected = bool(addin.Connect)
        except Exception:
            connected = None
        logger.info(
            "%s COM add-in %s: ProgId=%s | Connected=%s | Description=%s",
            app_label,
            i,
            prog_id,
            connected,
            desc,
        )
        if "think" in f"{prog_id} {desc}".lower() or "cell" in f"{prog_id} {desc}".lower():
            if prog_id:
                candidates.append(prog_id)
    return candidates


def ensure_excel_has_workbook(excel_app, logger: logging.Logger) -> None:
    try:
        if excel_app.Workbooks.Count == 0:
            logger.info("Excel has no workbook open. Creating a blank workbook to initialize add-ins.")
            excel_app.Workbooks.Add()
            pump_office_ui(1.0)
    except Exception as exc:
        logger.warning("Could not ensure Excel workbook exists: %s", exc)


def try_get_addin_object(app, app_label: str, prog_id: str, logger: logging.Logger):
    try:
        addin = app.COMAddIns.Item(prog_id)
    except Exception as exc:
        logger.debug("%s does not expose COM add-in %s: %s", app_label, prog_id, exc)
        return None
    try:
        connected = bool(addin.Connect)
    except Exception:
        connected = None
    logger.info("Trying think-cell add-in from %s: %s | Connected=%s", app_label, prog_id, connected)
    try:
        if hasattr(addin, "Connect") and not addin.Connect:
            logger.info("%s / %s exists but is not connected. Connecting now.", app_label, prog_id)
            addin.Connect = True
            pump_office_ui(1.5)
    except Exception as exc:
        logger.warning("Could not set Connect=True for %s / %s: %s", app_label, prog_id, exc)

    try:
        obj = addin.Object
    except Exception as exc:
        logger.warning("Could not read Object for %s / %s: %s", app_label, prog_id, exc)
        return None
    if obj is None:
        logger.warning("%s / %s found, but addin.Object is None.", app_label, prog_id)
        return None
    try:
        _ = obj.CreateUpdate
    except Exception as exc:
        logger.warning(
            "%s / %s returned an object, but CreateUpdate is not accessible: %s",
            app_label,
            prog_id,
            exc,
        )
        return None
    logger.info("Using think-cell automation object from %s / %s.", app_label, prog_id)
    return obj


def get_thinkcell_addin(excel_app, ppt_app, logger: logging.Logger):
    ensure_excel_has_workbook(excel_app, logger)
    known_prog_ids = ["thinkcell.addin", "thinkcell.addin64"]
    excel_candidates = list_com_addins(excel_app, "Excel", logger)
    ppt_candidates = list_com_addins(ppt_app, "PowerPoint", logger)

    ordered_checks = []
    for prog_id in known_prog_ids + excel_candidates:
        ordered_checks.append(("Excel", excel_app, prog_id))
    for prog_id in known_prog_ids + ppt_candidates:
        ordered_checks.append(("PowerPoint", ppt_app, prog_id))

    seen = set()
    for app_label, app, prog_id in ordered_checks:
        key = (app_label.lower(), prog_id.lower())
        if key in seen:
            continue
        seen.add(key)
        obj = try_get_addin_object(app, app_label, prog_id, logger)
        if obj is not None:
            return obj

    raise RuntimeError(
        "Could not get think-cell COM automation object. Open Excel and PowerPoint manually, "
        "confirm think-cell is loaded/enabled, then rerun."
    )


def iter_shapes(shapes):
    for i in range(1, shapes.Count + 1):
        shape = shapes.Item(i)
        yield shape
        try:
            if shape.Type == 6:
                yield from iter_shapes(shape.GroupItems)
        except Exception:
            pass


def get_slide_text(slide) -> str:
    parts = []
    for shape in iter_shapes(slide.Shapes):
        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                parts.append(shape.TextFrame.TextRange.Text)
        except Exception:
            pass
    return "\n".join(parts)


def find_chart_slide_index(pres, config: BuildConfig) -> int:
    for idx in range(1, pres.Slides.Count + 1):
        text = get_slide_text(pres.Slides(idx))
        if config.chart_slide_search_text in text:
            return idx
    return pres.Slides.Count


def keep_only_chart_slide(pres, config: BuildConfig) -> None:
    keep_idx = find_chart_slide_index(pres, config)
    for idx in range(pres.Slides.Count, 0, -1):
        if idx != keep_idx:
            pres.Slides(idx).Delete()


def dump_shapes_for_debug(pres, config: BuildConfig, logger: logging.Logger) -> None:
    if not config.dump_shapes:
        return
    slide_idx = find_chart_slide_index(pres, config)
    slide = pres.Slides(slide_idx)
    logger.debug("Shape dump for slide %s:", slide_idx)
    for shape in iter_shapes(slide.Shapes):
        try:
            text = ""
            try:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text.replace("\r", " ")
            except Exception:
                pass
            logger.debug(
                "name=%r type=%s left=%.1f top=%.1f width=%.1f height=%.1f text=%r",
                getattr(shape, "Name", ""),
                getattr(shape, "Type", ""),
                float(shape.Left),
                float(shape.Top),
                float(shape.Width),
                float(shape.Height),
                text[:120],
            )
        except Exception:
            pass


def write_chart_data_workbook(
    temp_dir: Path,
    company_headers: list,
    revenue_values: list,
    record: dict,
    config: BuildConfig,
) -> Path:
    function_name = record["function"]
    temp_excel = temp_dir / f"tc_data_{safe_text(function_name)}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "ChartData"
    ws.cell(row=config.chart_header_row, column=1).value = "Axis"
    for idx, company in enumerate(company_headers, start=2):
        ws.cell(row=config.chart_header_row, column=idx).value = company

    ws.cell(row=config.revenue_chart_row, column=1).value = "Revenue ($M)"
    for idx, value in enumerate(revenue_values, start=2):
        ws.cell(row=config.revenue_chart_row, column=idx).value = value
        if isinstance(value, (int, float)):
            ws.cell(row=config.revenue_chart_row, column=idx).number_format = "$0.00"

    for row_num in range(3, config.fte_chart_row):
        for col in range(1, config.chart_cols + 1):
            ws.cell(row=row_num, column=col).value = None

    ws.cell(row=config.fte_chart_row, column=1).value = "Bps"
    for idx, value in enumerate(record["fte_values"], start=2):
        ws.cell(row=config.fte_chart_row, column=idx).value = value
        ws.cell(row=config.fte_chart_row, column=idx).number_format = "0"

    for col in range(1, config.chart_cols + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    wb.save(temp_excel)
    return temp_excel


def shape_name_contains(shape, fragment: str) -> bool:
    try:
        return fragment.lower() in shape.Name.lower()
    except Exception:
        return False


def shape_has_text(shape) -> bool:
    try:
        return bool(shape.HasTextFrame and shape.TextFrame.HasText)
    except Exception:
        return False


def get_shape_text(shape) -> str:
    try:
        if shape_has_text(shape):
            return str(shape.TextFrame.TextRange.Text)
    except Exception:
        pass
    return ""


def find_group_item(group_shape, name_fragment: str):
    try:
        items = group_shape.GroupItems
    except Exception:
        return None
    for i in range(1, items.Count + 1):
        item = items.Item(i)
        if shape_name_contains(item, name_fragment):
            return item
        try:
            if item.Type == 6:
                nested = find_group_item(item, name_fragment)
                if nested is not None:
                    return nested
        except Exception:
            pass
    return None


def agenda_text_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def find_group_text_item_exact(group_shape, accepted_texts: set[str]):
    try:
        items = group_shape.GroupItems
    except Exception:
        return None
    accepted_keys = {agenda_text_key(text) for text in accepted_texts}
    for i in range(1, items.Count + 1):
        item = items.Item(i)
        try:
            if shape_has_text(item) and agenda_text_key(get_shape_text(item)) in accepted_keys:
                return item
        except Exception:
            pass
        try:
            if item.Type == 6:
                nested = find_group_text_item_exact(item, accepted_texts)
                if nested is not None:
                    return nested
        except Exception:
            pass
    return None


def find_running_agenda_group(slide):
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes.Item(i)
        try:
            if shape.Type == 6 and "runningagenda" in shape.Name.lower():
                return shape
        except Exception:
            pass
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes.Item(i)
        try:
            if shape.Type == 6 and (find_group_item(shape, "TextRight") or find_group_item(shape, "TextLeft")):
                return shape
        except Exception:
            pass
    return None


def set_shape_text_preserve_format(shape, new_text: str) -> None:
    try:
        shape.TextFrame.TextRange.Text = new_text
    except Exception:
        pass
    try:
        shape.TextFrame2.TextRange.Text = new_text
    except Exception:
        pass


def agenda_l0_display_text(l0_name: str) -> str:
    raw = clean_text(l0_name)
    compact = agenda_text_key(raw)
    mapping = {
        "ga": "G&A",
        "ganda": "G&A",
        "generalandadministrative": "G&A",
        "generaladministrative": "G&A",
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
    return mapping.get(compact, raw)


def update_bain_agenda_text_only(slide, function_name: str, l0_name: str, logger: logging.Logger) -> None:
    display_right = agenda_display_text(function_name)
    display_left = agenda_l0_display_text(l0_name)
    agenda_group = find_running_agenda_group(slide)
    updated_left = False
    updated_right = False

    if agenda_group is not None:
        text_left = find_group_item(agenda_group, "TextLeft")
        if text_left is None:
            text_left = find_group_text_item_exact(agenda_group, L0_FALLBACK_LABELS)
        if text_left is not None and display_left:
            set_shape_text_preserve_format(text_left, display_left)
            updated_left = True

        text_right = find_group_item(agenda_group, "TextRight")
        if text_right is not None:
            set_shape_text_preserve_format(text_right, display_right)
            updated_right = True

    l0_keys = {agenda_text_key(value) for value in L0_FALLBACK_LABELS}
    for shape in iter_shapes(slide.Shapes):
        try:
            if not shape_has_text(shape):
                continue
            current_text = get_shape_text(shape).strip()
            current_key = agenda_text_key(current_text)

            if not updated_left and display_left and current_key in l0_keys:
                set_shape_text_preserve_format(shape, display_left)
                updated_left = True
                continue

            if not updated_right and normalize_letters(current_text) == "finance":
                set_shape_text_preserve_format(shape, display_right)
                updated_right = True
                continue
        except Exception:
            pass

    if not updated_left:
        logger.debug("Could not find Bain agenda L0/dark-gray text for %s / %s", l0_name, function_name)
    if not updated_right:
        logger.debug("Could not find Bain agenda L1/light-gray text for %s", function_name)


def replace_template_tokens(text: str, client_company: str, quartile_label: str | None = None) -> str:
    new_text = text
    if client_company:
        new_text = re.sub(r"\[\s*client\s*\]", client_company, new_text, flags=re.I)
        new_text = re.sub(r"\[\s*company\s*\]", client_company, new_text, flags=re.I)
    if quartile_label:
        new_text = re.sub(r"\[\s*quartile\s*\]", quartile_label, new_text, flags=re.I)
    return new_text


def collapse_client_tag_if_needed(text: str, client_company: str) -> str | None:
    stripped = re.sub(r"\s+", " ", clean_text(text))
    lower = stripped.lower()

    if not client_company:
        return None

    if re.fullmatch(r"\[\s*client\s*\]\s*/\s*\[\s*company\s*\]", lower, flags=re.I):
        return client_company
    if re.fullmatch(r"\[\s*company\s*\]\s*/\s*\[\s*client\s*\]", lower, flags=re.I):
        return client_company
    if re.fullmatch(r"\[\s*client\s*\]", lower, flags=re.I):
        return client_company
    if re.fullmatch(r"\[\s*company\s*\]", lower, flags=re.I):
        return client_company

    if stripped in CLIENT_TAG_FALLBACK_TEXTS:
        return client_company

    return None


def is_title_text(text: str) -> bool:
    normalized = re.sub(r"\s+", " ", clean_text(text))
    lower = normalized.lower()
    return "|" in normalized and "fte" in lower and "revenue" in lower


def find_case_insensitive(haystack: str, needle: str) -> int:
    if not needle:
        return -1
    return haystack.lower().find(needle.lower())


def apply_substring_font(shape, substring: str, rgb_color: int | None = None, bold: bool | None = None) -> bool:
    try:
        tr = shape.TextFrame.TextRange
        text = str(tr.Text)
        idx = find_case_insensitive(text, substring)
        if idx < 0:
            return False
        chars = tr.Characters(idx + 1, len(substring))
        if bold is not None:
            chars.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        if rgb_color is not None:
            chars.Font.Color.RGB = rgb_color
        return True
    except Exception:
        return False


def apply_textframe2_substring_font(shape, substring: str, rgb_color: int | None = None, bold: bool | None = None) -> bool:
    try:
        tr = shape.TextFrame2.TextRange
        text = str(tr.Text)
        idx = find_case_insensitive(text, substring)
        if idx < 0:
            return False
        chars = tr.Characters(idx + 1, len(substring))
        if bold is not None:
            chars.Font.Bold = MSO_TRUE if bold else MSO_FALSE
        if rgb_color is not None:
            chars.Font.Fill.Visible = MSO_TRUE
            chars.Font.Fill.ForeColor.RGB = rgb_color
        return True
    except Exception:
        return False


def apply_quartile_phrase_format(shape, quartile_label: str) -> None:
    color = QUARTILE_TEXT_COLORS.get(quartile_label, QUARTILE_TEXT_COLORS["Unknown"])
    label = clean_text(quartile_label) or "Unknown"
    phrase = f"{label} quartile"

    applied_1 = apply_substring_font(shape, phrase, rgb_color=color, bold=True)
    applied_2 = apply_textframe2_substring_font(shape, phrase, rgb_color=color, bold=True)

    if not applied_1:
        apply_substring_font(shape, label, rgb_color=color, bold=True)
    if not applied_2:
        apply_textframe2_substring_font(shape, label, rgb_color=color, bold=True)


def replace_title_client_function_phrase(text: str, function_name: str, client_company: str) -> str:
    if "|" not in text or not re.search(r"\bFTEs\b", text, flags=re.I):
        return text

    left, _, right = text.partition("|")
    new_right = re.sub(
        r"^\s*.*?\bFTEs\b",
        f" {client_company} {function_name} FTEs",
        right,
        count=1,
        flags=re.I,
    )
    return f"{left.strip()} |{new_right}"


def update_title_shape(
    shape,
    function_name: str,
    client_company: str,
    quartile_label: str,
    logger: logging.Logger,
) -> None:
    try:
        tr = shape.TextFrame.TextRange
    except Exception:
        return

    old_text = str(tr.Text)
    new_text = old_text

    new_text = re.sub(r"^\s*[^|]+\|", f"{function_name} |", new_text, count=1)
    new_text = replace_template_tokens(new_text, client_company, quartile_label)
    new_text = replace_title_client_function_phrase(new_text, function_name, client_company)

    new_text = re.sub(
        r"\b(top|second|third|bottom)\s+quartile\b",
        f"{quartile_label} quartile",
        new_text,
        count=1,
        flags=re.I,
    )

    try:
        tr.Text = new_text
    except Exception as exc:
        logger.debug("Could not set title text: %s", exc)
        return

    try:
        tr.Font.Bold = MSO_FALSE
    except Exception:
        pass

    apply_substring_font(shape, function_name, bold=True)
    apply_textframe2_substring_font(shape, function_name, bold=True)
    apply_quartile_phrase_format(shape, quartile_label)


def update_generic_placeholder_shape(
    shape,
    client_company: str,
    quartile_label: str,
    logger: logging.Logger,
) -> bool:
    old_text = get_shape_text(shape)
    if not old_text:
        return False

    collapsed = collapse_client_tag_if_needed(old_text, client_company)
    if collapsed is not None:
        set_shape_text_preserve_format(shape, collapsed)
        return True

    new_text = replace_template_tokens(old_text, client_company, quartile_label)
    if new_text == old_text:
        return False

    try:
        shape.TextFrame.TextRange.Text = new_text
    except Exception as exc:
        logger.debug("Could not replace placeholder text: %s", exc)
        return False

    apply_quartile_phrase_format(shape, quartile_label)
    return True


def iter_related_shape_collections(slide):
    try:
        yield slide.Shapes
    except Exception:
        pass
    try:
        yield slide.CustomLayout.Shapes
    except Exception:
        pass
    try:
        yield slide.Master.Shapes
    except Exception:
        pass


def update_slide_text(
    pres,
    function_name: str,
    l0_name: str,
    client_company: str,
    quartile_label: str,
    config: BuildConfig,
    logger: logging.Logger,
) -> None:
    client_company = clean_text(client_company) or "Client"
    quartile_label = clean_text(quartile_label) or "Unknown"
    l0_name = clean_text(l0_name)

    for slide_idx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(slide_idx)

        for shapes in iter_related_shape_collections(slide):
            for shape in iter_shapes(shapes):
                try:
                    if not shape_has_text(shape):
                        continue
                    old_text = get_shape_text(shape)
                    if is_title_text(old_text):
                        update_title_shape(shape, function_name, client_company, quartile_label, logger)
                    else:
                        update_generic_placeholder_shape(shape, client_company, quartile_label, logger)
                except Exception as exc:
                    logger.debug("Text update skipped for one shape: %s", exc)

        if config.update_agenda_text:
            update_bain_agenda_text_only(slide, function_name, l0_name, logger)


def update_thinkcell_chart(
    tc_addin,
    pres,
    excel_app,
    chart_data_path: Path,
    config: BuildConfig,
    logger: logging.Logger,
) -> None:
    if tc_addin is None:
        raise RuntimeError("think-cell COM automation object is None before chart update.")
    last_error = None
    for attempt in range(1, config.thinkcell_retries + 1):
        wb = None
        try:
            logger.info("think-cell update attempt %s...", attempt)
            wb = excel_app.Workbooks.Open(str(chart_data_path), UpdateLinks=0, ReadOnly=True)
            ws = wb.Worksheets("ChartData")
            rng = ws.Range(ws.Cells(1, 1), ws.Cells(config.chart_rows, config.chart_cols))
            pump_office_ui(0.4)
            update = tc_addin.CreateUpdate()
            update.AddRangeData(pres, config.chart_name, rng, False)
            pump_office_ui(0.4)
            update.Send()
            pump_office_ui(max(0.8, config.post_chart_update_wait_seconds))
            logger.info("think-cell update succeeded.")
            return
        except Exception as exc:
            last_error = exc
            logger.warning("think-cell update failed on attempt %s: %s", attempt, exc)
            pump_office_ui(1.5)
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
    raise RuntimeError("think-cell update failed after retries.") from last_error


def run_com_check(config: BuildConfig, logger: logging.Logger) -> None:
    logger.info("Running COM health check...")
    excel = None
    ppt = None
    try:
        excel = start_excel_app(config, logger)
        ppt = start_powerpoint_app(config, logger)
        try:
            logger.info("Excel version: %s", excel.Version)
        except Exception:
            logger.info("Excel started, but version could not be read.")
        try:
            logger.info("PowerPoint version: %s", ppt.Version)
        except Exception:
            logger.info("PowerPoint started, but version could not be read.")
        get_thinkcell_addin(excel, ppt, logger)
        logger.info("COM health check passed.")
    finally:
        if config.quit_office_apps:
            try:
                if ppt is not None:
                    ppt.Quit()
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass