from __future__ import annotations

import logging
import re
import time
from pathlib import Path

from geocost_common import (
    AUTOLABEL_PREFIX,
    DEFAULT_THINKCELL_ATTEMPTS,
    GeoCostData,
    GeoCostTemplateLayout,
    LabelTarget,
    MSO_ANCHOR_MIDDLE,
    MSO_FALSE,
    MSO_LINE,
    PLAIN_PERCENT_NUMBER_FORMAT,
    PLAIN_VALUE_NUMBER_FORMAT,
    PP_ALIGN_CENTER,
    XL_DATA_LABELS_SHOW_VALUE,
    XL_LABEL_POSITION_CENTER,
    canonical_country_key,
    clean_text,
    compact_text,
    country_key_set as _country_key_set,
    font_hex_for_bucket as _font_hex_for_bucket,
    is_disconnected_office_error as _is_disconnected_office_error,
    is_transient_office_error as _is_transient_office_error,
    pump_office_ui,
    require_windows_com,
    win32,
)

EXCEL_MANUAL_DIALOG_WAIT_SECONDS = 300.0
EXCEL_MANUAL_DIALOG_LOG_INTERVAL_SECONDS = 15.0


def _finish_excel_startup_with_manual_dialog_wait(
    excel,
    logger: logging.Logger,
    timeout_seconds: float = EXCEL_MANUAL_DIALOG_WAIT_SECONDS,
):
    """Wait for the user to dismiss a blocking Excel startup/security dialog."""
    deadline = time.monotonic() + max(5.0, float(timeout_seconds))
    last_error: BaseException | None = None
    announced = False
    next_status_log = 0.0

    while time.monotonic() < deadline:
        try:
            excel.Visible = True
            _ = int(excel.Workbooks.Count)
            excel.DisplayAlerts = False
            pump_office_ui(0.5)
            if announced:
                logger.info(
                    "Excel startup dialog was cleared. Geo Cost automation is resuming automatically."
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


def start_excel_app(logger: logging.Logger):
    """Start one Excel instance and pause for any managed startup dialog."""
    require_windows_com()
    logger.info("Starting Excel.Application...")
    try:
        excel = win32.DispatchEx("Excel.Application")
    except Exception:
        excel = win32.Dispatch("Excel.Application")

    try:
        return _finish_excel_startup_with_manual_dialog_wait(excel, logger)
    except Exception:
        try:
            excel.Quit()
        except Exception:
            pass
        raise


def start_powerpoint_app(logger: logging.Logger):
    require_windows_com()
    logger.info("Starting/attaching PowerPoint.Application...")
    try:
        ppt = win32.GetActiveObject("PowerPoint.Application")
    except Exception:
        try:
            ppt = win32.Dispatch("PowerPoint.Application")
        except Exception:
            ppt = win32.DispatchEx("PowerPoint.Application")
    try:
        ppt.Visible = True
    except Exception:
        pass
    pump_office_ui(1.0)
    return ppt


def ensure_excel_has_workbook(excel_app) -> None:
    try:
        if excel_app.Workbooks.Count == 0:
            excel_app.Workbooks.Add()
            pump_office_ui(0.8)
    except Exception:
        pass


def list_com_addins(app) -> list[str]:
    result: list[str] = []
    try:
        count = app.COMAddIns.Count
    except Exception:
        return result
    for index in range(1, count + 1):
        try:
            addin = app.COMAddIns.Item(index)
            text = f"{addin.ProgId} {addin.Description}"
            if "think" in text.lower() or "cell" in text.lower():
                result.append(str(addin.ProgId))
        except Exception:
            pass
    return result


def try_get_thinkcell_object(app, prog_id: str):
    try:
        addin = app.COMAddIns.Item(prog_id)
    except Exception:
        return None
    try:
        if hasattr(addin, "Connect") and not addin.Connect:
            addin.Connect = True
            pump_office_ui(1.0)
    except Exception:
        pass
    try:
        obj = addin.Object
    except Exception:
        return None
    if obj is None:
        return None
    try:
        _ = obj.CreateUpdate
    except Exception:
        return None
    return obj


def get_thinkcell_addin(excel_app, ppt_app, logger: logging.Logger):
    ensure_excel_has_workbook(excel_app)
    candidates = ["thinkcell.addin", "thinkcell.addin64"]
    candidates += list_com_addins(excel_app)
    candidates += list_com_addins(ppt_app)
    seen: set[str] = set()
    for prog_id in candidates:
        key = prog_id.lower()
        if key in seen:
            continue
        seen.add(key)
        for app in (excel_app, ppt_app):
            obj = try_get_thinkcell_object(app, prog_id)
            if obj is not None:
                logger.info("Using think-cell add-in: %s", prog_id)
                return obj
    raise RuntimeError("Could not access the think-cell COM add-in.")


def reacquire_thinkcell_addin(
    excel_app,
    ppt_app,
    logger: logging.Logger,
    max_attempts: int = 3,
):
    """Acquire a fresh think-cell COM proxy instead of reusing a dead proxy."""
    last_error: BaseException | None = None
    attempts = max(1, int(max_attempts))
    for attempt in range(1, attempts + 1):
        try:
            obj = get_thinkcell_addin(excel_app, ppt_app, logger)
            logger.debug("Acquired fresh think-cell COM proxy on attempt %s.", attempt)
            return obj
        except Exception as exc:
            last_error = exc
            if attempt >= attempts:
                break
            logger.warning(
                "Could not reacquire think-cell on attempt %s/%s; retrying after Office settles.",
                attempt,
                attempts,
            )
            logger.debug("think-cell reacquisition error: %r", exc)
            pump_office_ui(1.2 * attempt)
    raise RuntimeError("Could not acquire a healthy think-cell COM proxy.") from last_error


def _get_slide_with_retry(
    pres,
    slide_index: int,
    logger: logging.Logger,
    context: str,
    max_attempts: int = 8,
):
    """Return a fully hydrated PowerPoint Slide COM proxy.

    ``pres.Slides(index)`` relies on the collection's default COM member. After
    repeated think-cell template opens, pywin32 can occasionally return a
    late-bound ``<unknown>`` proxy from that shorthand. The proxy exists, but
    accessing ``.Shapes`` then raises ``AttributeError: <unknown>.Shapes``.

    Use the explicit ``Slides.Item`` member, force both the slide and its Shapes
    collection to hydrate, and reacquire them during short PowerPoint busy
    windows. This operation is read-only and therefore safe to retry.
    """
    index = int(slide_index)
    attempts = max(1, int(max_attempts))
    last_error: BaseException | None = None

    for attempt in range(1, attempts + 1):
        try:
            pump_office_ui(0.25)
            slides = pres.Slides
            slide_count = int(slides.Count)
            if index < 1 or index > slide_count:
                raise IndexError(
                    f"Slide index {index} is outside the presentation's "
                    f"1..{slide_count} range."
                )
            slide = slides.Item(index)

            # Force pywin32 to resolve a real PowerPoint Slide dispatch object,
            # not an unready/late-bound <unknown> proxy.
            _ = int(slide.SlideIndex)
            shapes = slide.Shapes
            _ = int(shapes.Count)
            return slide
        except Exception as exc:
            last_error = exc
            if attempt >= attempts:
                break
            logger.warning(
                "PowerPoint slide %s was not COM-ready during %s "
                "(attempt %s/%s); reacquiring the slide proxy.",
                index,
                context,
                attempt,
                attempts,
            )
            logger.debug(
                "Slide COM-readiness error during %s: %r",
                context,
                exc,
            )
            pump_office_ui(min(0.45 * attempt, 2.0))

    raise RuntimeError(
        f"PowerPoint did not expose a usable Slide/Shapes COM object for slide "
        f"{index} during {context} after {attempts} attempts."
    ) from last_error


def activate_first_slide(pres) -> None:
    try:
        pres.Windows(1).Activate()
    except Exception:
        pass
    try:
        pres.Application.ActiveWindow.View.GotoSlide(1)
    except Exception:
        pass
    try:
        pres.Slides.Item(1).Select()
    except Exception:
        pass


def update_thinkcell_chart(
    tc_addin,
    pres,
    excel_app,
    ppt_app,
    chart_data_path: Path,
    template_layout: GeoCostTemplateLayout,
    cols_count: int,
    chart_name: str,
    logger: logging.Logger,
    post_wait: float,
    max_attempts: int = DEFAULT_THINKCELL_ATTEMPTS,
):
    """Update one think-cell chart and return the healthy COM proxy used.

    A workbook and range are reopened for every retry. When think-cell reports
    that its object has disconnected, the stale proxy is discarded and a fresh
    add-in object is acquired before the next attempt.
    """
    chart_name = clean_text(chart_name) or template_layout.chart_name
    if not chart_name:
        raise ValueError("think-cell chart name must not be empty.")

    last_error: BaseException | None = None
    attempts = max(1, int(max_attempts))

    for attempt in range(1, attempts + 1):
        workbook = None
        try:
            if tc_addin is None or attempt > 1:
                tc_addin = reacquire_thinkcell_addin(
                    excel_app,
                    ppt_app,
                    logger,
                    max_attempts=3,
                )

            workbook = excel_app.Workbooks.Open(
                str(chart_data_path),
                UpdateLinks=0,
                ReadOnly=False,
            )
            worksheet = workbook.Worksheets("ChartData")

            country_values = worksheet.Range(
                worksheet.Cells(template_layout.country_first_row, 2),
                worksheet.Cells(template_layout.country_last_row, cols_count),
            )
            country_values.Value2 = country_values.Value2
            country_values.NumberFormat = PLAIN_VALUE_NUMBER_FORMAT

            if template_layout.country_last_row + 1 <= template_layout.data_row - 1:
                spacer_values = worksheet.Range(
                    worksheet.Cells(template_layout.country_last_row + 1, 2),
                    worksheet.Cells(template_layout.data_row - 1, cols_count),
                )
                spacer_values.ClearContents()
                spacer_values.NumberFormat = "General"

            low_cost_values = worksheet.Range(
                worksheet.Cells(template_layout.data_row, 2),
                worksheet.Cells(template_layout.data_row, cols_count),
            )
            low_cost_values.Value2 = low_cost_values.Value2
            low_cost_values.NumberFormat = PLAIN_PERCENT_NUMBER_FORMAT
            workbook.Save()

            data_range = worksheet.Range(
                worksheet.Cells(1, 1),
                worksheet.Cells(template_layout.data_row, cols_count),
            )
            logger.info(
                "Prepared AddRangeData source as plain values: country cells=%s, "
                "number format=%r; no quoted labels or hidden ;;; formats.",
                int(country_values.Count),
                PLAIN_VALUE_NUMBER_FORMAT,
            )

            activate_first_slide(pres)
            pump_office_ui(max(0.8, 0.8 * attempt))
            logger.info(
                "Updating think-cell element %r (attempt %s/%s)...",
                chart_name,
                attempt,
                attempts,
            )
            update = tc_addin.CreateUpdate()
            update.AddRangeData(pres, chart_name, data_range, False)
            pump_office_ui(0.5)
            update.Send()
            pump_office_ui(max(1.0, post_wait))
            logger.info("think-cell data update succeeded for %r.", chart_name)
            return tc_addin

        except Exception as exc:
            last_error = exc
            disconnected = _is_disconnected_office_error(exc)
            transient = _is_transient_office_error(exc)
            if disconnected:
                tc_addin = None
                logger.warning(
                    "think-cell COM proxy disconnected on attempt %s/%s; "
                    "a fresh proxy will be acquired.",
                    attempt,
                    attempts,
                )
            elif attempt < attempts:
                logger.warning(
                    "think-cell update failed on attempt %s/%s; retrying after Office settles.",
                    attempt,
                    attempts,
                )
            logger.debug("think-cell update error: %r", exc)

            if attempt >= attempts or (not transient and not disconnected):
                break
            pump_office_ui(1.5 * attempt)

        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=True)
            except Exception:
                pass

    raise RuntimeError(
        f"think-cell could not update verified template element {chart_name!r} "
        f"after {attempts} attempts."
    ) from last_error


# =============================================================================
# Slide text and shape helpers


def iter_shapes(shapes):
    for index in range(1, shapes.Count + 1):
        shape = shapes.Item(index)
        yield shape
        try:
            if int(shape.Type) == 6:
                yield from iter_shapes(shape.GroupItems)
        except Exception:
            pass


def shape_has_text(shape) -> bool:
    try:
        return bool(shape.HasTextFrame and shape.TextFrame.HasText)
    except Exception:
        return False


def get_shape_text(shape) -> str:
    try:
        return str(shape.TextFrame.TextRange.Text) if shape_has_text(shape) else ""
    except Exception:
        return ""


def set_shape_text(shape, text: str) -> None:
    try:
        shape.TextFrame.TextRange.Text = text
    except Exception:
        pass
    try:
        shape.TextFrame2.TextRange.Text = text
    except Exception:
        pass


def update_slide_text(pres, data: GeoCostData, logger: logging.Logger) -> None:
    client = data.companies[0]
    client_ratio = data.low_cost_ratio[0]
    client_ratio_text = "-" if client_ratio is None else f"{client_ratio * 100:.0f}%"
    function_keys = {
        "overall", "ga", "finance", "hr", "it", "legal", "facilitiesmanagement",
        "sm", "sales", "marketing", "customerservice", "ops", "operations",
        "rd", "researchdevelopment",
    }

    # Let a newly opened think-cell template finish hydrating before traversing
    # PowerPoint shapes. The category labels themselves are not scanned here;
    # they remain part of the chart and are left untouched.
    activate_first_slide(pres)
    pump_office_ui(0.35)
    slide_count = int(pres.Slides.Count)

    for slide_index in range(1, slide_count + 1):
        slide = _get_slide_with_retry(
            pres,
            slide_index,
            logger,
            context=f"Geo Cost text update for {data.function_name}",
        )
        shapes = slide.Shapes
        for shape in iter_shapes(shapes):
            if not shape_has_text(shape):
                continue
            old_text = get_shape_text(shape)
            stripped = clean_text(old_text)
            new_text = old_text
            if "low-cost locations" in stripped.lower() and "has" in stripped.lower():
                new_text = re.sub(
                    r"^.*?\bhas\s+~?\d+%",
                    f"{client} has ~{client_ratio_text}",
                    new_text,
                    count=1,
                    flags=re.I | re.S,
                )
            if compact_text(stripped) in function_keys:
                new_text = data.function_name.upper() if stripped.isupper() else data.function_name
            if new_text != old_text:
                set_shape_text(shape, new_text)
    logger.info("Updated slide title/function text.")



# =============================================================================
# Live per-segment point labels



def _shape_bounds_com(shape) -> tuple[float, float, float, float] | None:
    try:
        left = float(shape.Left)
        top = float(shape.Top)
        width = abs(float(shape.Width))
        height = abs(float(shape.Height))
    except Exception:
        return None
    return left, top, width, height


def _shape_center_in_bounds_com(
    shape,
    target_bounds: tuple[float, float, float, float],
    margin: float,
) -> bool:
    bounds = _shape_bounds_com(shape)
    if bounds is None:
        return False
    left, top, width, height = bounds
    target_left, target_top, target_width, target_height = target_bounds
    center_x = left + width / 2.0
    center_y = top + height / 2.0
    return (
        target_left - margin <= center_x <= target_left + target_width + margin
        and target_top - margin <= center_y <= target_top + target_height + margin
    )


def _shape_text_com(shape) -> str:
    try:
        if bool(shape.HasTextFrame) and bool(shape.TextFrame.HasText):
            return str(shape.TextFrame.TextRange.Text).strip()
    except Exception:
        pass
    return ""


def _find_largest_chart_shape_com(slide):
    candidates = []
    for index in range(1, int(slide.Shapes.Count) + 1):
        shape = slide.Shapes.Item(index)
        try:
            if not bool(shape.HasChart):
                continue
        except Exception:
            continue
        bounds = _shape_bounds_com(shape)
        area = 0.0 if bounds is None else bounds[2] * bounds[3]
        candidates.append((area, shape))
    if not candidates:
        raise RuntimeError("No PowerPoint chart object was found on the GeoCost slide.")
    return max(candidates, key=lambda item: item[0])[1]


def _remove_stale_country_shapes_com(
    slide,
    chart_shape,
    template_country_names: list[str],
) -> int:
    keys = _country_key_set(template_country_names)
    chart_bounds = _shape_bounds_com(chart_shape)
    if chart_bounds is None:
        return 0

    margin = 25.2
    removed = 0
    for index in range(int(slide.Shapes.Count), 0, -1):
        shape = slide.Shapes.Item(index)
        try:
            if shape.Id == chart_shape.Id:
                continue
        except Exception:
            pass

        try:
            shape_name = str(shape.Name)
        except Exception:
            shape_name = ""
        text = _shape_text_com(shape)
        is_generated = shape_name.startswith(AUTOLABEL_PREFIX)
        is_country = canonical_country_key(text) in keys

        try:
            shape_type = int(shape.Type)
        except Exception:
            shape_type = -1

        should_remove = False
        if (is_generated or is_country) and _shape_center_in_bounds_com(
            shape, chart_bounds, margin
        ):
            should_remove = True
        elif shape_type == MSO_LINE and _shape_center_in_bounds_com(
            shape, chart_bounds, margin
        ):
            should_remove = True

        if should_remove:
            try:
                shape.Delete()
                removed += 1
            except Exception:
                pass
    return removed


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


def _delete_series_labels_com(series) -> None:
    try:
        series.HasDataLabels = False
        return
    except Exception:
        pass
    try:
        series.DataLabels().Delete()
    except Exception:
        pass


def _set_point_label_text_com(label, text: str) -> None:
    last_error = None
    for setter in (
        lambda: setattr(label, "Text", text),
        lambda: setattr(label.Format.TextFrame2.TextRange, "Text", text),
        lambda: setattr(label.Format.TextFrame.TextRange, "Text", text),
    ):
        try:
            setter()
            return
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"Could not set chart-point label text to {text!r}.") from last_error


def _format_point_label_com(label, target: LabelTarget) -> None:
    try:
        label.Position = XL_LABEL_POSITION_CENTER
    except Exception as exc:
        raise RuntimeError(
            f"Could not center label for {target.company_index=}, {target.country!r}."
        ) from exc

    font_rgb = int(_font_hex_for_bucket(target.segment_bucket), 16)
    try:
        frame = label.Format.TextFrame2
        frame.VerticalAnchor = MSO_ANCHOR_MIDDLE
        text_range = frame.TextRange
        text_range.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        text_range.Font.Size = 10
        try:
            text_range.Font.Name = "Arial"
        except Exception:
            pass
        text_range.Font.Fill.ForeColor.RGB = font_rgb
        try:
            text_range.Font.Bold = MSO_FALSE
        except Exception:
            pass
    except Exception:
        try:
            label.Font.Size = 10
        except Exception:
            pass
        try:
            label.Font.Name = "Arial"
        except Exception:
            pass
        try:
            label.Font.Bold = MSO_FALSE
        except Exception:
            pass
        try:
            label.Font.Color = font_rgb
        except Exception:
            pass

    try:
        label.Format.Fill.Visible = MSO_FALSE
    except Exception:
        pass
    try:
        label.Format.Line.Visible = MSO_FALSE
    except Exception:
        pass


def _apply_one_point_label_com(series_collection, target: LabelTarget) -> None:
    series = series_collection.Item(target.row_index + 1)
    points = _points_collection_com(series)
    point = points.Item(target.company_index + 1)
    applied = False
    for args in ((XL_DATA_LABELS_SHOW_VALUE,), tuple()):
        try:
            point.ApplyDataLabels(*args)
            applied = True
            break
        except Exception:
            continue
    if not applied:
        raise RuntimeError(
            f"PowerPoint could not add a data label for company index "
            f"{target.company_index}, country row {target.row_index}."
        )
    label = point.DataLabel
    try:
        label.AutoText = False
    except Exception:
        pass
    _set_point_label_text_com(label, target.display_label)
    _format_point_label_com(label, target)
    try:
        label.ShowLegendKey = False
    except Exception:
        pass
    try:
        label.ShowCategoryName = False
    except Exception:
        pass
    try:
        label.ShowSeriesName = False
    except Exception:
        pass
    try:
        label.ShowPercentage = False
    except Exception:
        pass
    try:
        label.ShowBubbleSize = False
    except Exception:
        pass


def _read_point_label_text_com(point) -> str:
    try:
        return clean_text(point.DataLabel.Text)
    except Exception:
        pass
    try:
        return clean_text(point.DataLabel.Format.TextFrame2.TextRange.Text)
    except Exception:
        return ""


def apply_centered_labels_to_slide_com(
    slide,
    data: GeoCostData,
    logger: logging.Logger,
) -> tuple[int, int]:
    chart_shape = _find_largest_chart_shape_com(slide)
    chart = chart_shape.Chart
    series_collection = _series_collection_com(chart)
    series_count = int(series_collection.Count)
    expected_series = data.template_layout.country_count
    if series_count < expected_series:
        raise RuntimeError(
            f"The live PowerPoint chart exposes {series_count} series; "
            f"expected at least {expected_series}."
        )
    removed = 0
    for series_index in range(1, expected_series + 1):
        try:
            _delete_series_labels_com(series_collection.Item(series_index))
        except Exception:
            pass
    for target in data.label_targets:
        _apply_one_point_label_com(series_collection, target)
    try:
        chart.Refresh()
    except Exception:
        pass
    pump_office_ui(0.4)
    actual: dict[tuple[int, int], str] = {}
    for target in data.label_targets:
        series = series_collection.Item(target.row_index + 1)
        point = _points_collection_com(series).Item(target.company_index + 1)
        try:
            has_label = bool(point.HasDataLabel)
        except Exception:
            has_label = True
        if not has_label:
            raise RuntimeError(
                f"Missing chart label for {target.company_index=}, {target.row_index=}."
            )
        text = _read_point_label_text_com(point)
        actual[(target.company_index, target.row_index)] = text
        try:
            position = int(point.DataLabel.Position)
        except Exception:
            position = XL_LABEL_POSITION_CENTER
        if position != XL_LABEL_POSITION_CENTER:
            raise RuntimeError(
                f"Chart label for {target.country!r} is not centered; position={position}."
            )
    expected = {
        (target.company_index, target.row_index): target.display_label
        for target in data.label_targets
    }
    if actual != expected:
        raise RuntimeError(
            "Live PowerPoint chart-label verification failed. "
            f"Expected {sorted(expected.items())}, got {sorted(actual.items())}."
        )
    logger.info(
        "Applied and verified %s centered per-bar label(s); removed %s stale "
        "country-label/leader-line shape(s) on the live PowerPoint chart.",
        len(data.label_targets),
        removed,
    )
    return len(data.label_targets), removed


def apply_centered_labels_to_presentation_com(
    pres,
    data_by_slide: list[GeoCostData],
    logger: logging.Logger,
) -> None:
    if int(pres.Slides.Count) < len(data_by_slide):
        raise RuntimeError(
            f"Combined deck has {pres.Slides.Count} slide(s), but "
            f"{len(data_by_slide)} GeoCost slide(s) were built."
        )
    total_labels = 0
    total_removed = 0
    for slide_index, data in enumerate(data_by_slide, start=1):
        slide = _get_slide_with_retry(
            pres,
            slide_index,
            logger,
            context=f"final Geo Cost label pass for {data.function_name}",
        )
        labels, removed = apply_centered_labels_to_slide_com(
            slide, data, logger
        )
        total_labels += labels
        total_removed += removed
    logger.info(
        "Final live-deck label pass complete: %s centered label(s), %s stale "
        "shape(s) removed; no PPTX/OLE binary patching used.",
        total_labels,
        total_removed,
    )


def insert_slide_with_retry(
    final_pres,
    temp_ppt: Path,
    logger: logging.Logger,
    max_attempts: int = 3,
) -> None:
    last_error = None
    for attempt in range(1, max_attempts + 1):
        try:
            pump_office_ui(0.5 * attempt)
            final_pres.Slides.InsertFromFile(
                str(temp_ppt), int(final_pres.Slides.Count), 1, 1
            )
            pump_office_ui(0.4)
            logger.info("Appended slide from: %s", temp_ppt.name)
            return
        except Exception as exc:
            last_error = exc
            logger.warning(
                "Slide insertion attempt %s/%s failed; retrying after PowerPoint settles.",
                attempt,
                max_attempts,
            )
            logger.debug("InsertFromFile COM error: %r", exc)
            pump_office_ui(1.0 * attempt)
    raise RuntimeError(f"Could not insert slide from {temp_ppt}.") from last_error
