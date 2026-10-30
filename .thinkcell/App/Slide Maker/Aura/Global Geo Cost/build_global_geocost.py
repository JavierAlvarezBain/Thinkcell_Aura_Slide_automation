from __future__ import annotations

import logging
import re
import shutil
import tempfile
import time
import zipfile
from pathlib import Path

from geocost_common import (
    BUILD_ID,
    GeoCostData,
    GeoCostScope,
    GeoCostSourceLayout,
    MSO_FALSE,
    MSO_TRUE,
    close_presentation_with_retry,
    pump_office_ui,
    pythoncom,
    require_windows_com,
    retry_office_call,
    safe_filename,
    save_presentation_as_with_retry,
    wait_for_saved_file_stable,
)
from geocost_office import (
    apply_centered_labels_to_presentation_com,
    insert_slide_with_retry,
    reacquire_thinkcell_addin,
    start_excel_app,
    start_powerpoint_app,
    update_slide_text,
    update_thinkcell_chart,
)
from geocost_source import (
    choose_function_blocks,
    detect_geocost_sheets,
    detect_source_layout,
    extract_template_layout,
    find_function_blocks,
    parse_geocost_block,
    read_sheet_values,
)


def _resolve_sheet_and_layout(
    excel_path: Path,
    scope: GeoCostScope,
    sheet_name: str | None,
    source_layout: GeoCostSourceLayout | None,
    logger: logging.Logger,
) -> tuple[str, list[list[object]], GeoCostSourceLayout]:
    if sheet_name is None:
        detection = detect_geocost_sheets(excel_path, logger=logger)
        match = detection.global_sheet if scope == "global" else detection.us_sheet
        if match is None:
            raise ValueError(f"No {scope.upper()} Geo Cost worksheet was detected in {excel_path.name}.")
        sheet_name = match.sheet_name
        source_layout = match.source_layout

    rows = read_sheet_values(excel_path, sheet_name)
    if source_layout is None:
        detected = detect_source_layout(rows)
        if detected is None:
            raise ValueError(
                f"Worksheet {sheet_name!r} does not match the expected Geo Cost structure. "
                "Expected >>> function markers, Region, High/Low, and Country/State headers."
            )
        source_layout = detected[0]
    return sheet_name, rows, source_layout


def prepare_data_workbooks(
    excel_path: Path,
    template_path: Path,
    output_dir: Path,
    requested_function: str,
    max_functions: int | None,
    max_companies: int,
    label_threshold: float,
    logger: logging.Logger,
    *,
    scope: GeoCostScope = "global",
    sheet_name: str | None = None,
    source_layout: GeoCostSourceLayout | None = None,
    chart_name: str = "",
    company_sort_order: str = "sheet_order",
    company_segment_sort_direction: str = "revenue_descending",
    company_sort_assignments: list[dict] | None = None,
) -> list[GeoCostData]:
    sheet_name, rows, source_layout = _resolve_sheet_and_layout(
        excel_path, scope, sheet_name, source_layout, logger
    )
    template_layout = extract_template_layout(template_path, chart_name, logger)
    blocks = find_function_blocks(rows, source_layout)
    if not blocks:
        raise ValueError(
            f"No Geo Cost function blocks were found on {sheet_name!r}; "
            f"expected >>> markers on row {source_layout.marker_row}."
        )
    selected = choose_function_blocks(blocks, requested_function, max_functions)
    output_dir.mkdir(parents=True, exist_ok=True)
    return [
        parse_geocost_block(
            scope,
            sheet_name,
            rows,
            source_layout,
            block,
            max_companies,
            label_threshold,
            template_layout,
            output_dir,
            logger,
            company_sort_order=company_sort_order,
            company_segment_sort_direction=company_segment_sort_direction,
            company_sort_assignments=company_sort_assignments,
        )
        for block in selected
    ]


_SLIDE_PART_RE = re.compile(r"^ppt/slides/slide[0-9]+[.]xml$")


def _wait_for_closed_pptx_copy(
    path: Path,
    expected_slides: int,
    logger: logging.Logger,
    timeout_seconds: float = 90.0,
) -> None:
    """Verify that SaveCopyAs produced an independent, readable PPTX package.

    PowerPoint may keep the live working presentation open, but it must not own
    the section file returned to the integrated workflow.  The caller can safely
    register and merge the section only after this check succeeds.
    """
    deadline = time.monotonic() + max(1.0, float(timeout_seconds))
    last_error: BaseException | None = None
    while time.monotonic() < deadline:
        try:
            with zipfile.ZipFile(path, "r") as archive:
                archive.read("[Content_Types].xml")
                bad_member = archive.testzip()
                if bad_member is not None:
                    raise zipfile.BadZipFile(
                        f"PPTX member failed CRC validation: {bad_member}"
                    )
                slide_count = sum(
                    1 for name in archive.namelist() if _SLIDE_PART_RE.fullmatch(name)
                )
            if slide_count != int(expected_slides):
                raise RuntimeError(
                    f"Geo Cost section contains {slide_count} slide(s); "
                    f"expected {expected_slides}."
                )
            logger.info(
                "Verified closed Geo Cost section copy: %s (%s slides).",
                path,
                slide_count,
            )
            return
        except (PermissionError, OSError, KeyError, zipfile.BadZipFile, RuntimeError) as exc:
            last_error = exc
            pump_office_ui(0.25)
            time.sleep(0.25)
    raise RuntimeError(
        f"Geo Cost section copy remained locked or unreadable after "
        f"{timeout_seconds:.0f}s: {path}"
    ) from last_error


def build_geocost_deck(
    excel_path: Path,
    template_path: Path,
    output_path: Path,
    requested_function: str,
    max_functions: int | None,
    max_companies: int,
    chart_name: str,
    label_threshold: float,
    keep_temp: bool,
    quit_office: bool,
    post_chart_wait: float,
    logger: logging.Logger,
    *,
    scope: GeoCostScope = "global",
    sheet_name: str | None = None,
    source_layout: GeoCostSourceLayout | None = None,
    company_sort_order: str = "sheet_order",
    company_segment_sort_direction: str = "revenue_descending",
    company_sort_assignments: list[dict] | None = None,
) -> Path:
    """Build one Global or US Geo Cost deck from a detected worksheet and template."""
    require_windows_com()
    sheet_name, rows, source_layout = _resolve_sheet_and_layout(
        excel_path, scope, sheet_name, source_layout, logger
    )
    template_layout = extract_template_layout(template_path, chart_name, logger)
    blocks = find_function_blocks(rows, source_layout)
    if not blocks:
        raise ValueError(
            f"No Geo Cost function blocks were found on {sheet_name!r}; "
            f"expected >>> markers on row {source_layout.marker_row}."
        )
    selected = choose_function_blocks(blocks, requested_function, max_functions)

    temp_dir = Path(tempfile.mkdtemp(prefix=f"geocost_{scope}_"))
    final_ppt_temp = temp_dir / f"combined_geocost_{scope}_working.pptx"
    excel_app = None
    ppt_app = None
    final_pres = None
    current_pres = None
    success = False
    built_data: list[GeoCostData] = []

    pythoncom.CoInitialize()
    try:
        excel_app = start_excel_app(logger)
        ppt_app = start_powerpoint_app(logger)
        tc_addin = reacquire_thinkcell_addin(excel_app, ppt_app, logger)
        logger.info("Build: %s", BUILD_ID)
        logger.info(
            "Geo Cost scope=%s | sheet=%s | template=%s | functions=%s",
            scope.upper(),
            sheet_name,
            template_path.name,
            [block.name for block in selected],
        )
        logger.info(
            "Flow: detect source sheet/layout -> read template entity order -> build complete ChartData -> "
            "one AddRangeData update per function -> combine slides -> add and verify centered >= %.1f%% labels.",
            label_threshold,
        )

        for slide_index, block in enumerate(selected, start=1):
            logger.info(
                "========== %s Geo Cost slide %s: %s ==========",
                scope.upper(),
                slide_index,
                block.name,
            )
            data = parse_geocost_block(
                scope,
                sheet_name,
                rows,
                source_layout,
                block,
                max_companies,
                label_threshold,
                template_layout,
                temp_dir,
                logger,
                company_sort_order=company_sort_order,
                company_segment_sort_direction=company_segment_sort_direction,
                company_sort_assignments=company_sort_assignments,
            )
            temp_ppt = temp_dir / (
                f"geocost_{scope}_slide_{slide_index}_{safe_filename(block.name)}.pptx"
            )
            current_pres = ppt_app.Presentations.Open(
                str(template_path),
                MSO_TRUE,
                MSO_FALSE,
                MSO_TRUE,
            )
            update_slide_text(current_pres, data, logger)
            tc_addin = reacquire_thinkcell_addin(excel_app, ppt_app, logger)
            tc_addin = update_thinkcell_chart(
                tc_addin,
                current_pres,
                excel_app,
                ppt_app,
                data.chart_data_path,
                template_layout,
                1 + len(data.companies),
                template_layout.chart_name,
                logger,
                post_chart_wait,
            )
            pump_office_ui(max(0.8, post_chart_wait / 2.0))

            if final_pres is None:
                save_presentation_as_with_retry(current_pres, final_ppt_temp, logger)
                final_pres = current_pres
                current_pres = None
                logger.info("Initialized live %s Geo Cost combined deck.", scope.upper())
            else:
                save_presentation_as_with_retry(current_pres, temp_ppt, logger)
                close_presentation_with_retry(
                    current_pres,
                    logger,
                    f"temporary {scope} Geo Cost slide {slide_index}",
                    raise_on_failure=True,
                )
                current_pres = None
                insert_slide_with_retry(final_pres, temp_ppt, logger)
            built_data.append(data)
            pump_office_ui(0.5)

        if final_pres is None:
            raise RuntimeError(f"No {scope.upper()} Geo Cost slides were generated.")

        logger.info("Applying final centered labels on the live %s Geo Cost deck...", scope.upper())
        apply_centered_labels_to_presentation_com(final_pres, built_data, logger)
        pump_office_ui(max(0.8, post_chart_wait / 2.0))

        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.unlink(missing_ok=True)

        # Do not SaveAs the live working presentation onto the section output.
        # SaveAs makes PowerPoint own/lock output_path until Close succeeds.
        # SaveCopyAs leaves final_pres attached to its local temporary working
        # file and publishes an independent section package for registration.
        retry_office_call(
            f"SaveCopyAs {output_path.name}",
            lambda: final_pres.SaveCopyAs(str(output_path)),
            logger,
            max_attempts=8,
            base_delay=1.0,
        )
        wait_for_saved_file_stable(output_path, logger)
        _wait_for_closed_pptx_copy(
            output_path,
            expected_slides=len(built_data),
            logger=logger,
            timeout_seconds=90.0,
        )

        # Closing the disposable working presentation is still attempted, but
        # the returned section file is no longer the file held by that window.
        # Keep the COM reference alive when Close fails so the finally block can
        # retry instead of silently abandoning an open presentation.
        closed = close_presentation_with_retry(
            final_pres,
            logger,
            f"final {scope} Geo Cost working presentation",
            max_attempts=10,
            raise_on_failure=False,
        )
        if closed:
            final_pres = None
        else:
            logger.warning(
                "PowerPoint kept only the disposable Geo Cost working presentation "
                "open. The returned section is an independent verified copy, so "
                "section registration can continue without a file lock."
            )

        success = True
        logger.info("Created %s Geo Cost deck: %s", scope.upper(), output_path)
        return output_path
    finally:
        if current_pres is not None:
            close_presentation_with_retry(
                current_pres,
                logger,
                "current temporary Geo Cost presentation",
                max_attempts=3,
                raise_on_failure=False,
            )
        if final_pres is not None:
            close_presentation_with_retry(
                final_pres,
                logger,
                "combined Geo Cost working presentation",
                max_attempts=3,
                raise_on_failure=False,
            )
        if quit_office:
            try:
                if ppt_app is not None:
                    ppt_app.Quit()
            except Exception:
                pass
            try:
                if excel_app is not None:
                    excel_app.Quit()
            except Exception:
                pass
        if keep_temp or not success:
            logger.warning("Geo Cost temp folder kept: %s", temp_dir)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)
        pythoncom.CoUninitialize()


def main(argv: list[str] | None = None) -> None:
    """Backward-compatible CLI entry point; the application runner owns argument parsing."""
    from run_geocost import main as run_geocost_main

    run_geocost_main(argv)


__all__ = ["build_geocost_deck", "prepare_data_workbooks", "main"]


if __name__ == "__main__":
    main()
