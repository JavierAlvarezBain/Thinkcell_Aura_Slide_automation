from __future__ import annotations

import logging
import shutil
import tempfile
from pathlib import Path
from typing import Any

import pythoncom

from shared import (
    BuildConfig,
    MSO_FALSE,
    MSO_TRUE,
    apply_company_order,
    build_company_order,
    company_sort_settings_from_config,
    normalize_company_segment_sort_direction,
    normalize_company_sort_order,
    get_output_path,
    kill_office_processes,
    require_file,
    safe_text,
    stage,
)
from fte_quartile_table_detector import read_source_data
from chart_updater import (
    dump_shapes_for_debug,
    get_thinkcell_addin,
    keep_only_chart_slide,
    run_com_check,
    start_excel_app,
    start_powerpoint_app,
    update_slide_text,
    update_thinkcell_chart,
    write_chart_data_workbook,
)
from quartile_line_handler import (
    GEOMETRY_CARRYOVER_BUILD_ID,
    patch_l1_quartiles_in_deck,
    settle_after_thinkcell_chart_update,
)


QUARTILE_LABELS = {
    "top": "Top",
    "second": "Second",
    "third": "Third",
    "bottom": "Bottom",
    "unknown": "Unknown",
}


def number_or_none(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str):
        cleaned = value.replace("$", "").replace(",", "").replace("\u00a0", "").strip()
        if cleaned in {"", "-", "NA", "N/A", "nan", "None", "<no data>"}:
            return None
        try:
            return float(cleaned)
        except Exception:
            return None
    try:
        return float(value)
    except Exception:
        return None


L1_SORT_ORDER_LABELS = {
    "revenue_ascending": "increasing revenue",
    "revenue_descending": "decreasing revenue",
    "sheet_order": "Excel sheet order",
    "segment": "segment order",
}


def normalize_l1_sort_order(value: Any) -> str:
    """Backward-compatible wrapper for the universal company sort normalizer."""
    return normalize_company_sort_order(value)


def normalize_l1_segment_sort_direction(value: Any) -> str:
    """Backward-compatible wrapper for the universal segment direction."""
    return normalize_company_segment_sort_direction(value)


def _normalized_company_key(value: Any) -> str:
    import re
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def order_l1_companies(
    company_headers: list[Any],
    revenue_values: list[Any],
    records: list[dict[str, Any]],
    logger: logging.Logger,
    sort_order: str = "revenue_descending",
    segment_assignments: list[dict[str, Any]] | None = None,
    segment_sort_direction: str = "revenue_descending",
) -> tuple[list[Any], list[Any], list[dict[str, Any]]]:
    """Apply the universal client-first company order to Aura L1 records."""
    order = build_company_order(
        company_headers,
        revenue_values,
        sort_order=sort_order,
        reference_assignments=segment_assignments,
        segment_sort_direction=segment_sort_direction,
        logger=logger,
        context="Aura L1",
    )
    ordered_headers, ordered_revenue, ordered_records = apply_company_order(
        company_headers,
        order,
        revenue_values=revenue_values,
        records=records,
        record_value_key="fte_values",
    )
    return (
        ordered_headers,
        list(ordered_revenue or []),
        list(ordered_records or []),
    )


def order_l1_companies_by_revenue(
    company_headers: list[Any],
    revenue_values: list[Any],
    records: list[dict[str, Any]],
    logger: logging.Logger,
) -> tuple[list[Any], list[Any], list[dict[str, Any]]]:
    """Backward-compatible wrapper for the former descending-only behavior."""
    return order_l1_companies(
        company_headers,
        revenue_values,
        records,
        logger,
        sort_order="revenue_descending",
    )


def classify_client_quartile(record: dict, logger: logging.Logger) -> dict[str, Any]:
    """
    Classify the leftmost company/client against the function row's quartiles.

    Lower FTE/$B is better:
      client <= Top Q      -> Top
      client <= Median     -> Second
      client <= Bottom Q   -> Third
      client > Bottom Q    -> Bottom
    """
    fte_values = record.get("fte_values") or []
    quartile_values = record.get("quartile_values") or []

    client_value = number_or_none(fte_values[0]) if fte_values else None
    quartiles = sorted(value for value in (number_or_none(v) for v in quartile_values) if value is not None)

    if client_value is None or len(quartiles) < 3:
        logger.warning(
            "Could not classify client quartile for %s. client_value=%s quartiles=%s",
            record.get("function"),
            client_value,
            quartile_values,
        )
        return {
            "label": QUARTILE_LABELS["unknown"],
            "client_value": client_value,
            "top_q": None,
            "median": None,
            "bottom_q": None,
        }

    top_q, median, bottom_q = quartiles[0], quartiles[1], quartiles[2]

    if client_value <= top_q:
        label = QUARTILE_LABELS["top"]
    elif client_value <= median:
        label = QUARTILE_LABELS["second"]
    elif client_value <= bottom_q:
        label = QUARTILE_LABELS["third"]
    else:
        label = QUARTILE_LABELS["bottom"]

    return {
        "label": label,
        "client_value": client_value,
        "top_q": top_q,
        "median": median,
        "bottom_q": bottom_q,
    }


def filter_records(records: list[dict], config: BuildConfig, logger: logging.Logger) -> list[dict]:
    if config.only_function:
        wanted = config.only_function.strip().lower()
        records = [record for record in records if record["function"].strip().lower() == wanted]
        if not records:
            raise ValueError(f"No source row found for --only-function {config.only_function!r}")
        logger.info("Only-function mode: creating slide for %s", records[0]["function"])
    if config.max_slides is not None:
        records = records[: config.max_slides]
        logger.info("Debug mode: only creating %s slide(s)", config.max_slides)
    return records




def prepare_l1_source(
    config: BuildConfig,
    logger: logging.Logger,
) -> tuple[list[Any], list[Any], list[dict[str, Any]]]:
    """Read, order, and filter the exact Aura L1 source used for slide creation.

    Keeping this preparation in one function lets the integrated workflow reuse
    the identical record order when it patches L1 quartiles in the final closed
    consolidated deck.
    """
    company_headers, revenue_values, records = read_source_data(config, logger)
    sort_order, segment_direction, segment_assignments = (
        company_sort_settings_from_config(config)
    )
    company_headers, revenue_values, records = order_l1_companies(
        company_headers,
        revenue_values,
        records,
        logger,
        sort_order=sort_order,
        segment_assignments=segment_assignments,
        segment_sort_direction=segment_direction,
    )
    records = filter_records(records, config, logger)
    return company_headers, revenue_values, records


def append_or_create_final_deck(
    ppt,
    final_pres,
    final_ppt_temp: Path,
    temp_ppt: Path,
    config: BuildConfig,
    logger: logging.Logger,
):
    if final_pres is None:
        with stage(logger, "Create combined deck"):
            shutil.copyfile(temp_ppt, final_ppt_temp)
            return ppt.Presentations.Open(
                str(final_ppt_temp),
                MSO_FALSE,
                MSO_FALSE,
                MSO_TRUE if config.with_window else MSO_FALSE,
            )
    with stage(logger, "Append slide to combined deck"):
        final_pres.Slides.InsertFromFile(str(temp_ppt), final_pres.Slides.Count, 1, 1)
    return final_pres


def build_deck(
    config: BuildConfig,
    logger: logging.Logger,
    prepared_source: tuple[list[Any], list[Any], list[dict[str, Any]]] | None = None,
    *,
    finalize_quartiles: bool = True,
) -> Path:
    pythoncom.CoInitialize()
    run_temp_dir = Path(tempfile.mkdtemp(prefix="thinkcell_run_"))
    output_ppt = get_output_path(config)

    excel = None
    ppt = None
    final_pres = None
    current_pres = None
    success = False

    try:
        require_file(config.excel_path, "Excel workbook")
        require_file(config.template_ppt, "PowerPoint template")

        if config.kill_office_first:
            kill_office_processes(logger)

        if prepared_source is None:
            company_headers, revenue_values, records = prepare_l1_source(config, logger)
        else:
            company_headers, revenue_values, records = prepared_source

        if not company_headers:
            raise ValueError("No company headers were detected in the Aura Summary table.")
        client_company = str(company_headers[0]).strip() or "Client"
        logger.info("Client/company selected from leftmost Aura Summary company column: %s", client_company)

        logger.info("Temp folder: %s", run_temp_dir)
        logger.info("Output will be saved to: %s", output_ppt)
        logger.info("Build: %s", GEOMETRY_CARRYOVER_BUILD_ID)
        logger.info(
            "Architecture: fte_quartile_table_detector -> chart_updater -> quartile_line_handler -> slide_collator."
        )
        logger.info(
            "Carry-over mode: deferred manual. No CommandBars, MSAA, OLE, mouse, or keyboard automation."
        )

        with stage(logger, "Start Excel COM"):
            excel = start_excel_app(config, logger)
        with stage(logger, "Start PowerPoint COM"):
            ppt = start_powerpoint_app(config, logger)
        with stage(logger, "Get think-cell add-in"):
            tc_addin = get_thinkcell_addin(excel, ppt, logger)

        final_ppt_temp = run_temp_dir / "combined_working.pptx"

        for idx, record in enumerate(records, start=1):
            function_name = record["function"]
            l0_name = str(record.get("l0") or record.get("level0") or "").strip()
            client_quartile = classify_client_quartile(record, logger)
            quartile_label = client_quartile["label"]

            logger.info("========== Slide %s: %s / %s ==========", idx, l0_name or "<no L0>", function_name)
            logger.info(
                "Client quartile classification for %s: %s | client=%.4g top=%.4g median=%.4g bottom=%.4g",
                function_name,
                quartile_label,
                client_quartile["client_value"] if client_quartile["client_value"] is not None else float("nan"),
                client_quartile["top_q"] if client_quartile["top_q"] is not None else float("nan"),
                client_quartile["median"] if client_quartile["median"] is not None else float("nan"),
                client_quartile["bottom_q"] if client_quartile["bottom_q"] is not None else float("nan"),
            )

            with stage(logger, "Write chart data workbook"):
                chart_data_path = write_chart_data_workbook(
                    temp_dir=run_temp_dir,
                    company_headers=company_headers,
                    revenue_values=revenue_values,
                    record=record,
                    config=config,
                )

            temp_ppt = run_temp_dir / f"tc_slide_{idx}_{safe_text(function_name)}.pptx"

            with stage(logger, "Open PowerPoint template"):
                current_pres = ppt.Presentations.Open(
                    str(config.template_ppt),
                    MSO_TRUE,
                    MSO_FALSE,
                    MSO_TRUE if config.with_window else MSO_FALSE,
                )

            with stage(logger, "Update title, client tagline, quartile phrase, and Bain agenda text"):
                update_slide_text(
                    current_pres,
                    function_name,
                    l0_name,
                    client_company,
                    quartile_label,
                    config,
                    logger,
                )

            with stage(logger, "Update think-cell chart"):
                update_thinkcell_chart(
                    tc_addin=tc_addin,
                    pres=current_pres,
                    excel_app=excel,
                    chart_data_path=chart_data_path,
                    config=config,
                    logger=logger,
                )

            with stage(logger, "Settle chart axes after think-cell update"):
                settle_after_thinkcell_chart_update(current_pres, config, logger)

            dump_shapes_for_debug(current_pres, config, logger)

            with stage(logger, "Keep only chart slide"):
                keep_only_chart_slide(current_pres, config)

            with stage(logger, "Save temporary PPT"):
                current_pres.SaveAs(str(temp_ppt))
                current_pres.Close()
                current_pres = None

            if config.update_quartiles:
                logger.debug(
                    "Aura L1 quartile patch for %s is deferred until the complete "
                    "section deck has been saved and closed.",
                    function_name,
                )
            else:
                logger.info("Quartile geometry patch disabled by --skip-quartiles.")

            final_pres = append_or_create_final_deck(
                ppt=ppt,
                final_pres=final_pres,
                final_ppt_temp=final_ppt_temp,
                temp_ppt=temp_ppt,
                config=config,
                logger=logger,
            )

        with stage(logger, "Save and close Aura L1 section deck"):
            final_pres.SaveAs(str(output_ppt))
            final_pres.Close()
            final_pres = None

        if config.update_quartiles and finalize_quartiles:
            with stage(logger, "Finalize Aura L1 quartiles in the closed section deck"):
                patch_l1_quartiles_in_deck(
                    output_ppt=output_ppt,
                    slide_records=records,
                    logger=logger,
                )
        elif config.update_quartiles:
            logger.info(
                "Aura L1 section shell is complete. Native quartile carry-over data "
                "is intentionally deferred until the consolidated deck has been "
                "merged, saved, and closed."
            )

        success = True
        logger.info("Created deck: %s", output_ppt)
        if config.update_quartiles and finalize_quartiles:
            logger.warning(
                "Aura L1 native quartile values and geometry are verified in the closed deck. "
                "Manual think-cell carry-over is still required: open the deck; for each "
                "slide click the chart, choose 'Manually carry over changes to think-cell', "
                "click Finish, then save the deck."
            )
        return output_ppt

    finally:
        try:
            if current_pres is not None:
                current_pres.Close()
        except Exception:
            pass
        try:
            if final_pres is not None:
                final_pres.Close()
        except Exception:
            pass
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
        if success and not config.keep_temp:
            shutil.rmtree(run_temp_dir, ignore_errors=True)
        else:
            logger.warning("Temp folder kept for debugging: %s", run_temp_dir)
        pythoncom.CoUninitialize()


__all__ = ["build_deck", "prepare_l1_source", "run_com_check"]