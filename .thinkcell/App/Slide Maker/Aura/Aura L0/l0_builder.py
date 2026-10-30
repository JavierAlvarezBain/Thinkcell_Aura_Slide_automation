from __future__ import annotations

"""Aura L0 section builder and standalone recovery orchestration.

The builder creates an L0-only section deck through the accepted think-cell
AddRangeData loop.  Callers choose whether to finalize the native think-cell
state immediately (standalone deck) or defer it until after a larger deck has
been merged and closed (the consolidated app.py workflow).
"""

from dataclasses import replace
from pathlib import Path
import shutil
import tempfile
import zipfile
from typing import Any

from l0_config import (
    CHART_HEADER_ROW,
    CHART_NAME,
    DEFAULT_TEMPLATE_PPT,
    FIRST_SEGMENT_ROW,
    L0_BUILD_ID,
    L0BuildResult,
    REVENUE_CHART_ROW,
    column_letter,
)
from l0_native import (
    capture_l0_quartile_anchor_from_pptx,
    patch_l0_slides_in_deck,
    settle_after_thinkcell_chart_update,
)
from l0_source import (
    build_l0_slide_records,
    filter_l0_slide_records,
    segment_display_name,
    segment_style_log_items,
    write_l0_chart_data_workbook,
)
from shared import (
    BuildConfig,
    MSO_FALSE,
    MSO_TRUE,
    apply_company_order,
    build_company_order,
    clean_text,
    company_sort_settings_from_config,
    get_output_path,
    kill_office_processes,
    pump_office_ui,
    require_file,
    safe_text,
    stage,
)
from fte_quartile_table_detector import read_source_data


L0_BUILDER_INTEGRATION_ID = "l0-staged-template-same-session-v2026-08-10-r1"


def _stage_l0_template_copy(source_ppt: Path, run_temp_dir: Path, logger) -> Path:
    """Copy the L0 template once and use only that local copy for this run.

    The L0 native-anchor reader opens the PPTX as a ZIP package. PowerPoint or
    OneDrive can keep the project copy locked even though PowerPoint can still
    open it. Staging once before the chart loop removes that race without any
    watchdog, process termination, or alternate PowerPoint instance.
    """
    source_ppt = Path(source_ppt).resolve()
    runtime_ppt = run_temp_dir / f"runtime_{source_ppt.name}"
    with stage(logger, "Stage Aura L0 template locally"):
        shutil.copy2(source_ppt, runtime_ppt)
        with zipfile.ZipFile(runtime_ppt, "r") as archive:
            archive.read("[Content_Types].xml")
            bad_member = archive.testzip()
            if bad_member is not None:
                raise zipfile.BadZipFile(
                    f"The staged Aura L0 template failed CRC validation at {bad_member}."
                )
        logger.info(
            "Aura L0 will read and open only the staged template copy: %s",
            runtime_ppt,
        )
    return runtime_ppt


def make_l0_runtime_config(
    base_config: BuildConfig,
    *,
    template_ppt: Path | None = None,
    output_ppt: Path | None = None,
    update_quartiles: bool | None = None,
    post_chart_wait_seconds: float | None = None,
) -> BuildConfig:
    """Clone shared workflow settings into an L0-specific runtime config."""
    template = Path(template_ppt or getattr(base_config, "l0_template_ppt", None) or DEFAULT_TEMPLATE_PPT).resolve()
    destination = Path(output_ppt or get_output_path(base_config)).resolve()
    quartiles_enabled = (
        bool(update_quartiles)
        if update_quartiles is not None
        else bool(getattr(base_config, "l0_update_quartiles", base_config.update_quartiles))
    )
    wait_seconds = (
        float(post_chart_wait_seconds)
        if post_chart_wait_seconds is not None
        else float(
            getattr(
                base_config,
                "l0_post_chart_wait_seconds",
                base_config.post_chart_update_wait_seconds,
            )
        )
    )

    return replace(
        base_config,
        template_ppt=template,
        output_dir=destination.parent,
        output_ppt_name=destination.name,
        chart_name=CHART_NAME,
        chart_header_row=CHART_HEADER_ROW,
        revenue_chart_row=REVENUE_CHART_ROW,
        fte_chart_row=FIRST_SEGMENT_ROW,
        chart_rows=11,
        chart_cols=11,
        chart_slide_search_text="FTEs per $B Revenue",
        max_slides=None,
        only_function=None,
        update_quartiles=quartiles_enabled,
        update_agenda_text=True,
        post_chart_update_wait_seconds=max(0.3, wait_seconds),
        carryover_mode="defer",
        include_geocost=False,
    )


def prepare_l0_records(
    config: BuildConfig,
    logger,
    *,
    only_l0: str | None = None,
    max_slides: int | None = None,
) -> tuple[list[str], list[Any], list[dict[str, Any]], str]:
    """Read the shared Aura table and create the complete L0-specific slide plan."""
    company_headers, revenue_values, l1_records = read_source_data(config, logger)
    if not company_headers:
        raise ValueError("No company headers were detected on the Aura Summary worksheet.")

    sort_order, segment_direction, segment_assignments = (
        company_sort_settings_from_config(config)
    )
    company_order = build_company_order(
        company_headers,
        revenue_values,
        sort_order=sort_order,
        reference_assignments=segment_assignments,
        segment_sort_direction=segment_direction,
        logger=logger,
        context="Aura L0",
    )
    company_headers, ordered_revenue, ordered_records = apply_company_order(
        company_headers,
        company_order,
        revenue_values=revenue_values,
        records=l1_records,
        record_value_key="fte_values",
    )
    revenue_values = list(ordered_revenue or [])
    l1_records = list(ordered_records or [])
    client_company = clean_text(company_headers[0]) or "Client"
    slide_records = build_l0_slide_records(
        company_headers=company_headers,
        records=l1_records,
        config=config,
        logger=logger,
    )
    slide_records = filter_l0_slide_records(
        slide_records,
        only_l0=only_l0,
        max_slides=max_slides,
    )
    return company_headers, revenue_values, slide_records, client_company


def patch_existing_l0_deck_from_source(
    config: BuildConfig,
    existing_ppt: Path,
    logger,
    *,
    only_l0: str | None = None,
    max_slides: int | None = None,
) -> Path:
    """Patch native colors and quartiles in an already-generated L0-only deck."""
    existing_ppt = Path(existing_ppt).resolve()
    require_file(config.excel_path, "Excel workbook")
    require_file(config.template_ppt, "Aura L0 PowerPoint template")
    require_file(existing_ppt, "Existing Aura L0 PowerPoint deck")

    _company_headers, _revenue_values, slide_records, _client_company = prepare_l0_records(
        config,
        logger,
        only_l0=only_l0,
        max_slides=max_slides,
    )
    template_anchor = capture_l0_quartile_anchor_from_pptx(
        config.template_ppt,
        logger,
        context="L0 template recovery anchor",
    )
    for slide_record in slide_records:
        slide_record["quartile_anchor"] = template_anchor

    logger.info(
        "Recovery mode: patch native think-cell segment colors and quartiles only; "
        "no Excel COM, PowerPoint COM, chart rebuild, or live recoloring."
    )
    logger.info("Existing Aura L0 deck: %s", existing_ppt)
    logger.info("Expected slide order: %s", [record["l0"] for record in slide_records])
    patch_l0_slides_in_deck(
        output_ppt=existing_ppt,
        slide_records=slide_records,
        logger=logger,
    )
    logger.info("Recovered Aura L0 deck in place: %s", existing_ppt)
    return existing_ppt


def build_l0_section_deck(
    base_config: BuildConfig,
    logger,
    *,
    output_ppt: Path | None = None,
    template_ppt: Path | None = None,
    only_l0: str | None = None,
    max_slides: int | None = None,
    update_quartiles: bool | None = None,
    post_chart_wait_seconds: float | None = None,
    finalize_native_patch: bool = True,
) -> L0BuildResult:
    """Run the complete accepted L0 creation loop and return its patch metadata.

    ``finalize_native_patch=False`` is required when the L0 deck will be inserted
    into another presentation.  The returned records must then be passed to
    :func:`l0_native.patch_l0_slides_in_deck` after the consolidated deck has
    been saved and PowerPoint has closed it.
    """
    try:
        import pythoncom
    except ImportError as exc:  # pragma: no cover - Windows/Office dependency
        raise RuntimeError(
            "pywin32 is not available. Run Aura L0 with the same "
            ".thinkcell\\Scripts\\python.exe used by app.py."
        ) from exc

    from chart_updater import (
        dump_shapes_for_debug,
        get_thinkcell_addin,
        keep_only_chart_slide,
        start_excel_app,
        start_powerpoint_app,
        update_thinkcell_chart,
    )
    from l0_powerpoint import append_or_create_combined_deck, update_l0_slide_text_safe

    config = make_l0_runtime_config(
        base_config,
        template_ppt=template_ppt,
        output_ppt=output_ppt,
        update_quartiles=update_quartiles,
        post_chart_wait_seconds=post_chart_wait_seconds,
    )
    pythoncom.CoInitialize()
    run_temp_dir = Path(tempfile.mkdtemp(prefix="aura_l0_"))
    output_path = get_output_path(config)
    runtime_template_ppt: Path | None = None

    excel = None
    ppt = None
    combined_pres = None
    current_pres = None
    success = False

    try:
        require_file(config.excel_path, "Excel workbook")
        require_file(config.template_ppt, "Aura L0 PowerPoint template")
        config.output_dir.mkdir(parents=True, exist_ok=True)

        if config.kill_office_first:
            kill_office_processes(logger)

        runtime_template_ppt = _stage_l0_template_copy(
            config.template_ppt,
            run_temp_dir,
            logger,
        )

        company_headers, revenue_values, slide_records, client_company = prepare_l0_records(
            config,
            logger,
            only_l0=only_l0,
            max_slides=max_slides,
        )

        logger.info("Client/company: %s", client_company)
        logger.info("Detected L0 slide order: %s", [record["l0"] for record in slide_records])
        logger.info("Revenue row values: %s", revenue_values)
        logger.info("L0 temp folder: %s", run_temp_dir)
        logger.info("L0 section output: %s", output_path)
        logger.info("L0 build: %s", L0_BUILD_ID)
        logger.info("L0 integration build: %s", L0_BUILDER_INTEGRATION_ID)
        logger.info(
            "L0 flow: read Aura Summary -> group L1 rows by L0 -> dedicated L0 quartiles if available, "
            "otherwise peer percentiles excluding client -> dynamic think-cell range from row 6 onward."
        )
        logger.info(
            "L0 segment row order: Other G&A/Other Ops first when present; remaining rows ascending by "
            "the first-company client value so the largest client segment renders at the bottom."
        )
        logger.info(
            "L0 colors: G&A=Forest, S&M=Sky, Ops=Sunset, R&D=Berry, OOS/Unmapped=Graphite. "
            "Other G&A/Other Ops remain patterned. Six-to-ten ordinary rows use dynamically interpolated "
            "Bain shades; every ordinary row beyond ten repeats the dynamic scale's lightest family shade."
        )
        logger.info(
            "L0 native state order: chart shell -> PowerPoint save/close -> native scalar/series colors -> "
            "native legend -> quartiles last -> no later PowerPoint save."
        )

        template_quartile_anchor = None
        if config.update_quartiles:
            with stage(logger, "Capture L0 template quartile identities"):
                if runtime_template_ppt is None:
                    raise RuntimeError("Aura L0 runtime template was not staged.")
                template_quartile_anchor = capture_l0_quartile_anchor_from_pptx(
                    runtime_template_ppt,
                    logger,
                    context="staged L0 template",
                )

        with stage(logger, "Start Excel COM for Aura L0"):
            excel = start_excel_app(config, logger)
        with stage(logger, "Start PowerPoint COM for Aura L0"):
            ppt = start_powerpoint_app(config, logger)
        with stage(logger, "Get think-cell add-in for Aura L0"):
            tc_addin = get_thinkcell_addin(excel, ppt, logger)

        combined_working_path = run_temp_dir / "Aura_L0_combined_working.pptx"

        for slide_index, slide_record in enumerate(slide_records, start=1):
            l0_name = slide_record["l0"]
            segments = slide_record["segments"]
            quartiles = slide_record["quartile_values"]
            classification = slide_record["client_quartile"]
            quartile_label = str(classification["label"])

            logger.info("========== L0 slide %s: %s ==========", slide_index, l0_name)
            logger.info(
                "L1 rows in stack / ChartData rows: %s",
                [
                    f"row {FIRST_SEGMENT_ROW + idx}: {segment_display_name(segment)}"
                    for idx, segment in enumerate(segments)
                ],
            )
            logger.info("L0 segment Bain color plan: %s", segment_style_log_items(l0_name, segments))
            logger.info("L0 totals by company: %s", slide_record["fte_values"])
            logger.info(
                "Quartiles for %s: top=%.4g median=%.4g bottom=%.4g | source=%s",
                l0_name,
                float(classification["top_q"]),
                float(classification["median"]),
                float(classification["bottom_q"]),
                slide_record["quartile_source"],
            )
            logger.info(
                "Client quartile classification for %s: %s | client=%.4g",
                l0_name,
                quartile_label,
                float(classification["client_value"])
                if classification["client_value"] is not None
                else float("nan"),
            )

            with stage(logger, "Write L0 chart data workbook"):
                chart_data_path = write_l0_chart_data_workbook(
                    temp_dir=run_temp_dir,
                    company_headers=company_headers,
                    revenue_values=revenue_values,
                    slide_record=slide_record,
                    config=config,
                )
                logger.info(
                    "ChartData range for %s: A1:%s%s (%s stacked series starting at row 6)",
                    l0_name,
                    column_letter(config.chart_cols),
                    config.chart_rows,
                    len(segments),
                )

            temp_ppt = run_temp_dir / f"l0_slide_{slide_index}_{safe_text(l0_name)}.pptx"

            with stage(logger, "Open staged Aura L0 PowerPoint template"):
                if runtime_template_ppt is None:
                    raise RuntimeError("Aura L0 runtime template was not staged.")
                current_pres = ppt.Presentations.Open(
                    str(runtime_template_ppt),
                    MSO_TRUE,
                    MSO_FALSE,
                    MSO_TRUE if config.with_window else MSO_FALSE,
                )

            with stage(logger, "Update L0 title, client, quartile phrase, and agenda text"):
                update_l0_slide_text_safe(
                    current_pres,
                    l0_name,
                    client_company,
                    quartile_label,
                    config,
                    logger,
                )

            with stage(logger, "Update L0 think-cell chart"):
                update_thinkcell_chart(
                    tc_addin=tc_addin,
                    pres=current_pres,
                    excel_app=excel,
                    chart_data_path=chart_data_path,
                    config=config,
                    logger=logger,
                )

            with stage(logger, "Settle L0 chart axes after think-cell update"):
                settle_after_thinkcell_chart_update(current_pres, config, logger)

            dump_shapes_for_debug(current_pres, config, logger)

            with stage(logger, "Keep only Aura L0 chart slide"):
                keep_only_chart_slide(current_pres, config)

            with stage(logger, "Save temporary Aura L0 slide"):
                current_pres.SaveAs(str(temp_ppt))
                current_pres.Close()
                current_pres = None

            if config.update_quartiles:
                with stage(logger, "Capture pre-color L0 quartile line identities"):
                    slide_record["quartile_anchor"] = capture_l0_quartile_anchor_from_pptx(
                        temp_ppt,
                        logger,
                        context=f"{l0_name} temporary chart",
                        preferred_anchor=template_quartile_anchor,
                    )
                logger.info(
                    "L0 quartile values retained for the final closed-deck patch: %s",
                    [round(float(value), 4) for value in quartiles],
                )
            else:
                logger.info("L0 native color/legend/quartile patch disabled for this build.")

            combined_pres = append_or_create_combined_deck(
                ppt=ppt,
                combined_pres=combined_pres,
                combined_working_path=combined_working_path,
                slide_path=temp_ppt,
                config=config,
                logger=logger,
            )

        with stage(logger, "Save and close Aura L0 section deck shell"):
            combined_pres.SaveAs(str(output_path))
            pump_office_ui(0.5)
            combined_pres.Close()
            combined_pres = None

        native_patch_pending = bool(config.update_quartiles and not finalize_native_patch)
        if config.update_quartiles and finalize_native_patch:
            with stage(logger, "Finalize Aura L0 native colors, legend, and quartiles"):
                patch_l0_slides_in_deck(
                    output_ppt=output_path,
                    slide_records=slide_records,
                    logger=logger,
                )
            native_patch_pending = False
        elif native_patch_pending:
            logger.info(
                "Aura L0 section shell is complete. Native L0 state is intentionally deferred until "
                "the consolidated deck has been merged, saved, and closed."
            )

        success = True
        logger.info("Created Aura L0 section deck: %s", output_path)
        logger.info("Aura L0 slides created: %s", len(slide_records))
        if config.update_quartiles and finalize_native_patch:
            logger.warning(
                "The deck contains verified native L0 color, legend, and quartile carry-over data. "
                "For each L0 slide, choose 'Manually carry over changes to think-cell', click Finish, "
                "and save the deck."
            )
        return L0BuildResult(
            deck_path=output_path,
            slide_records=slide_records,
            client_company=client_company,
            native_patch_pending=native_patch_pending,
        )

    finally:
        try:
            if current_pres is not None:
                current_pres.Close()
        except Exception:
            pass
        try:
            if combined_pres is not None:
                combined_pres.Close()
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
            logger.warning("L0 temp folder kept for debugging: %s", run_temp_dir)
        pythoncom.CoUninitialize()


# Backward-compatible name used by older L0 callers.
def build_l0_deck(config: BuildConfig, args, logger) -> Path:
    result = build_l0_section_deck(
        config,
        logger,
        template_ppt=config.template_ppt,
        only_l0=getattr(args, "only_l0", None),
        max_slides=getattr(args, "max_slides", None),
        update_quartiles=config.update_quartiles,
        post_chart_wait_seconds=config.post_chart_update_wait_seconds,
        finalize_native_patch=True,
    )
    return result.deck_path
