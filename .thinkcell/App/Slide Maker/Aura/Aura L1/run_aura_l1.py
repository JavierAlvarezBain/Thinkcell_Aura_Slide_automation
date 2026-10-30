from __future__ import annotations

import argparse
import faulthandler
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

WORKFLOW_DIR = Path(__file__).parent.resolve()
AURA_ROOT = WORKFLOW_DIR.parent.resolve()
APP_ROOT = AURA_ROOT.parent.resolve()
L0_DIR = AURA_ROOT / "Aura L0"
SUMMARY_DIR = AURA_ROOT / "Aura Summary"
GEOCOST_DIR = AURA_ROOT / "Global Geo Cost"

# Import priority: workflow-specific modules first, then sibling workflow modules,
# then universal Aura root modules.
for path in [
    str(APP_ROOT),
    str(AURA_ROOT),
    str(L0_DIR),
    str(SUMMARY_DIR),
    str(GEOCOST_DIR),
    str(WORKFLOW_DIR),
]:
    while path in sys.path:
        sys.path.remove(path)
sys.path.insert(0, str(APP_ROOT))
sys.path.insert(0, str(AURA_ROOT))
sys.path.insert(0, str(L0_DIR))
sys.path.insert(0, str(SUMMARY_DIR))
sys.path.insert(0, str(GEOCOST_DIR))
sys.path.insert(0, str(WORKFLOW_DIR))

from shared import (  # noqa: E402
    BuildConfig,
    DEFAULT_GEOCOST_GLOBAL_TEMPLATE_PPT,
    DEFAULT_GEOCOST_US_TEMPLATE_PPT,
    DEFAULT_L0_TEMPLATE_PPT,
    combine_presentation_files,
    get_output_path,
    kill_office_processes,
    presentation_slide_count,
    pythoncom,
    require_file,
    running_powerpoint_process_ids,
    setup_logging,
    stage,
)
from input import resolve_path, resolve_run_inputs  # noqa: E402
from aura_summary_builder import build_aura_summary_slide  # noqa: E402
from build_global_geocost import build_geocost_deck  # noqa: E402
from geocost_common import DEFAULT_LABEL_THRESHOLD  # noqa: E402
from geocost_source import detect_geocost_sheets  # noqa: E402


DEFAULT_TEMPLATE_PPT = WORKFLOW_DIR / "ThinkCell_Test.pptx"
DEFAULT_OUTPUT_DIR = WORKFLOW_DIR / "output"
DEFAULT_COM_CHECK_EXCEL = WORKFLOW_DIR / "Aura_Summary_Test.xlsx"
DEFAULT_SUMMARY_TEMPLATE_PPT = SUMMARY_DIR / "Table Test.pptx"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate Aura Summary + Aura L0 + Aura L1 slides and append detected Global/US Geo Cost slides."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument("--excel", type=Path, default=None)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_PPT)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--output-name", default="Aura_Integrated_Generated_Deck.pptx")

    parser.add_argument("--interactive-input", action="store_true")
    parser.add_argument("--no-input-prompt", action="store_true")
    parser.add_argument("--sheet-name", default=None)
    parser.add_argument(
        "--taxonomy-sheet",
        default=None,
        help=(
            "Worksheet containing the L0/L1 taxonomy columns. When omitted, "
            "the Aura L1 reader auto-detects a tab whose name contains "
            "'Taxonomy' (case-insensitive)."
        ),
    )
    parser.add_argument(
        "--aura-input-map",
        "--l1-input-map",
        dest="aura_input_map",
        type=Path,
        default=None,
        help=(
            "JSON manifest containing the standardized Aura workbook ranges. "
            "Taxonomy, company names, revenue, and Function + FTE/$B are shared "
            "by Aura L0 and Aura L1; L1 additionally uses the three benchmark columns. "
            "--l1-input-map is retained as a legacy alias."
        ),
    )
    parser.add_argument(
        "--company-sort-order",
        "--l1-sort-order",
        dest="company_sort_order",
        choices=[
            "revenue_ascending",
            "revenue_descending",
            "sheet_order",
            "segment",
        ],
        default="revenue_descending",
        help=(
            "Universal company-bar order for Aura L0, Aura L1, and Geo Cost. "
            "The leftmost client remains fixed; only peers are reordered. "
            "--l1-sort-order is retained as a legacy alias."
        ),
    )
    parser.add_argument(
        "--company-segment-sort-direction",
        "--l1-segment-sort-direction",
        dest="company_segment_sort_direction",
        choices=["revenue_ascending", "revenue_descending"],
        default="revenue_descending",
        help=(
            "Within-segment peer order for universal segment sorting. The client "
            "remains first and segment groups retain their detected order."
        ),
    )
    parser.add_argument(
        "--company-sort-manifest",
        type=Path,
        default=None,
        help=(
            "UI-generated company/revenue/segment profile used to apply one bar "
            "order consistently across Aura L0, Aura L1, and Geo Cost."
        ),
    )

    parser.add_argument("--skip-summary", action="store_true")
    parser.add_argument("--summary-template", type=Path, default=DEFAULT_SUMMARY_TEMPLATE_PPT)

    parser.add_argument("--skip-l0", action="store_true")
    parser.add_argument("--l0-template", type=Path, default=DEFAULT_L0_TEMPLATE_PPT)
    parser.add_argument("--only-l0", default=None)
    parser.add_argument("--l0-max-slides", type=int, default=None)
    parser.add_argument("--skip-l0-quartiles", action="store_true")
    parser.add_argument("--l0-post-chart-wait", type=float, default=1.5)

    parser.add_argument(
        "--skip-l1",
        action="store_true",
        help="Omit the Aura L1 section while retaining any other selected sections.",
    )
    parser.add_argument("--max-slides", type=int, default=None)
    parser.add_argument("--only-function", default=None)
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--dump-shapes", action="store_true")
    parser.add_argument("--keep-temp", action="store_true")

    parser.add_argument("--skip-quartiles", action="store_true")
    parser.add_argument("--skip-agenda", action="store_true")

    parser.add_argument("--skip-geocost", action="store_true")
    parser.add_argument(
        "--geocost-scope",
        choices=["auto", "all", "global", "us"],
        default="auto",
        help="auto/all builds every detected scope; global or us limits the appended section.",
    )
    parser.add_argument("--global-geocost-sheet", default=None)
    parser.add_argument("--us-geocost-sheet", default=None)
    parser.add_argument(
        "--geocost-input-map",
        type=Path,
        default=None,
        help=(
            "Standardized UI manifest containing the Overall Geo Cost table "
            "range and repeated >>> function-marker row."
        ),
    )
    parser.add_argument(
        "--global-geocost-template",
        type=Path,
        default=DEFAULT_GEOCOST_GLOBAL_TEMPLATE_PPT,
    )
    parser.add_argument(
        "--us-geocost-template",
        type=Path,
        default=DEFAULT_GEOCOST_US_TEMPLATE_PPT,
    )
    parser.add_argument("--global-geocost-chart-name", default="")
    parser.add_argument("--us-geocost-chart-name", default="")
    parser.add_argument("--geocost-function", default="ALL")
    parser.add_argument("--geocost-max-functions", type=int, default=None)
    parser.add_argument(
        "--geocost-max-companies",
        type=int,
        default=0,
        help=(
            "Optional Geo Cost company cap for debugging. Use 0 (the default) "
            "to include every company detected in each function block."
        ),
    )
    parser.add_argument(
        "--geocost-label-threshold",
        type=float,
        default=DEFAULT_LABEL_THRESHOLD,
    )
    parser.add_argument("--geocost-post-chart-wait", type=float, default=2.0)

    parser.add_argument("--hidden", action="store_true")
    parser.add_argument("--no-window", action="store_true")
    parser.add_argument("--quit-office", action="store_true")
    parser.add_argument("--kill-office", action="store_true")
    parser.add_argument("--com-check-only", action="store_true")

    parser.add_argument(
        "--hang-dump-seconds",
        type=int,
        default=0,
        help="Optional repeating Python stack dump interval; 0 disables watchdog dumps.",
    )
    parser.add_argument("--thinkcell-retries", type=int, default=3)
    parser.add_argument("--post-chart-wait", type=float, default=1.2)
    return parser.parse_args(argv)


def _segment_assignments_from_input_map(path: Path | None) -> list[dict]:
    """Read optional UI-detected segment assignments from the Aura manifest."""
    if path is None:
        return []
    resolved = Path(path).expanduser().resolve()
    if not resolved.is_file():
        return []
    try:
        payload = json.loads(resolved.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Aura L1 input map is not readable JSON: {resolved}") from exc
    assignments = payload.get("assignments")
    if isinstance(assignments, list):
        return [dict(item) for item in assignments if isinstance(item, dict)]

    segments = payload.get("segments")
    if not isinstance(segments, dict) or not bool(segments.get("detected")):
        return []
    assignments = segments.get("assignments")
    if not isinstance(assignments, list):
        return []
    return [dict(item) for item in assignments if isinstance(item, dict)]


def config_from_args(args: argparse.Namespace) -> BuildConfig:
    if args.geocost_max_companies < 0:
        raise ValueError(
            "--geocost-max-companies must be 0 (all companies) or a positive integer."
        )
    excel_path = args.excel.resolve() if args.excel is not None else DEFAULT_COM_CHECK_EXCEL.resolve()
    aura_input_map_path = (
        resolve_path(args.aura_input_map, WORKFLOW_DIR)
        if args.aura_input_map is not None
        else None
    )
    company_sort_manifest_path = (
        resolve_path(args.company_sort_manifest, WORKFLOW_DIR)
        if args.company_sort_manifest is not None
        else None
    )
    segment_assignments = _segment_assignments_from_input_map(
        company_sort_manifest_path or aura_input_map_path
    )
    if args.company_sort_order == "segment" and not segment_assignments:
        raise ValueError(
            "--company-sort-order segment requires a company sort manifest or "
            "Aura input manifest with detected segment assignments."
        )
    return BuildConfig(
        excel_path=excel_path,
        template_ppt=args.template.resolve(),
        output_dir=args.output_dir.resolve(),
        output_ppt_name=args.output_name,
        sheet_name=args.sheet_name,
        taxonomy_sheet_name=args.taxonomy_sheet,
        aura_input_map_path=aura_input_map_path,
        l1_input_map_path=aura_input_map_path,
        company_sort_order=args.company_sort_order,
        company_segment_sort_direction=args.company_segment_sort_direction,
        company_sort_manifest_path=company_sort_manifest_path,
        company_sort_assignments=segment_assignments,
        l1_sort_order=args.company_sort_order,
        l1_segment_sort_direction=args.company_segment_sort_direction,
        l1_segment_assignments=segment_assignments,
        max_slides=args.max_slides,
        only_function=args.only_function,
        update_quartiles=not args.skip_quartiles,
        update_agenda_text=not args.skip_agenda,
        ppt_visible=not args.hidden,
        excel_visible=not args.hidden,
        with_window=not args.no_window,
        quit_office_apps=args.quit_office,
        keep_temp=args.keep_temp,
        debug=args.debug,
        dump_shapes=args.dump_shapes,
        kill_office_first=args.kill_office,
        com_check_only=args.com_check_only,
        hang_dump_seconds=args.hang_dump_seconds,
        thinkcell_retries=args.thinkcell_retries,
        post_chart_update_wait_seconds=args.post_chart_wait,
        carryover_mode="defer",
        include_l0=not args.skip_l0,
        l0_template_ppt=resolve_path(args.l0_template, L0_DIR),
        l0_only_group=args.only_l0,
        l0_max_slides=args.l0_max_slides,
        l0_update_quartiles=not args.skip_l0_quartiles,
        l0_post_chart_wait_seconds=max(0.3, float(args.l0_post_chart_wait)),
        include_geocost=not args.skip_geocost,
        geocost_scope=args.geocost_scope,
        geocost_input_map_path=(
            resolve_path(args.geocost_input_map, WORKFLOW_DIR)
            if args.geocost_input_map is not None
            else None
        ),
        geocost_global_sheet_name=args.global_geocost_sheet,
        geocost_us_sheet_name=args.us_geocost_sheet,
        geocost_global_template_ppt=resolve_path(args.global_geocost_template, GEOCOST_DIR),
        geocost_us_template_ppt=resolve_path(args.us_geocost_template, GEOCOST_DIR),
        geocost_requested_function=args.geocost_function,
        geocost_max_functions=args.geocost_max_functions,
        geocost_max_companies=args.geocost_max_companies,
        geocost_label_threshold=args.geocost_label_threshold,
        geocost_global_chart_name=args.global_geocost_chart_name,
        geocost_us_chart_name=args.us_geocost_chart_name,
        geocost_post_chart_wait_seconds=args.geocost_post_chart_wait,
    )


def resolve_summary_template_path(args: argparse.Namespace) -> Path:
    return resolve_path(args.summary_template, SUMMARY_DIR)


def _requested_geocost_scopes(config: BuildConfig) -> set[str]:
    if config.geocost_scope in {"auto", "all"}:
        return {"global", "us"}
    return {config.geocost_scope}


def _template_for_geocost_scope(config: BuildConfig, scope: str) -> Path:
    path = (
        config.geocost_global_template_ppt
        if scope == "global"
        else config.geocost_us_template_ppt
    )
    if path is None or not path.exists():
        if scope == "us":
            raise FileNotFoundError(
                f"US Geo Cost sheet detected, but the US PowerPoint template is missing: {path}\n"
                "Add Geo_Cost_US_Template.pptx to the Global Geo Cost folder, pass "
                "--us-geocost-template, or run with --geocost-scope global. "
                "The Global template is not reused because its entity axis is different."
            )
        raise FileNotFoundError(f"Global Geo Cost PowerPoint template not found: {path}")
    return path


def build_integrated_workflow(
    config: BuildConfig,
    args: argparse.Namespace,
    logger,
) -> Path:
    """Build ordered sections: Summary -> Aura L0 -> Aura L1 -> Geo Cost."""
    # Delay Windows-only PowerPoint imports so CLI help and workbook inventory
    # remain available outside the Office runtime. L1-specific modules are not
    # imported when a custom section selection excludes Aura L1.
    if not args.skip_l1:
        from slide_collator import build_deck, prepare_l1_source
        from quartile_line_handler import patch_l1_quartiles_in_deck

    if config.include_l0:
        from l0_builder import build_l0_section_deck
        from l0_manifest import DEFAULT_PATCH_MANIFEST_NAME, write_l0_patch_manifest
        from l0_native import patch_l0_slides_in_deck

    if not 0 <= config.geocost_label_threshold <= 100:
        raise ValueError("--geocost-label-threshold must be between 0 and 100.")

    # Record PowerPoint before any Office automation begins. PowerPoint is a
    # single-process COM server, so the final merge must first close only the
    # process started by this workflow (never a pre-existing user session).
    powerpoint_pids_at_start = running_powerpoint_process_ids()
    logger.info(
        "PowerPoint process(es) present before workflow: %s",
        sorted(powerpoint_pids_at_start) or "none",
    )

    if config.include_l0:
        if config.l0_template_ppt is None:
            raise ValueError("Aura L0 generation is enabled but no L0 template was configured.")
        require_file(config.l0_template_ppt, "Aura L0 PowerPoint template")
        if not config.l0_update_quartiles:
            logger.warning(
                "Aura L0 native color/legend/quartile finalization is disabled by "
                "--skip-l0-quartiles. The section shell will still be generated."
            )
    else:
        logger.info("Aura L0 generation disabled by --skip-l0.")

    if args.skip_l1:
        logger.info("Aura L1 generation disabled by --skip-l1.")

    # Detect and validate Geo Cost inputs before slower Office automation.
    geocost_matches = []
    if config.include_geocost:
        detection = detect_geocost_sheets(
            config.excel_path,
            global_sheet_override=config.geocost_global_sheet_name,
            us_sheet_override=config.geocost_us_sheet_name,
            logger=logger,
            input_map_path=config.geocost_input_map_path,
        )
        requested_scopes = _requested_geocost_scopes(config)
        geocost_matches = detection.selected(requested_scopes)
        if config.geocost_scope in {"global", "us"} and not geocost_matches:
            raise ValueError(
                f"--geocost-scope {config.geocost_scope} was requested, "
                "but no matching worksheet was detected."
            )
        if not geocost_matches:
            logger.info("No Geo Cost tab detected; final deck will contain Aura sections only.")
        for match in geocost_matches:
            _template_for_geocost_scope(config, match.scope)
    else:
        logger.info("Geo Cost generation disabled by --skip-geocost.")

    final_output_ppt = get_output_path(config)
    working_dir = config.output_dir / "_workflow_temp" / datetime.now().strftime("%Y%m%d_%H%M%S")
    working_dir.mkdir(parents=True, exist_ok=True)
    section_decks: list[Path] = []
    section_slide_cursor = 0
    l0_result = None
    l0_final_slide_indices: list[int] = []
    l1_final_slide_indices: list[int] = []
    l1_records: list[dict] = []
    workflow_success = False

    def register_section(path: Path, label: str) -> list[int]:
        nonlocal section_slide_cursor
        require_file(path, f"{label} deck")
        count = presentation_slide_count(path)
        if count <= 0:
            raise ValueError(f"{label} deck contains no slides: {path}")
        start_slide = section_slide_cursor + 1
        end_slide = section_slide_cursor + count
        section_decks.append(path)
        section_slide_cursor = end_slide
        logger.info(
            "Registered final-deck section %s: slides %s-%s from %s",
            label,
            start_slide,
            end_slide,
            path.name,
        )
        return list(range(start_slide, end_slide + 1))

    try:
        if not args.skip_summary:
            summary_ppt = working_dir / "00_Aura_Summary.pptx"
            with stage(logger, "Build Aura Summary slide"):
                build_aura_summary_slide(
                    excel_path=config.excel_path,
                    template_ppt=resolve_summary_template_path(args),
                    output_ppt=summary_ppt,
                    logger=logger,
                    sheet_name=args.sheet_name,
                )
            register_section(summary_ppt, "Aura Summary")

        if config.include_l0:
            l0_ppt = working_dir / "01_Aura_L0_Only.pptx"
            # Aura L0 and Aura L1 now consume the same standardized Aura input
            # manifest. L0 uses the shared Taxonomy/company/revenue/FTE ranges;
            # its own dedicated-L0 benchmark detection and peer fallback remain unchanged.
            with stage(logger, "Build Aura L0 section shell"):
                l0_result = build_l0_section_deck(
                    config,
                    logger,
                    output_ppt=l0_ppt,
                    template_ppt=config.l0_template_ppt,
                    only_l0=config.l0_only_group,
                    max_slides=config.l0_max_slides,
                    update_quartiles=config.l0_update_quartiles,
                    post_chart_wait_seconds=config.l0_post_chart_wait_seconds,
                    finalize_native_patch=False,
                )
            l0_final_slide_indices = register_section(l0_result.deck_path, "Aura L0")
            if len(l0_final_slide_indices) != l0_result.slide_count:
                raise RuntimeError(
                    "Aura L0 section slide mapping mismatch: "
                    f"registered={len(l0_final_slide_indices)}, records={l0_result.slide_count}."
                )
            logger.info(
                "Aura L0 native finalization will target consolidated slide(s): %s",
                l0_final_slide_indices,
            )
            l0_manifest_path = write_l0_patch_manifest(
                working_dir / DEFAULT_PATCH_MANIFEST_NAME,
                l0_result.slide_records,
                section_deck_name=l0_result.deck_path.name,
                native_patch_pending=l0_result.native_patch_pending,
            )
            logger.info("Saved Aura L0 recovery patch metadata: %s", l0_manifest_path)

        if not args.skip_l1:
            original_output_dir = config.output_dir
            original_output_name = config.output_ppt_name
            try:
                config.output_dir = working_dir
                config.output_ppt_name = "02_Aura_L1_Only.pptx"
                with stage(logger, "Prepare Aura L1 source and quartile plan"):
                    l1_prepared_source = prepare_l1_source(config, logger)
                    l1_records = l1_prepared_source[2]
                with stage(logger, "Build Aura L1 slides"):
                    l1_ppt = build_deck(
                        config,
                        logger,
                        prepared_source=l1_prepared_source,
                        finalize_quartiles=False,
                    )
            finally:
                config.output_dir = original_output_dir
                config.output_ppt_name = original_output_name
            l1_final_slide_indices = register_section(l1_ppt, "Aura L1")
            if len(l1_final_slide_indices) != len(l1_records):
                raise RuntimeError(
                    "Aura L1 section slide mapping mismatch: "
                    f"registered={len(l1_final_slide_indices)}, records={len(l1_records)}."
                )
            logger.info(
                "Aura L1 active-stream quartile finalization will target consolidated "
                "slide(s): %s",
                l1_final_slide_indices,
            )

        for match in geocost_matches:
            template = _template_for_geocost_scope(config, match.scope)
            chart_name = (
                config.geocost_global_chart_name
                if match.scope == "global"
                else config.geocost_us_chart_name
            )
            section_order = 3 if match.scope == "global" else 4
            geocost_ppt = working_dir / (
                f"{section_order:02d}_GeoCost_{match.scope.upper()}.pptx"
            )
            with stage(logger, f"Build {match.scope.upper()} Geo Cost slides"):
                build_geocost_deck(
                    excel_path=config.excel_path,
                    template_path=template,
                    output_path=geocost_ppt,
                    requested_function=config.geocost_requested_function,
                    max_functions=config.geocost_max_functions,
                    max_companies=config.geocost_max_companies,
                    chart_name=chart_name,
                    label_threshold=config.geocost_label_threshold,
                    keep_temp=config.keep_temp,
                    quit_office=False,
                    post_chart_wait=config.geocost_post_chart_wait_seconds,
                    logger=logger,
                    scope=match.scope,
                    sheet_name=match.sheet_name,
                    source_layout=match.source_layout,
                    company_sort_order=(
                        config.company_sort_order or config.l1_sort_order
                    ),
                    company_segment_sort_direction=(
                        config.company_segment_sort_direction
                        or config.l1_segment_sort_direction
                    ),
                    company_sort_assignments=(
                        config.company_sort_assignments
                        or config.l1_segment_assignments
                        or []
                    ),
                )
            register_section(geocost_ppt, f"{match.scope.upper()} Geo Cost")

        if not section_decks:
            raise ValueError(
                "No slide sections were generated. Select at least one section and "
                "ensure its workbook source is available."
            )

        logger.info(
            "All selected section decks are complete. Final merge will run in a separate Python/PowerPoint process."
        )
        result = combine_presentation_files(
            section_decks,
            final_output_ppt,
            logger,
            visible=config.ppt_visible,
            with_window=config.with_window,
            quit_office=config.quit_office_apps,
            stage_name="Combine selected Summary, Aura L0, Aura L1, and Geo Cost sections",
            preexisting_powerpoint_pids=powerpoint_pids_at_start,
        )

        # Native carry-over data is written only after the consolidated deck has
        # been saved and closed. No PowerPoint reopen/save is allowed afterward.
        # L1 is finalized first; L0 remains the final package mutation whenever
        # it is present, preserving the accepted L0 closed-deck contract.
        if not args.skip_l1 and config.update_quartiles and l1_records:
            with stage(logger, "Finalize Aura L1 quartiles in the closed consolidated deck"):
                patch_l1_quartiles_in_deck(
                    output_ppt=result,
                    slide_records=l1_records,
                    slide_indices=l1_final_slide_indices,
                    logger=logger,
                )
            logger.warning(
                "Aura L1 quartile carry-over data is verified on consolidated slide(s) %s. "
                "For each L1 slide, choose 'Manually carry over changes to think-cell', "
                "click Finish, and save the deck.",
                l1_final_slide_indices,
            )

        if l0_result is not None and l0_result.native_patch_pending:
            with stage(logger, "Finalize Aura L0 state in the closed consolidated deck"):
                patch_l0_slides_in_deck(
                    output_ppt=result,
                    slide_records=l0_result.slide_records,
                    slide_indices=l0_final_slide_indices,
                    logger=logger,
                )
            logger.warning(
                "Aura L0 finalization is complete on consolidated slide(s) %s. For each L0 slide, "
                "choose 'Manually carry over changes to think-cell', click Finish, and save the deck.",
                l0_final_slide_indices,
            )

        workflow_success = True
        return result
    finally:
        if config.keep_temp or not workflow_success:
            logger.warning(
                "Workflow temp folder kept%s: %s",
                " because the workflow did not complete" if not workflow_success else "",
                working_dir,
            )
        else:
            shutil.rmtree(working_dir, ignore_errors=True)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    if not args.com_check_only:
        args = resolve_run_inputs(args, WORKFLOW_DIR)
    elif args.excel is None:
        args.excel = DEFAULT_COM_CHECK_EXCEL

    config = config_from_args(args)
    logger = setup_logging(config)

    if config.hang_dump_seconds and config.hang_dump_seconds > 0:
        faulthandler.enable()
        faulthandler.dump_traceback_later(config.hang_dump_seconds, repeat=True)

    try:
        logger.info("Workflow: Integrated Aura Summary + L0 + L1 + Geo Cost")
        logger.info("Aura root: %s", AURA_ROOT)
        logger.info("Aura L0 dir: %s", L0_DIR)
        logger.info("Aura L1 dir: %s", WORKFLOW_DIR)
        logger.info("Aura Summary dir: %s", SUMMARY_DIR)
        logger.info("Geo Cost dir: %s", GEOCOST_DIR)
        logger.info("Excel path: %s", config.excel_path)
        if config.aura_input_map_path is not None:
            logger.info("Standardized Aura input map: %s", config.aura_input_map_path)
        effective_sort_order = config.company_sort_order or config.l1_sort_order
        effective_segment_direction = (
            config.company_segment_sort_direction
            or config.l1_segment_sort_direction
        )
        logger.info(
            "Universal company order: client fixed first; mode=%s",
            effective_sort_order,
        )
        if effective_sort_order == "segment":
            logger.info(
                "Universal segment order: detected segment order; within-segment mode=%s; assignments=%s",
                effective_segment_direction,
                len(config.company_sort_assignments or config.l1_segment_assignments or []),
            )
        if not args.skip_l1:
            logger.info(
                "Aura L1 taxonomy worksheet: %s",
                config.taxonomy_sheet_name or "auto-detect by tab name containing 'Taxonomy'",
            )
        logger.info("Aura L1 template PPT: %s", config.template_ppt)
        if config.include_l0:
            logger.info("Aura L0 template PPT: %s", config.l0_template_ppt)
        logger.info("Aura Summary template PPT: %s", resolve_summary_template_path(args))
        if config.geocost_scope in {"auto", "all", "global"}:
            logger.info("Global Geo Cost template PPT: %s", config.geocost_global_template_ppt)
        if config.geocost_scope in {"auto", "all", "us"}:
            logger.info("US Geo Cost template PPT: %s", config.geocost_us_template_ppt)
        logger.info("Final output dir: %s", config.output_dir)

        if config.com_check_only:
            from slide_collator import run_com_check

            pythoncom.CoInitialize()
            try:
                if config.kill_office_first:
                    kill_office_processes(logger)
                run_com_check(config, logger)
            finally:
                pythoncom.CoUninitialize()
            return

        build_integrated_workflow(config, args, logger)
    finally:
        try:
            faulthandler.cancel_dump_traceback_later()
        except Exception:
            pass


if __name__ == "__main__":
    main()
