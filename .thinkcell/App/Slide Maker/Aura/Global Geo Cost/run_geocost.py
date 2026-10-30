from __future__ import annotations

import argparse
import shutil
import sys
from datetime import datetime
from pathlib import Path

GEOCOST_DIR = Path(__file__).parent.resolve()
AURA_ROOT = GEOCOST_DIR.parent.resolve()
for path in [str(AURA_ROOT), str(GEOCOST_DIR)]:
    while path in sys.path:
        sys.path.remove(path)
sys.path.insert(0, str(AURA_ROOT))
sys.path.insert(0, str(GEOCOST_DIR))

from build_global_geocost import build_geocost_deck, prepare_data_workbooks  # noqa: E402
from geocost_common import BUILD_ID, DEFAULT_LABEL_THRESHOLD  # noqa: E402
from geocost_source import detect_geocost_sheets  # noqa: E402
from input import resolve_run_inputs, resolve_path  # noqa: E402
from shared import (  # noqa: E402
    BuildConfig,
    DEFAULT_GEOCOST_GLOBAL_TEMPLATE_PPT,
    DEFAULT_GEOCOST_US_TEMPLATE_PPT,
    combine_presentation_files,
    get_output_path,
    load_company_sort_manifest,
    require_file,
    setup_logging,
    stage,
)

DEFAULT_OUTPUT_DIR = GEOCOST_DIR / "output"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build detected Global and/or US Geo Cost think-cell slides.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--excel", type=Path, default=None)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--output-name", default="GeoCost_Generated_Deck.pptx")
    parser.add_argument("--interactive-input", action="store_true")
    parser.add_argument("--no-input-prompt", action="store_true")

    parser.add_argument("--scope", choices=["auto", "all", "global", "us"], default="auto")
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
        "--global-template",
        "--template",
        dest="global_template",
        type=Path,
        default=DEFAULT_GEOCOST_GLOBAL_TEMPLATE_PPT,
        help="Global Geo Cost PowerPoint template; --template is a legacy alias.",
    )
    parser.add_argument(
        "--us-template",
        type=Path,
        default=DEFAULT_GEOCOST_US_TEMPLATE_PPT,
    )
    parser.add_argument(
        "--global-chart-name",
        "--chart-name",
        dest="global_chart_name",
        default="",
        help="Global think-cell element name; --chart-name is a legacy alias.",
    )
    parser.add_argument("--us-chart-name", default="")

    parser.add_argument(
        "--company-sort-order",
        choices=["revenue_ascending", "revenue_descending", "sheet_order", "segment"],
        default="sheet_order",
        help=(
            "Universal company-bar order. Revenue and segment modes use the "
            "UI-generated --company-sort-manifest; the first company remains fixed."
        ),
    )
    parser.add_argument(
        "--company-segment-sort-direction",
        choices=["revenue_ascending", "revenue_descending"],
        default="revenue_descending",
    )
    parser.add_argument(
        "--company-sort-manifest",
        type=Path,
        default=None,
        help="UI-generated company/revenue/segment profile shared across slide types.",
    )
    parser.add_argument("--function", default="ALL")
    parser.add_argument("--max-functions", type=int, default=None)
    parser.add_argument(
        "--max-companies",
        type=int,
        default=0,
        help=(
            "Optional company cap for debugging. Use 0 (the default) to include "
            "every company detected in each Geo Cost function block."
        ),
    )
    parser.add_argument("--label-threshold", type=float, default=DEFAULT_LABEL_THRESHOLD)
    parser.add_argument("--post-chart-wait", type=float, default=2.0)
    parser.add_argument("--chart-data-only", action="store_true")
    parser.add_argument("--keep-temp", action="store_true")
    parser.add_argument("--quit-office", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--version", action="version", version=f"%(prog)s {BUILD_ID}")
    return parser.parse_args(argv)


def _selected_matches(args: argparse.Namespace, detection):
    requested = {"global", "us"} if args.scope in {"auto", "all"} else {args.scope}
    matches = detection.selected(requested)
    found_scopes = {match.scope for match in matches}
    if args.scope in {"global", "us"} and args.scope not in found_scopes:
        raise ValueError(
            f"--scope {args.scope} was requested, but no matching Geo Cost worksheet was detected."
        )
    if not matches:
        raise ValueError(
            "No structurally valid Geo Cost worksheet was detected. "
            "Supported names include Geo Cost, Geo-Cost, Global Geo Cost, and US Geo Cost."
        )
    return matches


def _template_for_scope(args: argparse.Namespace, scope: str) -> Path:
    value = args.global_template if scope == "global" else args.us_template
    path = resolve_path(value, GEOCOST_DIR)
    if not path.exists():
        if scope == "us":
            raise FileNotFoundError(
                f"US Geo Cost sheet detected, but the US PowerPoint template is missing: {path}\n"
                "Add Geo_Cost_US_Template.pptx to the Global Geo Cost folder or pass --us-template. "
                "The Global template is intentionally not reused because its entity axis is different."
            )
        raise FileNotFoundError(f"Global Geo Cost PowerPoint template not found: {path}")
    return path


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    args = resolve_run_inputs(args, GEOCOST_DIR, resolve_template=False)
    if not 0 <= args.label_threshold <= 100:
        raise ValueError("--label-threshold must be between 0 and 100.")
    if args.max_companies < 0:
        raise ValueError("--max-companies must be 0 (all companies) or a positive integer.")

    sort_manifest_path = (
        resolve_path(args.company_sort_manifest, GEOCOST_DIR)
        if args.company_sort_manifest is not None
        else None
    )
    sort_payload = load_company_sort_manifest(sort_manifest_path)
    sort_assignments = [
        dict(item)
        for item in (sort_payload.get("assignments") or [])
        if isinstance(item, dict)
    ]
    if args.company_sort_order == "segment" and not sort_assignments:
        raise ValueError(
            "--company-sort-order segment requires --company-sort-manifest with "
            "detected company-to-segment assignments."
        )

    config = BuildConfig(
        excel_path=args.excel.resolve(),
        template_ppt=resolve_path(args.global_template, GEOCOST_DIR),
        output_dir=args.output_dir.resolve(),
        output_ppt_name=args.output_name,
        debug=args.debug,
        keep_temp=args.keep_temp,
        quit_office_apps=args.quit_office,
        company_sort_order=args.company_sort_order,
        company_segment_sort_direction=args.company_segment_sort_direction,
        company_sort_manifest_path=sort_manifest_path,
        company_sort_assignments=sort_assignments,
        l1_sort_order=args.company_sort_order,
        l1_segment_sort_direction=args.company_segment_sort_direction,
        l1_segment_assignments=sort_assignments,
        geocost_input_map_path=(
            resolve_path(args.geocost_input_map, GEOCOST_DIR)
            if args.geocost_input_map is not None
            else None
        ),
    )
    logger = setup_logging(config)
    logger.info("Workflow: Geo Cost")
    logger.info("Build: %s", BUILD_ID)
    logger.info("Excel path: %s", config.excel_path)

    detection = detect_geocost_sheets(
        config.excel_path,
        global_sheet_override=args.global_geocost_sheet,
        us_sheet_override=args.us_geocost_sheet,
        logger=logger,
        input_map_path=config.geocost_input_map_path,
    )
    matches = _selected_matches(args, detection)

    if args.chart_data_only:
        audit_root = config.output_dir / "chart_data_audit"
        total = 0
        for match in matches:
            template = _template_for_scope(args, match.scope)
            chart_name = args.global_chart_name if match.scope == "global" else args.us_chart_name
            with stage(logger, f"Audit {match.scope.upper()} Geo Cost ChartData"):
                results = prepare_data_workbooks(
                    excel_path=config.excel_path,
                    template_path=template,
                    output_dir=audit_root / match.scope,
                    requested_function=args.function,
                    max_functions=args.max_functions,
                    max_companies=args.max_companies,
                    label_threshold=args.label_threshold,
                    logger=logger,
                    scope=match.scope,
                    sheet_name=match.sheet_name,
                    source_layout=match.source_layout,
                    chart_name=chart_name,
                    company_sort_order=args.company_sort_order,
                    company_segment_sort_direction=args.company_segment_sort_direction,
                    company_sort_assignments=sort_assignments,
                )
                total += len(results)
        logger.info("Generated %s audited Geo Cost ChartData workbook(s) in %s", total, audit_root)
        return

    final_output = get_output_path(config)
    if len(matches) == 1:
        match = matches[0]
        template = _template_for_scope(args, match.scope)
        chart_name = args.global_chart_name if match.scope == "global" else args.us_chart_name
        build_geocost_deck(
            excel_path=config.excel_path,
            template_path=template,
            output_path=final_output,
            requested_function=args.function,
            max_functions=args.max_functions,
            max_companies=args.max_companies,
            chart_name=chart_name,
            label_threshold=args.label_threshold,
            keep_temp=args.keep_temp,
            quit_office=args.quit_office,
            post_chart_wait=args.post_chart_wait,
            logger=logger,
            scope=match.scope,
            sheet_name=match.sheet_name,
            source_layout=match.source_layout,
            company_sort_order=args.company_sort_order,
            company_segment_sort_direction=args.company_segment_sort_direction,
            company_sort_assignments=sort_assignments,
        )
        return

    working_dir = config.output_dir / "_workflow_temp" / datetime.now().strftime("%Y%m%d_%H%M%S")
    working_dir.mkdir(parents=True, exist_ok=True)
    built_decks: list[Path] = []
    try:
        for order, match in enumerate(matches, start=1):
            template = _template_for_scope(args, match.scope)
            chart_name = args.global_chart_name if match.scope == "global" else args.us_chart_name
            section_path = working_dir / f"{order:02d}_GeoCost_{match.scope.upper()}.pptx"
            build_geocost_deck(
                excel_path=config.excel_path,
                template_path=template,
                output_path=section_path,
                requested_function=args.function,
                max_functions=args.max_functions,
                max_companies=args.max_companies,
                chart_name=chart_name,
                label_threshold=args.label_threshold,
                keep_temp=args.keep_temp,
                quit_office=False,
                post_chart_wait=args.post_chart_wait,
                logger=logger,
                scope=match.scope,
                sheet_name=match.sheet_name,
                source_layout=match.source_layout,
                company_sort_order=args.company_sort_order,
                company_segment_sort_direction=args.company_segment_sort_direction,
                company_sort_assignments=sort_assignments,
            )
            require_file(section_path, f"{match.scope.upper()} Geo Cost deck")
            built_decks.append(section_path)
        combine_presentation_files(
            built_decks,
            final_output,
            logger,
            visible=True,
            with_window=True,
            quit_office=args.quit_office,
            stage_name="Combine Global and US Geo Cost sections",
        )
    finally:
        if args.keep_temp:
            logger.warning("Geo Cost workflow temp folder kept: %s", working_dir)
        else:
            shutil.rmtree(working_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
