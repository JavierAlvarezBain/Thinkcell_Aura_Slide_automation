from __future__ import annotations

"""Standalone entry point for the modular Aura L0 workflow."""

import argparse
from pathlib import Path

from l0_config import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_OUTPUT_NAME,
    DEFAULT_TEMPLATE_PPT,
    L0_BUILD_ID,
    L0_DIR,
    PPT_SUFFIXES,
)
from input import resolve_path, resolve_run_inputs
from shared import BuildConfig, load_company_sort_manifest, setup_logging


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Build only Aura L0 think-cell slides from the shared Aura Summary table. "
            "The accepted L0 aggregation, quartile, ordering, color, legend, and native-patch "
            "logic stays entirely inside the Aura L0 folder."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--excel", type=Path, default=None)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_PPT)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--output-name", default=DEFAULT_OUTPUT_NAME)
    parser.add_argument("--interactive-input", action="store_true")
    parser.add_argument("--no-input-prompt", action="store_true")
    parser.add_argument("--sheet", "--sheet-name", dest="sheet_name", default=None)
    parser.add_argument(
        "--taxonomy-sheet",
        default=None,
        help="Optional Taxonomy worksheet override used by the standardized Aura input map.",
    )
    parser.add_argument(
        "--aura-input-map",
        type=Path,
        default=None,
        help=(
            "JSON manifest containing the shared Taxonomy, company, revenue, and "
            "Function + FTE/$B ranges used by Aura L0 and Aura L1."
        ),
    )
    parser.add_argument(
        "--company-sort-order",
        choices=["revenue_ascending", "revenue_descending", "sheet_order", "segment"],
        default="revenue_descending",
        help=(
            "Universal company-bar order. The leftmost client remains fixed; "
            "only peers are reordered."
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
        help="UI-generated company/revenue/segment sort profile shared across slide types.",
    )
    parser.add_argument("--only-l0", default=None)
    parser.add_argument("--max-slides", type=int, default=None)
    parser.add_argument("--skip-quartiles", action="store_true")
    parser.add_argument(
        "--patch-existing-quartiles",
        metavar="PPTX",
        help=(
            "Patch native L0 segment colors, legend order, and quartile lines in an existing "
            "L0-only deck without rebuilding charts."
        ),
    )
    parser.add_argument("--hidden", action="store_true")
    parser.add_argument("--no-window", action="store_true")
    parser.add_argument("--kill-office", action="store_true")
    parser.add_argument("--quit-office", action="store_true")
    parser.add_argument("--keep-temp", action="store_true")
    parser.add_argument("--debug", action="store_true")
    parser.add_argument("--dump-shapes", action="store_true")
    parser.add_argument("--thinkcell-retries", type=int, default=3)
    parser.add_argument("--post-chart-wait", type=float, default=1.5)
    return parser


def config_from_args(args: argparse.Namespace) -> BuildConfig:
    template = Path(args.template).resolve()
    aura_input_map_path = (
        resolve_path(args.aura_input_map, L0_DIR)
        if args.aura_input_map is not None
        else None
    )
    sort_manifest_path = (
        resolve_path(args.company_sort_manifest, L0_DIR)
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
    return BuildConfig(
        excel_path=Path(args.excel).resolve(),
        template_ppt=template,
        output_dir=Path(args.output_dir).resolve(),
        output_ppt_name=args.output_name,
        sheet_name=args.sheet_name,
        taxonomy_sheet_name=args.taxonomy_sheet,
        aura_input_map_path=aura_input_map_path,
        l1_input_map_path=aura_input_map_path,
        company_sort_order=args.company_sort_order,
        company_segment_sort_direction=args.company_segment_sort_direction,
        company_sort_manifest_path=sort_manifest_path,
        company_sort_assignments=sort_assignments,
        l1_sort_order=args.company_sort_order,
        l1_segment_sort_direction=args.company_segment_sort_direction,
        l1_segment_assignments=sort_assignments,
        update_quartiles=not args.skip_quartiles,
        update_agenda_text=True,
        excel_visible=not args.hidden,
        ppt_visible=not args.hidden,
        with_window=not args.no_window,
        quit_office_apps=bool(args.quit_office),
        keep_temp=bool(args.keep_temp),
        debug=bool(args.debug),
        dump_shapes=bool(args.dump_shapes),
        kill_office_first=bool(args.kill_office),
        thinkcell_retries=max(1, int(args.thinkcell_retries)),
        post_chart_update_wait_seconds=max(0.3, float(args.post_chart_wait)),
        carryover_mode="defer",
        include_l0=True,
        l0_template_ppt=template,
        l0_only_group=args.only_l0,
        l0_max_slides=args.max_slides,
        l0_update_quartiles=not args.skip_quartiles,
        l0_post_chart_wait_seconds=max(0.3, float(args.post_chart_wait)),
        include_geocost=False,
    )


def main(argv: list[str] | None = None) -> Path:
    print(f"\nAura L0 build: {L0_BUILD_ID}")
    parser = build_parser()
    args = parser.parse_args(argv)
    args = resolve_run_inputs(args, L0_DIR)

    if args.patch_existing_quartiles:
        args.patch_existing_quartiles = resolve_path(args.patch_existing_quartiles, L0_DIR)
        if not args.patch_existing_quartiles.exists():
            raise FileNotFoundError(
                f"Existing Aura L0 deck not found: {args.patch_existing_quartiles}"
            )
        if args.patch_existing_quartiles.suffix.lower() not in PPT_SUFFIXES:
            raise ValueError("--patch-existing-quartiles must point to a .pptx or .pptm file.")
        if args.skip_quartiles:
            raise ValueError("--skip-quartiles cannot be combined with --patch-existing-quartiles.")

    config = config_from_args(args)
    logger = setup_logging(config)
    logger.info(
        "Workflow: %s",
        "Aura L0 native-state recovery" if args.patch_existing_quartiles else "Aura L0 standalone",
    )
    logger.info("Aura L0 directory: %s", L0_DIR)
    logger.info("Excel path: %s", config.excel_path)
    logger.info("Aura L0 template: %s", config.template_ppt)
    logger.info("Output directory: %s", config.output_dir)

    # Import Windows/Office automation only after argparse has handled --help.
    from l0_builder import build_l0_section_deck, patch_existing_l0_deck_from_source

    if args.patch_existing_quartiles:
        return patch_existing_l0_deck_from_source(
            config,
            args.patch_existing_quartiles,
            logger,
            only_l0=args.only_l0,
            max_slides=args.max_slides,
        )

    result = build_l0_section_deck(
        config,
        logger,
        only_l0=args.only_l0,
        max_slides=args.max_slides,
        finalize_native_patch=True,
    )
    return result.deck_path


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nAura L0 run cancelled.")
        raise SystemExit(130)
