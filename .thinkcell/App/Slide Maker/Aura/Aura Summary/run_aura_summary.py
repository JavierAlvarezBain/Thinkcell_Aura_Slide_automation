from __future__ import annotations

import argparse
import sys
from pathlib import Path

SUMMARY_DIR = Path(__file__).parent.resolve()
AURA_ROOT = SUMMARY_DIR.parent.resolve()
AURA_L1_DIR = AURA_ROOT / "Aura L1"

# The Summary builder reuses the Aura L1 Excel detector, so both workflow
# folders and the universal Aura root must be importable during direct runs.
for path in [str(AURA_ROOT), str(AURA_L1_DIR), str(SUMMARY_DIR)]:
    while path in sys.path:
        sys.path.remove(path)
sys.path.insert(0, str(AURA_ROOT))
sys.path.insert(0, str(AURA_L1_DIR))
sys.path.insert(0, str(SUMMARY_DIR))

from shared import BuildConfig, setup_logging  # noqa: E402
from input import resolve_run_inputs  # noqa: E402
from aura_summary_builder import build_aura_summary_slide  # noqa: E402


DEFAULT_TEMPLATE_PPT = SUMMARY_DIR / "Table Test.pptx"
DEFAULT_OUTPUT_DIR = SUMMARY_DIR / "output"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate Aura Summary Bain table slide.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("--excel", type=Path, default=None)
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE_PPT)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--output-name", default="Aura_Summary_Slide.pptx")
    parser.add_argument("--interactive-input", action="store_true")
    parser.add_argument("--no-input-prompt", action="store_true")
    parser.add_argument("--sheet-name", default=None)
    parser.add_argument("--debug", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    args = parse_args(argv)
    args = resolve_run_inputs(args, SUMMARY_DIR)

    config = BuildConfig(
        excel_path=args.excel.resolve(),
        template_ppt=args.template.resolve(),
        output_dir=args.output_dir.resolve(),
        output_ppt_name=args.output_name,
        sheet_name=args.sheet_name,
        debug=args.debug,
    )
    logger = setup_logging(config)

    logger.info("Workflow: Aura Summary")
    logger.info("Aura root: %s", AURA_ROOT)
    logger.info("Workflow dir: %s", SUMMARY_DIR)
    logger.info("Excel path: %s", config.excel_path)
    logger.info("Template PPT: %s", config.template_ppt)
    logger.info("Output PPT: %s", config.output_ppt)

    build_aura_summary_slide(
        excel_path=config.excel_path,
        template_ppt=config.template_ppt,
        output_ppt=config.output_ppt,
        logger=logger,
        sheet_name=args.sheet_name,
    )


if __name__ == "__main__":
    main()
