from __future__ import annotations

import argparse
from datetime import datetime
import logging
from pathlib import Path
import sys

TOOLS_DIR = Path(__file__).resolve().parent
APP_DIR = TOOLS_DIR.parent
AURA_DIR = APP_DIR / "Aura"
L0_DIR = AURA_DIR / "Aura L0"
for directory in (APP_DIR, AURA_DIR, L0_DIR):
    text = str(directory)
    if text not in sys.path:
        sys.path.insert(0, text)

from shared import (  # noqa: E402
    combine_presentation_files,
    presentation_slide_count,
    running_powerpoint_process_ids,
)
from l0_manifest import (  # noqa: E402
    DEFAULT_PATCH_MANIFEST_NAME,
    read_l0_patch_manifest,
)
from l0_native import patch_l0_slides_in_deck  # noqa: E402


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Merge already-generated Aura Summary, Aura L0, Aura L1, and Geo Cost "
            "section decks without rebuilding them. When an L0 section is present, "
            "its native colors, legend, and quartiles are finalized in the recovered "
            "combined deck after PowerPoint closes."
        )
    )
    parser.add_argument(
        "--workflow-temp",
        type=Path,
        required=True,
        help=(
            "A preserved output\\_workflow_temp\\YYYYMMDD_HHMMSS folder or a "
            "*_MERGE_RECOVERY_* folder created after a failed final merge."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Optional final PPTX path. Defaults to the Aura L1 output folder.",
    )
    parser.add_argument(
        "--l0-patch-metadata",
        type=Path,
        default=None,
        help=(
            "Optional Aura L0 patch metadata JSON. Defaults to the metadata sidecar "
            "inside --workflow-temp."
        ),
    )
    return parser.parse_args()


def build_logger() -> logging.Logger:
    logger = logging.getLogger("aura_merge_recovery")
    logger.handlers.clear()
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(
        logging.Formatter(
            fmt="%(asctime)s | %(levelname)-8s | %(message)s",
            datefmt="%H:%M:%S",
        )
    )
    logger.addHandler(handler)
    return logger


def _first_existing(workflow_temp: Path, names: list[str]) -> Path | None:
    for name in names:
        candidate = workflow_temp / name
        if candidate.is_file():
            return candidate
        # shared.py prefixes preserved recovery decks with their required merge
        # position (for example 02_01_Aura_L0_Only.pptx). Accept that form too.
        prefixed = sorted(workflow_temp.glob(f"*_{name}"))
        if prefixed:
            return prefixed[0]
    return None


def _matches_original_name(path: Path, original_name: str) -> bool:
    return path.name == original_name or path.name.endswith("_" + original_name)


def discover_section_decks(workflow_temp: Path) -> list[Path]:
    """Support the new L0-integrated names and the prior no-L0 recovery names."""
    slots = [
        ["00_Aura_Summary.pptx"],
        ["01_Aura_L0_Only.pptx"],
        ["02_Aura_L1_Only.pptx", "01_Aura_L1_Only.pptx"],
        ["03_GeoCost_GLOBAL.pptx", "02_GeoCost_GLOBAL.pptx"],
        ["04_GeoCost_US.pptx", "03_GeoCost_US.pptx"],
    ]
    result: list[Path] = []
    seen: set[Path] = set()
    for names in slots:
        path = _first_existing(workflow_temp, names)
        if path is not None and path not in seen:
            seen.add(path)
            result.append(path)
    return result


def main() -> None:
    args = parse_args()
    logger = build_logger()
    workflow_temp = args.workflow_temp.expanduser().resolve()
    if not workflow_temp.is_dir():
        raise FileNotFoundError(f"Workflow temp folder not found: {workflow_temp}")

    section_decks = discover_section_decks(workflow_temp)
    l1_deck = next(
        (path for path in section_decks if "Aura_L1_Only" in path.name),
        None,
    )
    if l1_deck is None:
        raise RuntimeError(
            "Could not find the completed Aura L1 section deck in the preserved workflow folder."
        )

    active_pids = running_powerpoint_process_ids()
    if active_pids:
        raise RuntimeError(
            "PowerPoint is still running. Save and close every PowerPoint window, then run "
            f"this recovery command again. Active POWERPNT.EXE PID(s): {sorted(active_pids)}"
        )

    if args.output is None:
        if workflow_temp.parent.name == "_workflow_temp":
            output_dir = workflow_temp.parent.parent
        elif "_MERGE_RECOVERY_" in workflow_temp.name:
            output_dir = workflow_temp.parent
        else:
            output_dir = workflow_temp.parent
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_ppt = output_dir / f"Generated_ThinkCell_Deck_RECOVERED_{timestamp}.pptx"
    else:
        output_ppt = args.output.expanduser().resolve()

    l0_deck = next(
        (
            path
            for path in section_decks
            if _matches_original_name(path, "01_Aura_L0_Only.pptx")
        ),
        None,
    )
    l0_records = None
    l0_native_patch_pending = False
    l0_final_slide_indices: list[int] = []
    if l0_deck is not None:
        metadata_path = (
            args.l0_patch_metadata.expanduser().resolve()
            if args.l0_patch_metadata is not None
            else workflow_temp / DEFAULT_PATCH_MANIFEST_NAME
        )
        payload = read_l0_patch_manifest(metadata_path)
        expected_l0_name = str(payload.get("section_deck_name") or "")
        if not _matches_original_name(l0_deck, expected_l0_name):
            raise ValueError(
                "Aura L0 patch metadata belongs to a different section deck: "
                f"metadata={expected_l0_name!r}, found={l0_deck.name!r}."
            )
        l0_records = payload["slide_records"]
        l0_native_patch_pending = bool(payload.get("native_patch_pending", True))
        l0_section_count = presentation_slide_count(l0_deck)
        if l0_section_count != len(l0_records):
            raise RuntimeError(
                "Aura L0 section and patch metadata disagree on slide count: "
                f"deck={l0_section_count}, metadata={len(l0_records)}."
            )
        slides_before_l0 = sum(
            presentation_slide_count(path)
            for path in section_decks[: section_decks.index(l0_deck)]
        )
        l0_final_slide_indices = list(
            range(slides_before_l0 + 1, slides_before_l0 + l0_section_count + 1)
        )
        if l0_native_patch_pending:
            logger.info(
                "Loaded Aura L0 patch metadata for recovered final slide(s): %s",
                l0_final_slide_indices,
            )
        else:
            logger.info(
                "Aura L0 patch metadata records native finalization as disabled; "
                "the recovered deck will preserve the unpatched section shell."
            )

    logger.info("Recovering completed sections in this order:")
    for deck in section_decks:
        logger.info("  %s", deck.name)

    result = combine_presentation_files(
        section_decks,
        output_ppt,
        logger,
        visible=True,
        with_window=True,
        preexisting_powerpoint_pids=set(),
        stage_name="Recover final deck from completed sections",
    )

    if l0_records is not None and l0_native_patch_pending:
        patch_l0_slides_in_deck(
            output_ppt=result,
            slide_records=l0_records,
            slide_indices=l0_final_slide_indices,
            logger=logger,
        )
        logger.warning(
            "Recovered Aura L0 slides %s now contain their native carry-over state. "
            "Use one manual think-cell carry-over per L0 slide, click Finish, and save.",
            l0_final_slide_indices,
        )

    logger.info("Recovery complete: %s", result)


if __name__ == "__main__":
    main()
