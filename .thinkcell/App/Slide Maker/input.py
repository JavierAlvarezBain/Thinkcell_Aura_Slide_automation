from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

from shared import (
    APP_ROOT,
    AURA_L0_DIR,
    AURA_L1_DIR,
    AURA_ROOT,
    AURA_SUMMARY_DIR,
    DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE_PPT,
    GEOCOST_DIR,
    geocost_scope_hint_from_sheet_name,
    looks_like_geocost_sheet_name,
)

# openpyxl reads .xlsx/.xlsm reliably. Older .xls and binary .xlsb files must be
# saved as .xlsx before running this automation.
EXCEL_SUFFIXES = {".xlsx", ".xlsm"}
PPT_SUFFIXES = {".pptx", ".pptm"}


def strip_outer_quotes(value: str) -> str:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1].strip()
    return value


def resolve_path(value: str | Path, base_dir: Path) -> Path:
    path = Path(strip_outer_quotes(str(value))).expanduser()
    if not path.is_absolute():
        search_roots = [
            base_dir.resolve(),
            AURA_ROOT.resolve(),
            APP_ROOT.resolve(),
            AURA_L0_DIR.resolve(),
            AURA_L1_DIR.resolve(),
            AURA_SUMMARY_DIR.resolve(),
            GEOCOST_DIR.resolve(),
        ]
        candidates = [(root / path).resolve() for root in search_roots]
        path = next((candidate for candidate in candidates if candidate.exists()), candidates[0])
    return path.resolve()


def _candidate_search_dirs(base_dir: Path) -> list[Path]:
    result: list[Path] = []
    for directory in (
        base_dir.resolve(),
        AURA_ROOT.resolve(),
        APP_ROOT.resolve(),
        AURA_L0_DIR.resolve(),
        AURA_L1_DIR.resolve(),
        AURA_SUMMARY_DIR.resolve(),
        GEOCOST_DIR.resolve(),
    ):
        if directory.is_dir() and directory not in result:
            result.append(directory)
    return result


def _is_internal_excel_template(path: Path) -> bool:
    key = "".join(ch for ch in path.stem.lower() if ch.isalnum())
    return key in {
        "geocostglobaltemplate",
        "geocostustemplate",
        "tcdata",
    } or key.startswith("tcdata")


def excel_candidates(base_dir: Path) -> list[Path]:
    candidates: list[Path] = []
    seen: set[Path] = set()
    for directory in _candidate_search_dirs(base_dir):
        for path in directory.iterdir():
            if not path.is_file() or path.name.startswith("~$"):
                continue
            if path.suffix.lower() not in EXCEL_SUFFIXES:
                continue
            if _is_internal_excel_template(path):
                continue
            resolved = path.resolve()
            if resolved not in seen:
                seen.add(resolved)
                candidates.append(resolved)
    return sorted(candidates, key=lambda item: item.stat().st_mtime, reverse=True)


def workbook_sheet_names(excel_path: Path) -> list[str]:
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    try:
        return list(workbook.sheetnames)
    finally:
        workbook.close()


def looks_like_aura_source_sheet_name(value: str) -> bool:
    """Return a lightweight name hint; structural validation remains authoritative."""
    key = "".join(ch for ch in str(value).lower() if ch.isalnum())
    return key in {"aura", "aurasummary", "summary"} or (
        "aura" in key and ("summary" in key or "fte" in key)
    )


def attach_workbook_inventory(args: argparse.Namespace) -> argparse.Namespace:
    """Attach sheet names plus lightweight Aura and Geo Cost name hints."""
    names = workbook_sheet_names(Path(args.excel))
    args.workbook_sheet_names = names
    args.aura_sheet_hints = [name for name in names if looks_like_aura_source_sheet_name(name)]
    args.geocost_sheet_hints = [
        {
            "sheet_name": name,
            "scope_hint": geocost_scope_hint_from_sheet_name(name),
        }
        for name in names
        if looks_like_geocost_sheet_name(name)
    ]
    return args


def prompt_text(prompt: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    return value or (default or "")


def prompt_yes_no(prompt: str, default: bool = True) -> bool:
    default_text = "Y/n" if default else "y/N"
    while True:
        value = input(f"{prompt} [{default_text}]: ").strip().lower()
        if not value:
            return default
        if value in {"y", "yes"}:
            return True
        if value in {"n", "no"}:
            return False
        print("Please answer y or n.")


def prompt_existing_file(
    label: str,
    base_dir: Path,
    suffixes: set[str],
    default: Path | None = None,
    candidates: list[Path] | None = None,
) -> Path:
    if candidates:
        print(f"\nAvailable {label} files:")
        for idx, candidate in enumerate(candidates[:10], start=1):
            print(f"  {idx}. {candidate}")

    while True:
        default_text = str(default) if default and default.exists() else None
        value = prompt_text(f"Enter {label} path, or choose a number above", default_text)
        if candidates and value.isdigit():
            index = int(value)
            if 1 <= index <= len(candidates[:10]):
                return candidates[index - 1].resolve()
            print("That number is not in the list.")
            continue

        path = resolve_path(value, base_dir)
        if not path.exists():
            print(f"File not found: {path}")
            continue
        if path.suffix.lower() not in suffixes:
            print(f"Expected one of these file types: {', '.join(sorted(suffixes))}")
            continue
        return path


def maybe_prompt_template(args: argparse.Namespace, base_dir: Path) -> Path:
    current = (
        resolve_path(args.template, base_dir)
        if getattr(args, "template", None)
        else DEFAULT_TEMPLATE_PPT
    )
    if current.exists() and not getattr(args, "interactive_input", False):
        return current
    if current.exists() and prompt_yes_no(f"Use template {current.name}?", default=True):
        return current
    ppt_candidates: list[Path] = []
    for directory in _candidate_search_dirs(base_dir):
        ppt_candidates.extend(
            path for path in directory.iterdir()
            if path.is_file() and path.suffix.lower() in PPT_SUFFIXES
        )
    ppt_candidates = sorted(
        {path.resolve() for path in ppt_candidates},
        key=lambda item: item.stat().st_mtime,
        reverse=True,
    )
    return prompt_existing_file(
        "PowerPoint template",
        base_dir,
        PPT_SUFFIXES,
        default=current,
        candidates=ppt_candidates,
    )


def prompt_output_dir(args: argparse.Namespace, base_dir: Path) -> Path:
    current = (
        resolve_path(args.output_dir, base_dir)
        if getattr(args, "output_dir", None)
        else DEFAULT_OUTPUT_DIR
    )
    if not getattr(args, "interactive_input", False):
        current.mkdir(parents=True, exist_ok=True)
        return current
    value = prompt_text("Output folder", str(current))
    path = resolve_path(value, base_dir)
    path.mkdir(parents=True, exist_ok=True)
    return path


def prompt_output_name(args: argparse.Namespace) -> str:
    current = getattr(args, "output_name", None) or "Generated_ThinkCell_Deck.pptx"
    if not getattr(args, "interactive_input", False):
        return current
    value = prompt_text("Output PowerPoint file name", current)
    if not value.lower().endswith((".pptx", ".pptm")):
        value += ".pptx"
    return value


def resolve_client_workbook(
    args: argparse.Namespace,
    base_dir: Path,
) -> argparse.Namespace:
    excel_path = getattr(args, "excel", None)
    if excel_path is not None and not getattr(args, "interactive_input", False):
        args.excel = resolve_path(excel_path, base_dir)
        if not args.excel.exists():
            raise FileNotFoundError(f"Excel workbook not found: {args.excel}")
        return attach_workbook_inventory(args)

    if getattr(args, "no_input_prompt", False):
        raise SystemExit("--excel is required when --no-input-prompt is used.")

    candidates = excel_candidates(base_dir)
    default = (
        resolve_path(excel_path, base_dir)
        if excel_path is not None
        else (candidates[0] if candidates else None)
    )
    print(
        "\nUpload or copy the client Excel workbook into the App or Aura folder, "
        "then provide its path here. The same Aura Summary table feeds Summary, L0, and L1; "
        "the same workbook also supplies any Geo Cost tabs."
    )
    args.excel = prompt_existing_file(
        "Excel workbook",
        base_dir,
        EXCEL_SUFFIXES,
        default=default,
        candidates=candidates,
    )
    return attach_workbook_inventory(args)


def resolve_run_inputs(
    args: argparse.Namespace,
    base_dir: Path | None = None,
    *,
    resolve_template: bool = True,
) -> argparse.Namespace:
    """Resolve one client workbook plus workflow templates/output settings."""
    base_dir = (base_dir or Path(__file__).parent).resolve()
    if resolve_template:
        args.template = maybe_prompt_template(args, base_dir)
    args.output_dir = prompt_output_dir(args, base_dir)
    args.output_name = prompt_output_name(args)
    args = resolve_client_workbook(args, base_dir)

    aura_hints = getattr(args, "aura_sheet_hints", [])
    if aura_hints:
        print("\nAura source tab name hint(s) found in the selected workbook:")
        for name in aura_hints:
            print(f"  {name}")
        print("The shared Aura detector will validate the table before Summary, L0, or L1 generation.")

    hints = getattr(args, "geocost_sheet_hints", [])
    if hints:
        print("\nGeo Cost tab name hints found in the selected workbook:")
        for item in hints:
            print(f"  {item['sheet_name']} ({item['scope_hint'] or 'unknown scope'})")
        print("Structural Global/US validation will run before slide generation.")
    return args


if __name__ == "__main__":
    namespace = argparse.Namespace(
        excel=None,
        template=DEFAULT_TEMPLATE_PPT,
        output_dir=DEFAULT_OUTPUT_DIR,
        output_name="Generated_ThinkCell_Deck.pptx",
        interactive_input=True,
        no_input_prompt=False,
    )
    resolved = resolve_run_inputs(namespace)
    print("\nResolved inputs:")
    print(f"  Excel:    {resolved.excel}")
    print(f"  Template: {resolved.template}")
    print(f"  Output:   {resolved.output_dir / resolved.output_name}")
    print(f"  Sheets:   {resolved.workbook_sheet_names}")
    sys.exit(0)
