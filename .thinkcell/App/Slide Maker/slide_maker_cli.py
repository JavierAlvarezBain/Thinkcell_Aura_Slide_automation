from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

from shared import AURA_L0_DIR, AURA_ROOT


WORKFLOWS = {
    "aura-l1": {
        "path": AURA_ROOT / "Aura L1" / "run_aura_l1.py",
        "description": "Aura Summary + Aura L0 + Aura L1 + detected Global/US Geo Cost slides",
    },
    "aura-l0": {
        "path": AURA_L0_DIR / "run_aura_l0.py",
        "description": "Aura L0 slides only, using the dedicated L0 build and native patch loop",
    },
    "aura-summary": {
        "path": AURA_ROOT / "Aura Summary" / "run_aura_summary.py",
        "description": "Aura Summary Bain table slide only",
    },
    "geocost": {
        "path": AURA_ROOT / "Global Geo Cost" / "run_geocost.py",
        "description": "Detected Global and/or US Geo Cost slides only",
    },
}

ALIASES = {
    "aura": "aura-l1",
    "aura-l1": "aura-l1",
    "aura_l1": "aura-l1",
    "aura l1": "aura-l1",
    "l1": "aura-l1",
    "aura-l0": "aura-l0",
    "aura_l0": "aura-l0",
    "aura l0": "aura-l0",
    "l0": "aura-l0",
    "summary": "aura-summary",
    "aura-summary": "aura-summary",
    "aura_summary": "aura-summary",
    "aura summary": "aura-summary",
    "geocost": "geocost",
    "geo-cost": "geocost",
    "geo_cost": "geocost",
    "geo cost": "geocost",
    "global-geocost": "geocost",
    "global geo cost": "geocost",
}

DEFAULT_WORKFLOW = "aura-l1"


def print_root_help() -> None:
    print(
        """
ThinkCell Automation

Usage:
  python app.py [workflow] [workflow options]

Examples:
  python app.py
  python app.py aura-l1 --excel "C:\\path\\client.xlsx"
  python app.py aura-l1 --skip-l0
  python app.py aura-l1 --skip-geocost
  python app.py aura-l0 --excel "C:\\path\\client.xlsx"
  python app.py aura-summary --excel "C:\\path\\client.xlsx"
  python app.py geocost --excel "C:\\path\\client.xlsx" --scope auto
  python app.py geocost --chart-data-only --excel "C:\\path\\client.xlsx"

Available workflows:
"""
    )
    for name, info in WORKFLOWS.items():
        print(f"  {name:<14} {info['description']}")

    print(
        """
Notes:
  If no workflow is provided, aura-l1 is used.
  The default aura-l1 workflow builds Summary, L0, L1, then auto-detected Geo Cost sections.
  Aura L0 is finalized only after the consolidated deck has been saved and closed.
  Workflow-specific flags are passed through to the selected runner.
"""
    )


def normalize_workflow_name(value: str) -> str:
    return ALIASES.get(value.strip().lower(), value.strip().lower())


def load_module_from_path(module_name: str, path: Path) -> ModuleType:
    if not path.exists():
        raise FileNotFoundError(
            f"Workflow runner not found: {path}\n\n"
            "Expected folder structure:\n"
            "  Aura\\app.py\n"
            "  Aura\\input.py\n"
            "  Aura\\shared.py\n"
            "  Aura\\Aura L0\\run_aura_l0.py\n"
            "  Aura\\Aura L1\\run_aura_l1.py\n"
            "  Aura\\Aura Summary\\run_aura_summary.py\n"
            "  Aura\\Global Geo Cost\\run_geocost.py"
        )

    spec = importlib.util.spec_from_file_location(module_name, str(path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load workflow runner: {path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_workflow(workflow_name: str, workflow_args: list[str]) -> None:
    workflow_name = normalize_workflow_name(workflow_name)
    if workflow_name not in WORKFLOWS:
        print(f"Unknown workflow: {workflow_name}\n")
        print_root_help()
        raise SystemExit(2)

    runner_path = WORKFLOWS[workflow_name]["path"]
    workflow_dir = runner_path.parent

    # Workflow-local modules first; universal Aura root files second.
    for path in [str(AURA_ROOT), str(workflow_dir)]:
        while path in sys.path:
            sys.path.remove(path)
    sys.path.insert(0, str(AURA_ROOT))
    sys.path.insert(0, str(workflow_dir))

    module = load_module_from_path(
        f"{workflow_name.replace('-', '_')}_runner",
        runner_path,
    )
    if not hasattr(module, "main"):
        raise RuntimeError(f"Workflow runner has no main(argv) function: {runner_path}")
    module.main(workflow_args)


def main(argv: list[str] | None = None) -> None:
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] in {"-h", "--help", "help"}:
        print_root_help()
        return
    if args and args[0] in {"--list", "--list-workflows", "list"}:
        for name, info in WORKFLOWS.items():
            print(f"{name}: {info['description']}")
        return
    if args and not args[0].startswith("-"):
        workflow_name = args.pop(0)
    else:
        workflow_name = DEFAULT_WORKFLOW
    run_workflow(workflow_name, args)


if __name__ == "__main__":
    main()
