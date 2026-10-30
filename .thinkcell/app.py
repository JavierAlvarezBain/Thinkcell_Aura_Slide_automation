from __future__ import annotations


import os
import sys
from pathlib import Path

import streamlit as st


ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from UI.canvas import (  # noqa: E402
    apply_theme,
    initialize_canvas_state,
    inspect_uploaded_workbook,
    render_canvas,
)
from UI.control_panel import (  # noqa: E402
    AppPaths,
    WorkbookLoadProgress,
    initialize_control_panel_state,
    render_control_panel,
    render_input_sidebar,
    render_segment_sidebar,
    validate_project_layout,
    announce_ui_diagnostics,
    ui_diagnostic_log_path,
    ui_log,
    ui_stage,
)


ROOT_UI_BUILD_ID = "root-workbook-loader-v2026-09-19-r67"


def _first_existing(candidates: list[Path], fallback: Path) -> Path:
    for candidate in candidates:
        if candidate.exists():
            return candidate.resolve()
    return fallback.resolve()


def resolve_app_paths() -> AppPaths:
    """Resolve the new root layout while tolerating one legacy backend layout."""
    slide_maker_candidates = [
        ROOT_DIR / "App" / "Slide Maker",
        ROOT_DIR / "Slide Maker",
    ]
    slide_maker_dir = _first_existing(
        slide_maker_candidates,
        ROOT_DIR / "App" / "Slide Maker",
    )
    app_dir = slide_maker_dir.parent

    executable_name = "python.exe" if os.name == "nt" else "python"
    project_python = _first_existing(
        [
            ROOT_DIR / "Scripts" / executable_name,
            ROOT_DIR.parent / "Scripts" / executable_name,
        ],
        ROOT_DIR / "Scripts" / executable_name,
    )
    python_exe = project_python if project_python.is_file() else Path(sys.executable)

    return AppPaths(
        repository_root=ROOT_DIR,
        ui_dir=ROOT_DIR / "UI",
        app_dir=app_dir,
        slide_maker_dir=slide_maker_dir,
        backend_cli=slide_maker_dir / "slide_maker_cli.py",
        project_python=project_python,
        python_exe=python_exe.resolve(),
    )


def main() -> None:
    pass_started = __import__("time").perf_counter()
    diagnostics_path = announce_ui_diagnostics()
    ui_log(
        "STREAMLIT PASS START",
        build=ROOT_UI_BUILD_ID,
        repository_root=str(ROOT_DIR),
        diagnostic_log=str(diagnostics_path),
    )

    with ui_stage("ui.set_page_config", warn_after=5, stack_after=20):
        st.set_page_config(
            page_title="ThinkCell Slide Maker",
            page_icon="📊",
            layout="wide",
            initial_sidebar_state="expanded",
        )
    apply_theme()
    initialize_control_panel_state()
    initialize_canvas_state()

    paths = resolve_app_paths()
    layout_problems = validate_project_layout(paths)
    ui_log(
        "PROJECT PATHS",
        backend_cli=str(paths.backend_cli),
        python_exe=str(paths.python_exe),
        layout_problems=layout_problems,
    )

    # Reserve a canvas-level status area before the sidebar performs any
    # workbook work. A new upload can therefore show immediate, progressively
    # updated feedback while the Python process copies, validates, detects, and
    # prepares the formatted preview.
    with st.container(key="workbook_load_status"):
        workbook_loader_placeholder = st.empty()
    workbook_load_progress = WorkbookLoadProgress(workbook_loader_placeholder)

    # The native Streamlit sidebar is the primary dark control rail. The
    # standardized Inputs rail and optional Segment rail are rendered after the
    # workbook analysis so they can consume the detected state in the same pass.
    with ui_stage("ui.render_control_panel", warn_after=15, stack_after=40):
        with st.sidebar:
            with st.container(key="control_panel"):
                controls = render_control_panel(
                    paths,
                    layout_problems,
                    inspect_uploaded_workbook,
                    workbook_load_progress,
                )

    try:
        with ui_stage(
            "ui.render_canvas",
            warn_after=20,
            stack_after=45,
            workbook=str(controls.workbook_path or ""),
            selected_sections=list(controls.selected_sections),
        ):
            with st.container(key="canvas"):
                render_canvas(
                    paths,
                    controls,
                    layout_problems,
                    workbook_load_progress,
                )
    except Exception as exc:
        workbook_load_progress.fail(f"{type(exc).__name__}: {exc}")
        raise

    with ui_stage("ui.render_input_sidebar", warn_after=10, stack_after=30):
        render_input_sidebar(controls)
    with ui_stage("ui.render_segment_sidebar", warn_after=10, stack_after=30):
        render_segment_sidebar(controls)

    ui_log(
        "STREAMLIT PASS DONE",
        elapsed_seconds=round(__import__("time").perf_counter() - pass_started, 3),
        workbook=str(controls.workbook_path or ""),
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        ui_log(
            "STREAMLIT PASS FATAL",
            level="ERROR",
            error=f"{type(exc).__name__}: {exc}",
            diagnostic_log=str(ui_diagnostic_log_path()),
        )
        raise
