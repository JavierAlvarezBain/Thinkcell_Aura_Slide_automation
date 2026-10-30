from __future__ import annotations

"""Left-side controls for the ThinkCell Slide Maker Streamlit application.

This module owns the collapsible control rail: source upload, build selection,
standardized input mapping, output settings, and the Generate action. The rail opens
at roughly 25% width and can collapse to a compact 2.5% launcher. Slide-input
mapping lives in a second independently collapsible rail rendered beside it; workbook
rendering and the live build console live in ``UI/canvas.py``.
"""

import hashlib
import html
import json
import os
import re
import sys
import tempfile
import zipfile
import faulthandler
import threading
import time
import traceback
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

import streamlit as st
import streamlit.components.v1 as components


SUPPORTED_WORKBOOK_SUFFIXES = {".xlsx", ".xlsm"}
PRESENTATION_SUFFIXES = {".pptx", ".pptm"}


# -----------------------------------------------------------------------------
# UI diagnostics
# -----------------------------------------------------------------------------

_UI_DIAGNOSTIC_LOCK = threading.RLock()
_UI_DIAGNOSTIC_ANNOUNCED = False
_UI_DIAGNOSTIC_PROCESS_STARTED = time.time()
_UI_DIAGNOSTIC_BUILD_ID = "ui-segment-workspace-restored-v2026-09-21-r78"


def ui_diagnostic_log_path() -> Path:
    """Return the persistent Streamlit/UI diagnostic log path.

    The file is deliberately outside the repository so it remains writable even
    when the project is installed in a synchronized or read-only folder.
    """
    base = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
    folder = base / "ThinkCellSlideMaker" / "StreamlitUI" / "logs"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "ui_runtime.log"


def _diagnostic_value(value: Any) -> str:
    text = repr(value)
    return text if len(text) <= 500 else text[:497] + "..."


def ui_log(event: str, *, level: str = "INFO", **fields: Any) -> None:
    """Write one immediate diagnostic line to the shell and persistent log."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    details = " ".join(
        f"{key}={_diagnostic_value(value)}" for key, value in fields.items()
    )
    line = (
        f"{timestamp} | TC-UI | {level.upper():<7} | pid={os.getpid()} | "
        f"{event}" + (f" | {details}" if details else "")
    )
    with _UI_DIAGNOSTIC_LOCK:
        print(line, file=sys.stderr, flush=True)
        try:
            path = ui_diagnostic_log_path()
            with path.open("a", encoding="utf-8", errors="replace") as handle:
                handle.write(line + "\n")
        except Exception:
            # Diagnostics must never become the reason the UI fails.
            pass


def _dump_ui_threads(stage_name: str) -> None:
    """Dump all Python thread stacks to both shell and the persistent log."""
    ui_log("WATCHDOG STACK DUMP", level="WARNING", stage=stage_name)
    try:
        faulthandler.dump_traceback(file=sys.stderr, all_threads=True)
    except Exception:
        pass
    try:
        path = ui_diagnostic_log_path()
        with path.open("a", encoding="utf-8", errors="replace") as handle:
            handle.write(
                f"\n--- THREAD STACKS: {stage_name} @ {datetime.now().isoformat()} ---\n"
            )
            handle.flush()
            faulthandler.dump_traceback(file=handle, all_threads=True)
            handle.write("--- END THREAD STACKS ---\n\n")
    except Exception:
        pass


def announce_ui_diagnostics() -> Path:
    """Print the diagnostic path once per Streamlit server process."""
    global _UI_DIAGNOSTIC_ANNOUNCED
    path = ui_diagnostic_log_path()
    if not _UI_DIAGNOSTIC_ANNOUNCED:
        _UI_DIAGNOSTIC_ANNOUNCED = True
        try:
            faulthandler.enable(all_threads=True)
        except Exception:
            pass
        ui_log(
            "DIAGNOSTICS READY",
            build=_UI_DIAGNOSTIC_BUILD_ID,
            log_file=str(path),
            python=sys.executable,
            process_uptime_seconds=round(time.time() - _UI_DIAGNOSTIC_PROCESS_STARTED, 3),
        )
    return path


@contextmanager
def ui_stage(
    name: str,
    *,
    warn_after: float = 20.0,
    stack_after: float | None = 45.0,
    **fields: Any,
):
    """Log START/DONE/FAILED and expose long-running UI work in the shell.

    A daemon watchdog prints a warning while the main Streamlit thread is still
    blocked. At ``stack_after`` seconds it also dumps every Python thread stack,
    making a silent openpyxl/ZIP/HTML bottleneck diagnosable from the shell.
    """
    announce_ui_diagnostics()
    started = time.perf_counter()
    finished = threading.Event()
    warn_after = max(1.0, float(warn_after))
    stack_after_value = None if stack_after is None else max(warn_after, float(stack_after))

    def watchdog() -> None:
        if finished.wait(warn_after):
            return
        elapsed = time.perf_counter() - started
        ui_log("STILL RUNNING", level="WARNING", stage=name, elapsed_seconds=round(elapsed, 2))
        if stack_after_value is not None:
            remaining = max(0.0, stack_after_value - elapsed)
            if finished.wait(remaining):
                return
            _dump_ui_threads(name)
            while not finished.wait(60.0):
                ui_log(
                    "STILL RUNNING",
                    level="WARNING",
                    stage=name,
                    elapsed_seconds=round(time.perf_counter() - started, 2),
                )

    thread = threading.Thread(
        target=watchdog,
        name=f"tc-ui-watchdog-{re.sub(r'[^a-z0-9]+', '-', name.lower()).strip('-')[:48]}",
        daemon=True,
    )
    ui_log("START", stage=name, **fields)
    thread.start()
    try:
        yield
    except Exception as exc:
        ui_log(
            "FAILED",
            level="ERROR",
            stage=name,
            elapsed_seconds=round(time.perf_counter() - started, 3),
            error=f"{type(exc).__name__}: {exc}",
        )
        with _UI_DIAGNOSTIC_LOCK:
            traceback.print_exc(file=sys.stderr)
            try:
                with ui_diagnostic_log_path().open(
                    "a", encoding="utf-8", errors="replace"
                ) as handle:
                    traceback.print_exc(file=handle)
            except Exception:
                pass
        raise
    else:
        ui_log(
            "DONE",
            stage=name,
            elapsed_seconds=round(time.perf_counter() - started, 3),
        )
    finally:
        finished.set()


DECK_SECTION_LABELS = {
    "summary": "Summary",
    "l0": "Aura L0",
    "l1": "Aura L1",
    "geocost": "Geo Cost",
}
DECK_SECTION_ORDER = tuple(DECK_SECTION_LABELS)
ALL_DECK_SECTIONS = frozenset(DECK_SECTION_ORDER)
SORTABLE_DECK_SECTIONS = frozenset({"l0", "l1", "geocost"})

# Compatibility map used to migrate browser sessions created by the former
# single-select "Entire deck / ... only" control.
LEGACY_WORKFLOW_SECTIONS = {
    "complete": DECK_SECTION_ORDER,
    "summary": ("summary",),
    "l0": ("l0",),
    "l1": ("l1",),
    "geocost": ("geocost",),
}

INPUT_SECTION_LABELS = {
    "executive_summary": "Summary",
    "l0": "Aura L0",
    "l1": "Aura L1",
    "geocost": "Geo Cost",
}

INPUT_SECTION_BY_DECK_SECTION = {
    "summary": "executive_summary",
    "l0": "l0",
    "l1": "l1",
    "geocost": "geocost",
}


# Standard Aura range inputs shared by the Aura L0 and Aura L1 builders.
# Company names are selected as a value-only row. Revenue and FTE are
# self-describing ranges. Quartile benchmarks are selected as complete
# three-column tables instead of three separate columns; the detector resolves
# Top / Median / Bottom semantically even when their left-to-right order changes.
AURA_INPUT_FIELDS: tuple[dict[str, str], ...] = (
    {
        "key": "taxonomy",
        "label": "Taxonomy",
        "short_label": "Taxonomy",
        "help": (
            "Select the complete two-column L0/L1 taxonomy table, including its "
            "L0 and L1 header cells. Supported pairs include L0/L1, Bain L0/Bain L1, "
            "and VC L0/VC L1."
        ),
    },
    {
        "key": "company_headers",
        "label": "Company names",
        "short_label": "Company names",
        "help": (
            "Select only the horizontal company-name row. The first selected cell "
            "is the client and the remaining cells are peer companies; do not include "
            "a label cell to the left."
        ),
    },
    {
        "key": "revenue_values",
        "label": "Revenue row",
        "short_label": "Revenue + label",
        "help": (
            "Select one horizontal row beginning with the Revenue master label, "
            "followed by one revenue value per company."
        ),
    },
    {
        "key": "fte_values",
        "label": "Function + FTE / $B table",
        "short_label": "Function + FTE / $B",
        "help": (
            "Select the complete source table beginning with the L1/function column, "
            "followed by the company FTE / $B columns. Include the header row and all "
            "function rows. Do not include an L0 column; L0 is always taken from Taxonomy."
        ),
    },
    {
        "key": "l0_quartile_table",
        "label": "L0 quartile table detector",
        "short_label": "L0 quartile table",
        "help": (
            "Select the three consecutive L0 benchmark columns containing Top, Median, "
            "and Bottom quartiles. The columns may appear in any order. Auto-detection "
            "uses the nearest matching three-column header and the L0 labels on the same rows."
        ),
    },
    {
        "key": "l1_quartile_table",
        "label": "L1 quartile table detector",
        "short_label": "L1 quartile table",
        "help": (
            "Select the three consecutive L1 benchmark columns containing Top, Median, "
            "and Bottom quartiles. Include the header row when selecting manually; the "
            "columns may appear in any order."
        ),
    },
)
AURA_INPUT_KEYS = tuple(item["key"] for item in AURA_INPUT_FIELDS)
AURA_INPUT_BY_KEY = {item["key"]: item for item in AURA_INPUT_FIELDS}
AURA_SHARED_INPUT_KEYS = (
    "taxonomy",
    "company_headers",
    "revenue_values",
    "fte_values",
)
AURA_L0_QUARTILE_INPUT_KEYS = ("l0_quartile_table",)
AURA_L1_BENCHMARK_INPUT_KEYS = ("l1_quartile_table",)
AURA_INPUT_MAP_SCHEMA_VERSION = 9

# Geo Cost uses one visible Overall geography table plus the dedicated
# low-cost-percentage row that feeds the bottom row of the think-cell chart.
# Region / High-Low / raw geography helper columns and the >>> function-marker
# row are inferred around these two user-facing ranges. Subsequent L0/L1 blocks
# reuse the same table rows and low-cost row in columns to the right.
GEOCOST_INPUT_FIELDS: tuple[dict[str, str], ...] = (
    {
        "key": "geocost_table",
        "label": "Geo Cost table",
        "short_label": "Geo Cost table",
        "help": (
            "Select the complete Overall geography matrix from its local Country/State "
            "column through the final Overall company column, including the company "
            "header row and every geography row. Do not include the Region, High/Low, "
            "or raw-country helper columns on the far left; those are inferred automatically."
        ),
    },
    {
        "key": "geocost_data_row",
        "label": "Geo Cost data row",
        "short_label": "Low-cost FTE data row",
        "help": (
            "Select the horizontal '% of FTEs in Low Cost region' row that appears below "
            "the geography table. Include its row label and all Overall company percentage "
            "values. This is the data row shown at the bottom of the Geo Cost chart."
        ),
    },
)
GEOCOST_INPUT_KEYS = tuple(item["key"] for item in GEOCOST_INPUT_FIELDS)
GEOCOST_INPUT_BY_KEY = {item["key"]: item for item in GEOCOST_INPUT_FIELDS}
GEOCOST_INPUT_MAP_SCHEMA_VERSION = 2

STANDARD_INPUT_FIELDS = AURA_INPUT_FIELDS + GEOCOST_INPUT_FIELDS
STANDARD_INPUT_KEYS = tuple(item["key"] for item in STANDARD_INPUT_FIELDS)
STANDARD_INPUT_BY_KEY = {item["key"]: item for item in STANDARD_INPUT_FIELDS}

# Backward-compatible aliases retained while the rest of the codebase moves
# from the old L1-only mapper to the standardized Aura Inputs sidebar.
L1_INPUT_FIELDS = AURA_INPUT_FIELDS
L1_INPUT_KEYS = AURA_INPUT_KEYS
L1_INPUT_BY_KEY = AURA_INPUT_BY_KEY
L1_INPUT_MAP_SCHEMA_VERSION = AURA_INPUT_MAP_SCHEMA_VERSION

# Aura L1 company/bar ordering. The leftmost selected company remains the
# client in every mode; only peer companies from index 1 onward are reordered.
L1_SORT_ORDER_LABELS = {
    "revenue_ascending": "Basis Increasing revenue",
    "revenue_descending": "Basis Decreasing revenue",
    "sheet_order": "Sheet order (default Excel order)",
    "segment": "Sort by segment",
}
L1_SORT_ORDER_DEFAULT = "revenue_descending"
L1_SEGMENT_SORT_DIRECTION_LABELS = {
    "revenue_ascending": "Increasing revenue",
    "revenue_descending": "Decreasing revenue",
}
L1_SEGMENT_SORT_DIRECTION_DEFAULT = "revenue_descending"


def normalize_deck_sections(value: Any) -> tuple[str, ...]:
    """Return a de-duplicated canonical slide-section selection.

    ``workflow_mode`` used to contain one legacy string. It now stores the
    multi-select values, so this helper also migrates old browser sessions.
    """
    if isinstance(value, str):
        raw_values = LEGACY_WORKFLOW_SECTIONS.get(value, (value,))
    elif isinstance(value, (list, tuple, set, frozenset)):
        raw_values = tuple(str(item) for item in value)
    else:
        raw_values = ()
    selected = {item for item in raw_values if item in ALL_DECK_SECTIONS}
    return tuple(item for item in DECK_SECTION_ORDER if item in selected)


def workflow_mode_from_sections(sections: Any) -> str:
    """Return a compact mode name retained for progress/UI compatibility."""
    normalized = normalize_deck_sections(sections)
    if set(normalized) == set(DECK_SECTION_ORDER):
        return "complete"
    if len(normalized) == 1:
        return normalized[0]
    if normalized:
        return "custom"
    return "none"


def aura_input_keys_for_sections(sections: Any) -> tuple[str, ...]:
    """Return only Aura inputs required by the selected L0/L1 outputs."""
    selected = set(normalize_deck_sections(sections))
    keys: list[str] = []
    if selected.intersection({"l0", "l1"}):
        keys.extend(AURA_SHARED_INPUT_KEYS)
    if "l0" in selected:
        keys.extend(AURA_L0_QUARTILE_INPUT_KEYS)
    if "l1" in selected:
        keys.extend(AURA_L1_BENCHMARK_INPUT_KEYS)
    return tuple(keys)


def geocost_input_keys_for_sections(sections: Any) -> tuple[str, ...]:
    """Return the two standardized Geo Cost wrapper inputs when requested."""
    selected = set(normalize_deck_sections(sections))
    return GEOCOST_INPUT_KEYS if "geocost" in selected else ()


def standardized_input_keys_for_sections(sections: Any) -> tuple[str, ...]:
    """Return the one ordered input list used by the standardized Inputs rail."""
    return (
        aura_input_keys_for_sections(sections)
        + geocost_input_keys_for_sections(sections)
    )


@dataclass(frozen=True)
class AppPaths:
    """Resolved paths for the root UI and the unchanged Slide Maker backend."""

    repository_root: Path
    ui_dir: Path
    app_dir: Path
    slide_maker_dir: Path
    backend_cli: Path
    project_python: Path
    python_exe: Path

    @property
    def backend_working_dir(self) -> Path:
        return self.slide_maker_dir.parent


@dataclass(frozen=True)
class BuildOptions:
    mode: str
    selected_sections: tuple[str, ...]
    workbook: Path
    output_dir: Path
    output_name: str
    input_section: str = ""
    aura_sheet: str = ""
    taxonomy_sheet: str = ""
    geocost_global_sheet: str = ""
    geocost_us_sheet: str = ""
    only_l0: str = ""
    aura_input_mappings: dict[str, dict[str, Any]] = field(default_factory=dict)
    geocost_input_mappings: dict[str, dict[str, Any]] = field(default_factory=dict)
    # Legacy alias retained for older callers; both fields contain the same
    # standardized Aura mapping payload.
    l1_input_mappings: dict[str, dict[str, Any]] = field(default_factory=dict)
    l1_sort_order: str = L1_SORT_ORDER_DEFAULT
    l1_segment_sort_direction: str = L1_SEGMENT_SORT_DIRECTION_DEFAULT
    l1_segment_profile: dict[str, Any] = field(default_factory=dict)
    ignored_aura_taxonomy_functions: tuple[str, ...] = ()
    keep_temp: bool = False
    debug: bool = False


@dataclass(frozen=True)
class UploadInspection:
    preview_path: Path
    metadata: dict[str, Any] = field(default_factory=dict)
    preview_error: str | None = None
    relevant_tab_count: int = 0


@dataclass
class WorkbookLoadProgress:
    """Canvas-level workbook upload/preparation indicator.

    Streamlit's native uploader owns the browser-to-server transfer. Once the
    uploaded file is available to Python, this indicator takes over and shows
    the user exactly which server-side step is running: copying, validating,
    detecting workbook inputs, and preparing the formatted preview.

    The object is deliberately lightweight and passed from ``app.py`` into the
    control panel and canvas. Methods become no-ops unless a genuinely new
    upload has started, so ordinary reruns do not flash the loader.
    """

    placeholder: Any
    active: bool = False
    failed: bool = False
    file_name: str = ""
    file_size: int = 0
    started_at: float = 0.0
    percent: int = 0
    phase: str = "upload"

    _PHASE_INDEX = {
        "upload": 0,
        "verify": 1,
        "detect": 2,
        "preview": 3,
        "complete": 4,
        "error": -1,
    }
    _STEPS = ("Upload", "Verify", "Detect", "Preview")

    @staticmethod
    def _size_text(size_bytes: int) -> str:
        size = max(0, int(size_bytes or 0))
        if size >= 1024 ** 3:
            return f"{size / (1024 ** 3):.2f} GB"
        if size >= 1024 ** 2:
            return f"{size / (1024 ** 2):.1f} MB"
        if size >= 1024:
            return f"{size / 1024:.1f} KB"
        return f"{size} bytes" if size else "Workbook"

    def begin(self, file_name: str, file_size: int = 0) -> None:
        self.active = True
        self.failed = False
        self.file_name = str(file_name or "Workbook")
        self.file_size = max(0, int(file_size or 0))
        self.started_at = time.perf_counter()
        self.percent = 2
        self.phase = "upload"
        self.update(
            2,
            "Uploading workbook",
            "Receiving the selected file and preparing a safe local copy.",
            phase="upload",
            force=True,
        )

    def update(
        self,
        percent: int | float,
        title: str,
        detail: str = "",
        *,
        phase: str | None = None,
        force: bool = False,
    ) -> None:
        if not self.active and not force:
            return
        if phase:
            self.phase = str(phase)
        next_percent = max(0, min(100, int(round(float(percent)))))
        self.percent = max(self.percent, next_percent) if not force else next_percent
        self._render(title=title, detail=detail)

    def complete(self, detail: str = "Workbook preview is ready.") -> None:
        if not self.active:
            return
        self.phase = "complete"
        self.percent = 100
        self._render(title="Workbook ready", detail=detail, complete=True)

    def fail(self, detail: str) -> None:
        if not self.active:
            return
        self.failed = True
        self.phase = "error"
        self._render(
            title="Workbook could not be prepared",
            detail=str(detail or "Review the shell diagnostics for details."),
            error=True,
        )

    def clear(self, *, minimum_visible_seconds: float = 0.35) -> None:
        if not self.active:
            return
        elapsed = time.perf_counter() - self.started_at if self.started_at else 0.0
        remaining = max(0.0, float(minimum_visible_seconds) - elapsed)
        if remaining:
            time.sleep(min(remaining, 0.35))
        try:
            self.placeholder.empty()
        except Exception:
            pass
        self.active = False

    def _render(
        self,
        *,
        title: str,
        detail: str,
        complete: bool = False,
        error: bool = False,
    ) -> None:
        phase_index = self._PHASE_INDEX.get(self.phase, 0)
        step_nodes: list[str] = []
        for index, label in enumerate(self._STEPS):
            if complete or phase_index > index:
                state_class = "is-done"
                marker = "✓"
            elif phase_index == index and not error:
                state_class = "is-active"
                marker = str(index + 1)
            else:
                state_class = ""
                marker = str(index + 1)
            step_nodes.append(
                "<div class=\"tc-upload-loader-step "
                + state_class
                + "\"><span class=\"tc-upload-loader-step-marker\">"
                + html.escape(marker)
                + "</span><span>"
                + html.escape(label)
                + "</span></div>"
            )

        state_class = "is-error" if error else ("is-complete" if complete else "is-running")
        elapsed = max(0.0, time.perf_counter() - self.started_at) if self.started_at else 0.0
        file_meta = self._size_text(self.file_size)
        if self.file_name:
            file_meta = f"{self.file_name} · {file_meta}"
        elapsed_text = f"{elapsed:.1f}s"
        icon_text = "!" if error else ("✓" if complete else "XLSX")

        markup = f"""
        <div class="tc-upload-loader-shell" role="status" aria-live="polite">
          <div class="tc-upload-loader-card {state_class}">
            <div class="tc-upload-loader-main">
              <div class="tc-upload-loader-visual" aria-hidden="true">
                <span class="tc-upload-loader-orbit"></span>
                <span class="tc-upload-loader-icon">{html.escape(icon_text)}</span>
              </div>
              <div class="tc-upload-loader-copy">
                <div class="tc-upload-loader-kicker">SOURCE WORKBOOK</div>
                <div class="tc-upload-loader-title">{html.escape(str(title))}</div>
                <div class="tc-upload-loader-detail">{html.escape(str(detail))}</div>
                <div class="tc-upload-loader-file">{html.escape(file_meta)}</div>
              </div>
              <div class="tc-upload-loader-metric">
                <strong>{self.percent}%</strong>
                <span>{html.escape(elapsed_text)}</span>
              </div>
            </div>
            <div class="tc-upload-loader-track" aria-hidden="true">
              <span style="width: {self.percent}%"></span>
            </div>
            <div class="tc-upload-loader-steps">{''.join(step_nodes)}</div>
          </div>
        </div>
        """
        try:
            self.placeholder.markdown(markup, unsafe_allow_html=True)
        except Exception:
            # The loader is purely presentational and must not break uploads.
            pass


@dataclass(frozen=True)
class ControlPanelState:
    workbook_path: Path | None
    preview_workbook_path: Path | None
    preview_metadata: dict[str, Any]
    preview_error: str | None
    workbook_error: str | None
    workbook_display_name: str
    is_new_upload: bool
    workflow_mode: str
    selected_sections: tuple[str, ...]
    input_section: str
    l1_sort_order: str
    l1_segment_sort_direction: str
    only_l0: str
    output_dir: Path
    output_name: str
    keep_temp: bool
    debug: bool
    workbook_finalized: bool
    generate_clicked: bool

    def build_options(self) -> BuildOptions:
        if self.workbook_path is None:
            raise RuntimeError("The uploaded workbook is unavailable.")

        source_signature = str(st.session_state.get("uploaded_signature") or "")
        mapping_root = st.session_state.get("slide_input_mappings") or {}
        workbook_mappings = (
            mapping_root.get(source_signature, {})
            if isinstance(mapping_root, dict) and source_signature
            else {}
        )
        standardized_mappings: dict[str, dict[str, Any]] = {}
        if isinstance(workbook_mappings, dict):
            standardized_mappings = dict(
                workbook_mappings.get("inputs")
                or workbook_mappings.get("aura")
                or workbook_mappings.get("l1")
                or {}
            )
        aura_input_mappings = {
            key: value
            for key, value in standardized_mappings.items()
            if key in AURA_INPUT_KEYS and isinstance(value, dict)
        }
        geocost_input_mappings = {
            key: value
            for key, value in standardized_mappings.items()
            if key in GEOCOST_INPUT_KEYS and isinstance(value, dict)
        }

        return BuildOptions(
            mode=self.workflow_mode,
            selected_sections=self.selected_sections,
            workbook=self.workbook_path,
            output_dir=self.output_dir,
            output_name=normalize_output_name(self.output_name),
            input_section=self.input_section,
            aura_sheet=str(
                st.session_state.get("selected_aura_sheet_override") or ""
            ).strip(),
            taxonomy_sheet=str(
                st.session_state.get("selected_taxonomy_sheet_override") or ""
            ).strip(),
            geocost_global_sheet=str(
                st.session_state.get("selected_geocost_global_sheet_override") or ""
            ).strip(),
            geocost_us_sheet=str(
                st.session_state.get("selected_geocost_us_sheet_override") or ""
            ).strip(),
            only_l0=(
                self.only_l0.strip()
                if self.selected_sections == ("l0",)
                else ""
            ),
            aura_input_mappings=aura_input_mappings,
            geocost_input_mappings=geocost_input_mappings,
            l1_input_mappings=aura_input_mappings,
            l1_sort_order=self.l1_sort_order,
            l1_segment_sort_direction=self.l1_segment_sort_direction,
            l1_segment_profile=dict(
                st.session_state.get("l1_segment_profile") or {}
            ),
            ignored_aura_taxonomy_functions=tuple(
                get_ignored_aura_taxonomy_functions(source_signature)
            ),
            keep_temp=self.keep_temp,
            debug=self.debug,
        )


def get_standard_input_mappings(source_signature: str) -> dict[str, dict[str, Any]]:
    """Return the unified Aura + Geo Cost mapping payload for one workbook.

    Earlier releases stored Aura ranges under ``aura`` / ``l1``. Migrate those
    values into the standardized ``inputs`` bucket without discarding any new
    Geo Cost mappings.
    """
    root = st.session_state.get("slide_input_mappings") or {}
    if not isinstance(root, dict) or not source_signature:
        return {}
    workbook_bucket = root.get(source_signature) or {}
    if not isinstance(workbook_bucket, dict):
        return {}

    merged: dict[str, dict[str, Any]] = {}
    for bucket_name in ("aura", "l1", "geocost", "inputs"):
        bucket = workbook_bucket.get(bucket_name) or {}
        if not isinstance(bucket, dict):
            continue
        for key, value in bucket.items():
            if key in STANDARD_INPUT_KEYS and isinstance(value, dict):
                merged[key] = value

    if merged and workbook_bucket.get("inputs") != merged:
        migrated_root = dict(root)
        migrated_bucket = dict(workbook_bucket)
        migrated_bucket["inputs"] = dict(merged)
        migrated_bucket["aura"] = {
            key: value for key, value in merged.items() if key in AURA_INPUT_KEYS
        }
        migrated_bucket["l1"] = dict(migrated_bucket["aura"])
        migrated_bucket["geocost"] = {
            key: value for key, value in merged.items() if key in GEOCOST_INPUT_KEYS
        }
        migrated_root[source_signature] = migrated_bucket
        st.session_state["slide_input_mappings"] = migrated_root
    return merged


def set_standard_input_mappings(
    source_signature: str,
    mappings: dict[str, dict[str, Any]],
) -> None:
    """Persist one unified input payload while keeping legacy buckets synced."""
    root = dict(st.session_state.get("slide_input_mappings") or {})
    workbook_bucket = dict(root.get(source_signature) or {})
    cleaned = {
        key: value
        for key, value in mappings.items()
        if key in STANDARD_INPUT_KEYS and isinstance(value, dict)
    }
    aura = {key: value for key, value in cleaned.items() if key in AURA_INPUT_KEYS}
    geocost = {
        key: value for key, value in cleaned.items() if key in GEOCOST_INPUT_KEYS
    }
    workbook_bucket["inputs"] = dict(cleaned)
    workbook_bucket["aura"] = dict(aura)
    workbook_bucket["l1"] = dict(aura)
    workbook_bucket["geocost"] = dict(geocost)
    root[source_signature] = workbook_bucket
    st.session_state["slide_input_mappings"] = root


def get_aura_input_mappings(source_signature: str) -> dict[str, dict[str, Any]]:
    return {
        key: value
        for key, value in get_standard_input_mappings(source_signature).items()
        if key in AURA_INPUT_KEYS
    }


def set_aura_input_mappings(
    source_signature: str,
    mappings: dict[str, dict[str, Any]],
) -> None:
    current = get_standard_input_mappings(source_signature)
    for key in AURA_INPUT_KEYS:
        current.pop(key, None)
    current.update(
        {key: value for key, value in mappings.items() if key in AURA_INPUT_KEYS}
    )
    set_standard_input_mappings(source_signature, current)


def get_geocost_input_mappings(source_signature: str) -> dict[str, dict[str, Any]]:
    return {
        key: value
        for key, value in get_standard_input_mappings(source_signature).items()
        if key in GEOCOST_INPUT_KEYS
    }


def set_geocost_input_mappings(
    source_signature: str,
    mappings: dict[str, dict[str, Any]],
) -> None:
    current = get_standard_input_mappings(source_signature)
    for key in GEOCOST_INPUT_KEYS:
        current.pop(key, None)
    current.update(
        {key: value for key, value in mappings.items() if key in GEOCOST_INPUT_KEYS}
    )
    set_standard_input_mappings(source_signature, current)


# Backward-compatible function names used by the current canvas module. They
# intentionally return the unified payload so the same range picker can bind
# Aura and Geo Cost inputs. Aura-specific readers filter their own keys.
def get_l1_input_mappings(source_signature: str) -> dict[str, dict[str, Any]]:
    return get_standard_input_mappings(source_signature)


def set_l1_input_mappings(
    source_signature: str,
    mappings: dict[str, dict[str, Any]],
) -> None:
    set_standard_input_mappings(source_signature, mappings)


def _taxonomy_mismatch_key(value: Any) -> str:
    return " ".join(str(value or "").split()).casefold()


def get_ignored_aura_taxonomy_functions(source_signature: str) -> list[str]:
    """Return user-ignored Taxonomy→FTE missing-function mismatches for one workbook."""
    root = st.session_state.get("aura_ignored_taxonomy_functions") or {}
    if not isinstance(root, dict) or not source_signature:
        return []
    values = root.get(source_signature) or []
    if not isinstance(values, (list, tuple, set)):
        return []
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = " ".join(str(value or "").split()).strip()
        key = _taxonomy_mismatch_key(text)
        if text and key and key not in seen:
            seen.add(key)
            result.append(text)
    return result


def add_ignored_aura_taxonomy_function(source_signature: str, function_name: str) -> None:
    if not source_signature:
        return
    text = " ".join(str(function_name or "").split()).strip()
    if not text:
        return
    root = dict(st.session_state.get("aura_ignored_taxonomy_functions") or {})
    current = get_ignored_aura_taxonomy_functions(source_signature)
    target_key = _taxonomy_mismatch_key(text)
    if target_key not in {_taxonomy_mismatch_key(value) for value in current}:
        current.append(text)
    root[source_signature] = current
    st.session_state["aura_ignored_taxonomy_functions"] = root


def remove_ignored_aura_taxonomy_function(source_signature: str, function_name: str) -> None:
    if not source_signature:
        return
    target_key = _taxonomy_mismatch_key(function_name)
    root = dict(st.session_state.get("aura_ignored_taxonomy_functions") or {})
    current = [
        value for value in get_ignored_aura_taxonomy_functions(source_signature)
        if _taxonomy_mismatch_key(value) != target_key
    ]
    if current:
        root[source_signature] = current
    else:
        root.pop(source_signature, None)
    st.session_state["aura_ignored_taxonomy_functions"] = root



def get_l1_segment_name_overrides(source_signature: str) -> dict[str, str]:
    """Return persisted segment display-name overrides for one workbook."""
    root = st.session_state.get("l1_segment_name_overrides") or {}
    if not isinstance(root, dict) or not source_signature:
        return {}
    workbook_overrides = root.get(source_signature) or {}
    if not isinstance(workbook_overrides, dict):
        return {}
    return {
        str(original): str(renamed).strip()
        for original, renamed in workbook_overrides.items()
        if str(original).strip() and str(renamed).strip()
    }


def set_l1_segment_name_overrides(
    source_signature: str,
    overrides: dict[str, str],
) -> None:
    """Persist detected-segment display-name overrides for one workbook."""
    root = dict(st.session_state.get("l1_segment_name_overrides") or {})
    cleaned = {
        str(original).strip(): re.sub(r"\s+", " ", str(renamed)).strip()
        for original, renamed in overrides.items()
        if str(original).strip() and str(renamed).strip()
    }
    if source_signature:
        if cleaned:
            root[source_signature] = cleaned
        else:
            root.pop(source_signature, None)
    st.session_state["l1_segment_name_overrides"] = root


def apply_l1_segment_name_overrides(
    profile: dict[str, Any],
    overrides: dict[str, str],
) -> dict[str, Any]:
    """Apply display-name overrides without changing membership, colors, or order."""
    result = dict(profile or {})
    if not result:
        return result

    raw_order = [
        str(value).strip()
        for value in (
            result.get("detected_segment_order")
            or result.get("segment_order")
            or []
        )
        if str(value).strip()
    ]
    valid_originals = set(raw_order)
    cleaned_overrides = {
        str(original).strip(): re.sub(r"\s+", " ", str(renamed)).strip()
        for original, renamed in (overrides or {}).items()
        if str(original).strip() in valid_originals and str(renamed).strip()
    }

    effective: dict[str, str] = {}
    used_names: set[str] = set()
    for original in raw_order:
        candidate = cleaned_overrides.get(original, original)
        key = candidate.casefold()
        if key in used_names or key in {"client", "unassigned"}:
            candidate = original
            key = candidate.casefold()
        used_names.add(key)
        effective[original] = candidate

    renamed_order = [effective.get(original, original) for original in raw_order]

    renamed_segments: list[dict[str, Any]] = []
    for item in result.get("segments") or []:
        if not isinstance(item, dict):
            continue
        copied = dict(item)
        original = str(
            copied.get("detected_segment")
            or copied.get("segment")
            or ""
        ).strip()
        copied["detected_segment"] = original
        copied["segment"] = effective.get(original, original)
        renamed_segments.append(copied)

    renamed_assignments: list[dict[str, Any]] = []
    for index, item in enumerate(result.get("assignments") or []):
        if not isinstance(item, dict):
            continue
        copied = dict(item)
        original = str(
            copied.get("detected_segment")
            or copied.get("segment")
            or ("Client" if index == 0 else "Unassigned")
        ).strip()
        copied["detected_segment"] = original
        if index == 0 or original.casefold() == "client":
            copied["segment"] = "Client"
        else:
            copied["segment"] = effective.get(original, original)
        renamed_assignments.append(copied)

    result["detected_segment_order"] = raw_order
    result["segment_order"] = renamed_order
    result["segment_name_overrides"] = {
        original: renamed
        for original, renamed in effective.items()
        if renamed != original
    }
    if renamed_segments:
        result["segments"] = renamed_segments
    if renamed_assignments:
        result["assignments"] = renamed_assignments
    result["segment_count"] = len(renamed_order)
    return result



def _segment_company_key(company: Any) -> str:
    """Return a stable workbook-local key for one company assignment."""
    key = re.sub(r"[^a-z0-9]+", "", str(company or "").strip().lower())
    return key


def get_l1_segment_assignment_overrides(source_signature: str) -> dict[str, str]:
    """Return persisted per-company Segment assignments for one workbook."""
    root = st.session_state.get("l1_segment_assignment_overrides") or {}
    if not isinstance(root, dict) or not source_signature:
        return {}
    workbook_overrides = root.get(source_signature) or {}
    if not isinstance(workbook_overrides, dict):
        return {}
    return {
        str(company_key): re.sub(r"\s+", " ", str(segment)).strip()
        for company_key, segment in workbook_overrides.items()
        if str(company_key).strip() and str(segment).strip()
    }


def set_l1_segment_assignment_overrides(
    source_signature: str,
    overrides: dict[str, str],
) -> None:
    """Persist peer-company Segment assignments for one workbook."""
    root = dict(st.session_state.get("l1_segment_assignment_overrides") or {})
    cleaned = {
        str(company_key).strip(): re.sub(r"\s+", " ", str(segment)).strip()
        for company_key, segment in (overrides or {}).items()
        if str(company_key).strip() and str(segment).strip()
    }
    if source_signature:
        if cleaned:
            root[source_signature] = cleaned
        else:
            root.pop(source_signature, None)
    st.session_state["l1_segment_assignment_overrides"] = root


def apply_l1_segment_assignment_overrides(
    profile: dict[str, Any],
    overrides: dict[str, str],
) -> dict[str, Any]:
    """Apply per-company Segment edits while preserving client-first behavior.

    The detected Segment row remains the audit source. Each peer can then be
    reassigned independently by company name. Matching edited labels form one
    output segment; duplicates are therefore intentional. The first company is
    always the client and cannot be reassigned.
    """
    result = dict(profile or {})
    if not result:
        return result

    raw_assignments = [
        dict(item)
        for item in (result.get("assignments") or [])
        if isinstance(item, dict) and str(item.get("company") or "").strip()
    ]
    if not raw_assignments:
        result["segment_assignment_overrides"] = {}
        return result

    raw_assignments.sort(key=lambda item: int(item.get("company_index") or 0))
    cleaned_overrides = {
        str(company_key).strip(): re.sub(r"\s+", " ", str(segment)).strip()
        for company_key, segment in (overrides or {}).items()
        if str(company_key).strip() and str(segment).strip()
    }

    detected_color_by_segment: dict[str, str] = {}
    detected_order: list[str] = []
    for item in raw_assignments[1:]:
        detected = str(
            item.get("detected_segment")
            or item.get("segment")
            or "Unassigned"
        ).strip() or "Unassigned"
        color = str(item.get("segment_color") or "").strip()
        if detected.casefold() not in {"client", "unassigned"}:
            if detected not in detected_order:
                detected_order.append(detected)
            if color and detected not in detected_color_by_segment:
                detected_color_by_segment[detected] = color

    effective_assignments: list[dict[str, Any]] = []
    manual_by_company: dict[str, str] = {}
    for position, item in enumerate(raw_assignments):
        copied = dict(item)
        company_index = int(copied.get("company_index") or position)
        company = str(copied.get("company") or "").strip()
        detected = str(
            copied.get("detected_segment")
            or copied.get("segment")
            or ("Client" if company_index == 0 else "Unassigned")
        ).strip() or ("Client" if company_index == 0 else "Unassigned")
        copied["company_index"] = company_index
        copied["company"] = company
        copied["detected_segment"] = detected

        if company_index == 0:
            effective = "Client"
            copied["assignment_source"] = "fixed_client"
        else:
            company_key = _segment_company_key(company)
            requested = cleaned_overrides.get(company_key, "").strip()
            if requested and requested.casefold() not in {"client", "unassigned"}:
                effective = requested
                copied["assignment_source"] = "manual"
                if effective != detected:
                    manual_by_company[company] = effective
            else:
                effective = detected
                copied["assignment_source"] = "detected"

        copied["segment"] = effective
        effective_assignments.append(copied)

    # Use the detected color when an edited label points to an existing segment.
    # A newly created label inherits the first moved company's original color.
    effective_color_by_segment: dict[str, str] = {}
    for item in effective_assignments[1:]:
        segment = str(item.get("segment") or "Unassigned").strip() or "Unassigned"
        detected = str(item.get("detected_segment") or "").strip()
        color = detected_color_by_segment.get(segment) or str(
            item.get("segment_color") or detected_color_by_segment.get(detected) or ""
        ).strip()
        if segment not in effective_color_by_segment and color:
            effective_color_by_segment[segment] = color
    for item in effective_assignments[1:]:
        segment = str(item.get("segment") or "Unassigned").strip() or "Unassigned"
        item["segment_color"] = effective_color_by_segment.get(
            segment,
            str(item.get("segment_color") or ""),
        )

    segment_order: list[str] = []
    for item in effective_assignments[1:]:
        segment = str(item.get("segment") or "Unassigned").strip() or "Unassigned"
        if segment.casefold() in {"client", "unassigned"}:
            continue
        if segment not in segment_order:
            segment_order.append(segment)

    # Rebuild contiguous source-column spans so the profile stays internally
    # consistent even when one company is moved out of a detected segment.
    rebuilt_segments: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for item in effective_assignments[1:]:
        segment = str(item.get("segment") or "Unassigned").strip() or "Unassigned"
        try:
            source_column = int(item.get("source_column") or 0)
        except Exception:
            source_column = 0
        color = str(item.get("segment_color") or "").strip()
        manual = str(item.get("assignment_source") or "") == "manual"
        if (
            current is not None
            and current.get("segment") == segment
            and source_column
            and int(current.get("end_column") or 0) + 1 == source_column
        ):
            current["end_column"] = source_column
            current["manual_override"] = bool(current.get("manual_override")) or manual
            continue
        current = {
            "segment": segment,
            "start_column": source_column,
            "end_column": source_column,
            "segment_color": color,
            "detection_method": "manual_assignment" if manual else "detected_assignment",
            "manual_override": manual,
        }
        rebuilt_segments.append(current)

    unassigned = [
        str(item.get("company") or "").strip()
        for item in effective_assignments[1:]
        if str(item.get("segment") or "").strip().casefold()
        in {"", "client", "unassigned"}
    ]
    detected = bool(effective_assignments[1:] and segment_order and not unassigned)

    result["assignments"] = effective_assignments
    result["segments"] = rebuilt_segments
    result["detected_segment_order"] = list(
        result.get("detected_segment_order") or detected_order
    )
    result["segment_order"] = segment_order
    result["segment_count"] = len(segment_order)
    result["company_count"] = len(effective_assignments)
    result["segment_assignment_overrides"] = manual_by_company
    # Kept for older manifest readers; category-wide rename is no longer used.
    result["segment_name_overrides"] = {}
    result["detected"] = detected
    result["error"] = (
        ""
        if detected
        else (
            "Assign a segment to: " + ", ".join(unassigned)
            if unassigned
            else "No peer segments are available."
        )
    )
    return result

def _current_required_input_keys() -> tuple[str, ...]:
    sections = normalize_deck_sections(st.session_state.get("workflow_mode") or ())
    keys = standardized_input_keys_for_sections(sections)
    return keys or STANDARD_INPUT_KEYS


def active_l1_input_key(source_signature: str) -> str:
    """Return the active standardized input across Aura and Geo Cost."""
    active_root = dict(st.session_state.get("active_input_by_section") or {})
    state_key = f"{source_signature}:inputs"
    legacy_keys = (f"{source_signature}:aura", f"{source_signature}:l1")
    active = str(
        active_root.get(state_key)
        or next((active_root.get(key) for key in legacy_keys if active_root.get(key)), "")
        or ""
    )
    required_keys = _current_required_input_keys()
    if active in required_keys:
        active_root[state_key] = active
        st.session_state["active_input_by_section"] = active_root
        return active

    mappings = get_standard_input_mappings(source_signature)
    active = next(
        (key for key in required_keys if key not in mappings),
        required_keys[0],
    )
    active_root[state_key] = active
    st.session_state["active_input_by_section"] = active_root
    return active


def set_active_l1_input_key(source_signature: str, input_key: str) -> None:
    """Make one standardized Aura/Geo Cost field the range-assignment target."""
    if input_key not in STANDARD_INPUT_KEYS:
        raise ValueError(f"Unsupported standardized input key: {input_key!r}")
    active_root = dict(st.session_state.get("active_input_by_section") or {})
    active_root[f"{source_signature}:inputs"] = input_key
    # Keep the legacy key synchronized while older canvas paths remain.
    active_root[f"{source_signature}:aura"] = input_key
    st.session_state["active_input_by_section"] = active_root


def _l1_mapping_range_signature(mappings: dict[str, dict[str, Any]]) -> str:
    items = sorted(
        (
            key,
            str((mappings.get(key) or {}).get("sheet_name") or ""),
            str((mappings.get(key) or {}).get("range_a1") or ""),
        )
        for key in L1_INPUT_KEYS
        if mappings.get(key)
    )
    return hashlib.sha256(repr(items).encode("utf-8")).hexdigest()


def _selection_plain_row_length(
    selection: dict[str, Any] | None,
    label: str,
    errors: list[str],
) -> int | None:
    """Return the number of values in a required value-only horizontal row."""
    if not selection:
        return None
    try:
        rows = int(selection.get("row_count") or 0)
        columns = int(selection.get("column_count") or 0)
    except (TypeError, ValueError):
        errors.append(f"{label} has an invalid range shape.")
        return None
    if rows != 1 or columns < 1:
        errors.append(
            f"{label} must be one horizontal row whose first selected cell is the "
            f"client company, not {rows} × {columns}."
        )
        return None
    return columns


def _selection_labeled_row_length(
    selection: dict[str, Any] | None,
    label: str,
    errors: list[str],
) -> int | None:
    """Return the value count after the required left-hand master label."""
    if not selection:
        return None
    try:
        rows = int(selection.get("row_count") or 0)
        columns = int(selection.get("column_count") or 0)
    except (TypeError, ValueError):
        errors.append(f"{label} has an invalid range shape.")
        return None
    if rows != 1 or columns < 2:
        errors.append(
            f"{label} must be one horizontal row containing a master label in the "
            f"first cell and data to its right, not {rows} × {columns}."
        )
        return None
    return columns - 1


def _selection_labeled_column_length(
    selection: dict[str, Any] | None,
    label: str,
    errors: list[str],
) -> int | None:
    """Return the value count after the required top master label/header."""
    if not selection:
        return None
    try:
        rows = int(selection.get("row_count") or 0)
        columns = int(selection.get("column_count") or 0)
    except (TypeError, ValueError):
        errors.append(f"{label} has an invalid range shape.")
        return None
    if columns != 1 or rows < 2:
        errors.append(
            f"{label} must be one vertical column containing its header in the first "
            f"cell and values beneath it, not {rows} × {columns}."
        )
        return None
    return rows - 1


def _selection_quartile_table_shape(
    selection: dict[str, Any] | None,
    label: str,
    errors: list[str],
) -> tuple[int, int] | None:
    """Validate one mapped three-column quartile table range."""
    if not selection:
        return None
    try:
        rows = int(selection.get("row_count") or 0)
        columns = int(selection.get("column_count") or 0)
    except (TypeError, ValueError):
        errors.append(f"{label} has an invalid range shape.")
        return None
    if columns != 3 or rows < 1:
        errors.append(
            f"{label} must contain exactly three consecutive quartile columns; "
            f"the selected range is {rows} × {columns}."
        )
        return None
    return rows, columns


def validate_l1_input_mappings(
    mappings: dict[str, dict[str, Any]],
    *,
    source_signature: str = "",
    taxonomy_profile: dict[str, Any] | None = None,
    function_lookup_profile: dict[str, Any] | None = None,
    master_label_profile: dict[str, Any] | None = None,
    required_keys: tuple[str, ...] | None = None,
    ignored_taxonomy_functions: Iterable[str] | None = None,
) -> dict[str, Any]:
    """Validate the standardized Aura inputs and source-row lookup.

    Company names are a value-only horizontal row. Taxonomy, Revenue, FTE, and
    selected benchmark inputs are complete three-column quartile tables.
    Taxonomy controls slide order and final L0/L1 labels. ``required_keys``
    lets each output request the shared inputs plus its own quartile table.
    """
    requested_keys = AURA_INPUT_KEYS if required_keys is None else required_keys
    required = tuple(
        key for key in requested_keys
        if key in AURA_INPUT_KEYS
    )
    mappings = {
        key: value
        for key, value in mappings.items()
        if key in required and isinstance(value, dict)
    }
    missing = [key for key in required if not mappings.get(key)]
    errors: list[str] = []
    warnings: list[str] = []
    ignored_lookup_keys = {
        _taxonomy_mismatch_key(value)
        for value in (ignored_taxonomy_functions or ())
        if _taxonomy_mismatch_key(value)
    }

    for key, selection in mappings.items():
        if key not in L1_INPUT_KEYS or not isinstance(selection, dict):
            continue
        selection_signature = str(selection.get("source_signature") or "")
        if (
            source_signature
            and selection_signature
            and selection_signature != source_signature
        ):
            errors.append(
                f"{L1_INPUT_BY_KEY[key]['label']} belongs to a different workbook upload."
            )

    taxonomy_selection = mappings.get("taxonomy")
    taxonomy_rows = 0
    taxonomy_columns = 0
    if taxonomy_selection:
        try:
            taxonomy_rows = int(taxonomy_selection.get("row_count") or 0)
            taxonomy_columns = int(taxonomy_selection.get("column_count") or 0)
        except (TypeError, ValueError):
            taxonomy_rows = taxonomy_columns = 0
        if taxonomy_columns != 2:
            errors.append(
                f"Taxonomy must be a two-column table; the selected range has "
                f"{taxonomy_columns} columns."
            )
        if taxonomy_rows < 2:
            errors.append(
                "Taxonomy must include its header row and at least one L0/L1 data row."
            )

    profile = dict(
        taxonomy_profile
        if isinstance(taxonomy_profile, dict)
        else (st.session_state.get("taxonomy_sheet_profile") or {})
    )
    taxonomy_sheet = str(profile.get("sheet_name") or "").strip()
    taxonomy_error = str(profile.get("error") or "").strip()
    try:
        taxonomy_count = int(profile.get("row_count") or 0)
    except (TypeError, ValueError):
        taxonomy_count = 0

    if taxonomy_selection:
        selected_sheet = str(taxonomy_selection.get("sheet_name") or "").strip()
        if taxonomy_sheet and selected_sheet and taxonomy_sheet != selected_sheet:
            errors.append(
                "Taxonomy validation is out of sync with the selected Taxonomy range."
            )
        if taxonomy_error:
            errors.append(taxonomy_error)
        elif not taxonomy_sheet:
            errors.append("The selected Taxonomy range has no worksheet name.")
        elif taxonomy_count < 1:
            errors.append(
                f"Taxonomy range on {taxonomy_sheet!r} contains no L1 rows beneath "
                "its detected L0/L1 header pair."
            )

    company_count = _selection_plain_row_length(
        mappings.get("company_headers"), "Company names", errors
    )
    revenue_count = _selection_labeled_row_length(
        mappings.get("revenue_values"), "Revenue row", errors
    )
    l0_quartile_shape = _selection_quartile_table_shape(
        mappings.get("l0_quartile_table"),
        "L0 quartile table detector",
        errors,
    )
    l1_quartile_shape = _selection_quartile_table_shape(
        mappings.get("l1_quartile_table"),
        "L1 quartile table detector",
        errors,
    )

    if company_count is not None and company_count < 2:
        errors.append(
            "Company names must contain the client plus at least one peer company."
        )
    if (
        company_count is not None
        and revenue_count is not None
        and revenue_count != company_count
    ):
        errors.append(
            f"Revenue values contains {revenue_count} cells; Company names contains {company_count}."
        )

    matrix = mappings.get("fte_values")
    matrix_rows = None
    matrix_columns = None
    matrix_descriptor_columns = None
    matrix_value_columns = None
    if matrix:
        try:
            matrix_rows = int(matrix.get("row_count") or 0)
            matrix_columns = int(matrix.get("column_count") or 0)
        except (TypeError, ValueError):
            matrix_rows = matrix_columns = 0
        if matrix_rows < 2:
            errors.append(
                "Function + FTE / $B table must include its header row and at least "
                "one source-function data row."
            )
        if company_count is not None:
            matrix_value_columns = company_count
            matrix_descriptor_columns = matrix_columns - company_count
            if matrix_descriptor_columns != 1:
                errors.append(
                    "Function + FTE / $B table must start with exactly one L1/function "
                    f"column followed by {company_count} company-value columns; the selected "
                    f"range has {matrix_descriptor_columns} leading column(s). Do not include L0."
                )

    matrix_data_rows = (matrix_rows - 1) if matrix_rows is not None and matrix_rows > 0 else None

    # The combined L1 quartile table aligns to the FTE source rows beneath
    # its header. Taxonomy rows are subsequently matched by L1 name.

    labels = dict(
        master_label_profile
        if isinstance(master_label_profile, dict)
        else (st.session_state.get("l1_master_label_profile") or {})
    )
    label_errors = [
        str(value).strip()
        for value in (labels.get("errors") or [])
        if str(value).strip()
    ]
    profile_signature = str(labels.get("mapping_signature") or "")
    current_signature = _l1_mapping_range_signature(mappings)
    labels_are_current = not profile_signature or profile_signature == current_signature
    if bool(labels.get("attempted")) and labels_are_current:
        errors.extend(label_errors)

    quartile_profiles = dict(labels.get("quartile_tables") or {})
    l0_quartile_profile = dict(quartile_profiles.get("l0_quartile_table") or {})
    l1_quartile_profile = dict(quartile_profiles.get("l1_quartile_table") or {})
    try:
        l0_quartile_data_rows = int(l0_quartile_profile.get("data_row_count") or 0)
    except (TypeError, ValueError):
        l0_quartile_data_rows = 0
    try:
        l1_quartile_data_rows = int(l1_quartile_profile.get("data_row_count") or 0)
    except (TypeError, ValueError):
        l1_quartile_data_rows = 0

    if mappings.get("l0_quartile_table") and labels_are_current:
        if l0_quartile_data_rows < 1 and not l0_quartile_profile.get("error"):
            errors.append(
                "L0 quartile table detector contains no benchmark rows that can be matched to an L0 label."
            )

    if (
        mappings.get("l1_quartile_table")
        and labels_are_current
        and matrix_data_rows is not None
        and matrix_data_rows > 0
        and l1_quartile_data_rows > 0
        and l1_quartile_data_rows != matrix_data_rows
    ):
        errors.append(
            f"L1 quartile table contains {l1_quartile_data_rows} benchmark rows; "
            f"Function + FTE / $B contains {matrix_data_rows} source data rows beneath its header."
        )

    lookup = dict(
        function_lookup_profile
        if isinstance(function_lookup_profile, dict)
        else (st.session_state.get("l1_function_lookup_profile") or {})
    )
    lookup_attempted = bool(lookup.get("attempted"))
    lookup_error = str(lookup.get("error") or "").strip()
    lookup_missing = [str(value) for value in (lookup.get("missing") or []) if str(value)]
    lookup_ambiguous = [str(value) for value in (lookup.get("ambiguous") or []) if str(value)]
    try:
        matched_function_count = int(lookup.get("matched_count") or 0)
    except (TypeError, ValueError):
        matched_function_count = 0
    try:
        source_function_count = int(lookup.get("source_row_count") or 0)
    except (TypeError, ValueError):
        source_function_count = 0

    # Only enforce the lookup after all ranges needed to calculate it exist.
    lookup_inputs_present = all(
        mappings.get(key)
        for key in ("taxonomy", "fte_values")
    )
    ignored_missing = [
        value for value in lookup_missing
        if _taxonomy_mismatch_key(value) in ignored_lookup_keys
    ]
    active_missing = [
        value for value in lookup_missing
        if _taxonomy_mismatch_key(value) not in ignored_lookup_keys
    ]
    if lookup_inputs_present and lookup_attempted:
        if lookup_error:
            errors.append(lookup_error)
        elif lookup_ambiguous:
            errors.append(
                "Multiple FTE source rows match Taxonomy function(s): "
                + ", ".join(lookup_ambiguous[:6])
                + ("…" if len(lookup_ambiguous) > 6 else "")
            )
        elif active_missing:
            errors.append(
                "No FTE source row was found for Taxonomy function(s): "
                + ", ".join(active_missing[:6])
                + ("…" if len(active_missing) > 6 else "")
            )
        else:
            expected_matches = max(0, taxonomy_count - len(ignored_missing))
            if taxonomy_count > 0 and matched_function_count != expected_matches:
                errors.append(
                    f"Matched {matched_function_count} of {expected_matches} non-ignored Taxonomy functions "
                    "to FTE source rows."
                )
    # Explicitly ignored missing Taxonomy functions are part of the effective
    # taxonomy state, not a validation warning. They are excluded from the
    # Taxonomy→FTE crosswalk and from Aura slide generation. Keep them in the
    # result as informational metadata so the UI can explain/re-enable them,
    # while allowing Taxonomy/FTE to remain green when everything else is valid.

    result = {
        "ready": not missing and not errors,
        "mapped_count": sum(1 for key in required if mappings.get(key)),
        "required_count": len(required),
        "required_keys": list(required),
        "missing": missing,
        "errors": errors,
        "warnings": warnings,
        "ignored_taxonomy_functions": ignored_missing,
        "company_count": company_count,
        "function_count": max(0, taxonomy_count - len(ignored_missing)) or None,
        "taxonomy_sheet": taxonomy_sheet,
        "taxonomy_range": str(
            (taxonomy_selection or {}).get("range_a1") or ""
        ),
        "taxonomy_header_row": profile.get("header_row"),
        "taxonomy_l0_column": profile.get("l0_column"),
        "taxonomy_l1_column": profile.get("l1_column"),
        "taxonomy_header_prefix": profile.get("header_prefix"),
        "matrix_rows": matrix_rows,
        "matrix_data_rows": matrix_data_rows,
        "matrix_columns": matrix_columns,
        "matrix_descriptor_columns": matrix_descriptor_columns,
        "matrix_value_columns": matrix_value_columns,
        "l0_quartile_rows": l0_quartile_data_rows,
        "l1_quartile_rows": l1_quartile_data_rows,
        "l0_quartile_shape": l0_quartile_shape,
        "l1_quartile_shape": l1_quartile_shape,
        "matched_function_count": matched_function_count,
        "source_function_count": source_function_count,
        "function_label_range": str(lookup.get("function_label_range") or ""),
        "master_labels": dict(labels.get("labels") or {}),
    }
    st.session_state["l1_mapping_validation"] = result
    return result


def validate_geocost_input_mappings(
    mappings: dict[str, dict[str, Any]],
    *,
    source_signature: str = "",
    profile: dict[str, Any] | None = None,
    required_keys: tuple[str, ...] | None = None,
) -> dict[str, Any]:
    """Validate the two Geo Cost wrapper ranges.

    ``Geo Cost table`` establishes the visible geography/company matrix.
    ``Geo Cost data row`` is the low-cost FTE percentage row used at the bottom
    of the chart. Helper columns and repeated >>> function markers are inferred.
    """
    required = tuple(
        key
        for key in (GEOCOST_INPUT_KEYS if required_keys is None else required_keys)
        if key in GEOCOST_INPUT_KEYS
    )
    selected = {
        key: value
        for key, value in mappings.items()
        if key in required and isinstance(value, dict)
    }
    missing = [key for key in required if not selected.get(key)]
    errors: list[str] = []

    for key, selection in selected.items():
        selection_signature = str(selection.get("source_signature") or "")
        if (
            source_signature
            and selection_signature
            and selection_signature != source_signature
        ):
            errors.append(
                f"{GEOCOST_INPUT_BY_KEY[key]['label']} belongs to a different workbook upload."
            )

    table = selected.get("geocost_table")
    data_row = selected.get("geocost_data_row")
    table_rows = table_columns = 0
    marker_rows = marker_columns = 0
    if table:
        try:
            table_rows = int(table.get("row_count") or 0)
            table_columns = int(table.get("column_count") or 0)
        except (TypeError, ValueError):
            pass
        if table_rows < 2 or table_columns < 2:
            errors.append(
                "Geo Cost table must include the company header row, at least one geography "
                "row, the local Country/State label column, and at least one company column."
            )
    if data_row:
        try:
            marker_rows = int(data_row.get("row_count") or 0)
            marker_columns = int(data_row.get("column_count") or 0)
        except (TypeError, ValueError):
            pass
        if marker_rows != 1 or marker_columns < 2:
            errors.append(
                "Geo Cost data row must be one horizontal low-cost percentage row containing "
                "its label and at least one company value."
            )

    if table and data_row:
        table_sheet = str(table.get("sheet_name") or "").strip()
        marker_sheet = str(data_row.get("sheet_name") or "").strip()
        if not table_sheet or not marker_sheet or table_sheet.casefold() != marker_sheet.casefold():
            errors.append("Geo Cost table and Geo Cost data row must come from the same worksheet.")
        try:
            if int(data_row.get("start_row") or 0) <= int(table.get("end_row") or 0):
                errors.append(
                    "Geo Cost data row must be below the final geography row of the Geo Cost table."
                )
        except (TypeError, ValueError):
            pass

    details = dict(
        profile
        if isinstance(profile, dict)
        else (st.session_state.get("geocost_input_profile") or {})
    )
    profile_error = str(details.get("error") or "").strip()
    profile_signature = str(details.get("source_signature") or "")
    if profile_signature and source_signature and profile_signature != source_signature:
        profile_error = "Geo Cost input validation is out of sync with the current workbook."
    if table and data_row and profile_error:
        errors.append(profile_error)
    elif table and data_row and details and not bool(details.get("ready")):
        errors.append("The selected Geo Cost ranges do not form a valid repeated-block source layout.")

    result = {
        "ready": not missing and not errors,
        "mapped_count": sum(1 for key in required if selected.get(key)),
        "required_count": len(required),
        "required_keys": list(required),
        "missing": missing,
        "errors": errors,
        "sheet_name": str(details.get("sheet_name") or ""),
        "table_range": str((table or {}).get("range_a1") or ""),
        "data_row_range": str((data_row or {}).get("range_a1") or ""),
        "function_count": int(details.get("function_count") or 0),
        "data_row_count": int(details.get("data_row_count") or 0),
        "layout": dict(details.get("layout") or {}),
    }
    st.session_state["geocost_mapping_validation"] = result
    return result


def validate_standard_input_mappings(
    mappings: dict[str, dict[str, Any]],
    *,
    selected_sections: tuple[str, ...],
    source_signature: str = "",
    taxonomy_profile: dict[str, Any] | None = None,
    function_lookup_profile: dict[str, Any] | None = None,
    master_label_profile: dict[str, Any] | None = None,
    geocost_profile: dict[str, Any] | None = None,
    ignored_taxonomy_functions: Iterable[str] | None = None,
) -> dict[str, Any]:
    """Combine Aura and Geo Cost validation for the single Inputs rail."""
    aura_keys = aura_input_keys_for_sections(selected_sections)
    geocost_keys = geocost_input_keys_for_sections(selected_sections)
    results: list[dict[str, Any]] = []
    if aura_keys:
        results.append(
            validate_l1_input_mappings(
                mappings,
                source_signature=source_signature,
                taxonomy_profile=taxonomy_profile,
                function_lookup_profile=function_lookup_profile,
                master_label_profile=master_label_profile,
                required_keys=aura_keys,
                ignored_taxonomy_functions=ignored_taxonomy_functions,
            )
        )
    if geocost_keys:
        results.append(
            validate_geocost_input_mappings(
                mappings,
                source_signature=source_signature,
                profile=geocost_profile,
                required_keys=geocost_keys,
            )
        )
    required = standardized_input_keys_for_sections(selected_sections)
    missing: list[str] = []
    errors: list[str] = []
    warnings: list[str] = []
    for result in results:
        missing.extend(str(key) for key in result.get("missing") or [])
        errors.extend(str(value) for value in result.get("errors") or [] if str(value))
        warnings.extend(str(value) for value in result.get("warnings") or [] if str(value))
    return {
        "ready": bool(required) and all(bool(result.get("ready")) for result in results),
        "mapped_count": sum(1 for key in required if mappings.get(key)),
        "required_count": len(required),
        "required_keys": list(required),
        "missing": missing,
        "errors": errors,
        "warnings": warnings,
        "aura": results[0] if aura_keys and results else {},
        "geocost": results[-1] if geocost_keys and results else {},
    }


def _validation_errors_for_input(key: str, errors: list[str]) -> list[str]:
    """Return validation messages that most directly relate to one Inputs-rail field."""
    normalized_key = str(key or "").strip()
    needles = {
        "taxonomy": (
            "taxonomy",
            "l0/l1 header",
            "matched ",
            "no fte source row",
            "multiple fte source rows",
        ),
        "company_headers": (
            "company names",
            "client plus at least one peer",
        ),
        "revenue_values": (
            "revenue row",
            "revenue values",
        ),
        "fte_values": (
            "function + fte",
            "fte source",
            "source-function",
            "source data row",
            "leading column",
            "no fte source row",
            "multiple fte source rows",
            "matched ",
        ),
        "l0_quartile_table": (
            "l0 quartile",
        ),
        "l1_quartile_table": (
            "l1 quartile",
        ),
        "geocost_table": (
            "geo cost table",
            "repeated-block source layout",
            "geo cost input validation",
        ),
        "geocost_data_row": (
            "geo cost data row",
            "repeated-block source layout",
            "geo cost input validation",
        ),
    }.get(normalized_key, ())
    matched: list[str] = []
    for message in errors:
        text = str(message or "").strip()
        lower = text.casefold()
        if text and any(needle in lower for needle in needles):
            matched.append(text)
    return matched



def _compact_input_validation_message(message: str, *, limit: int = 280) -> str:
    """Return a concise, user-facing validation message for the status strip."""
    text = " ".join(str(message or "").split()).strip()
    prefix = "No FTE source row was found for Taxonomy function(s):"
    if text.casefold().startswith(prefix.casefold()):
        raw_names = text[len(prefix):].strip()
        names = [name.strip() for name in raw_names.split(",") if name.strip()]
        if names:
            preview = ", ".join(names[:5])
            suffix = "" if len(names) <= 5 else f", +{len(names) - 5} more"
            return (
                f"{len(names)} Taxonomy function{'s' if len(names) != 1 else ''} "
                f"have no matching FTE source row: {preview}{suffix}."
            )
    if len(text) > limit:
        return text[: max(0, limit - 1)].rstrip() + "…"
    return text


def standard_input_status(
    *,
    field: dict[str, str],
    mapping: dict[str, Any] | None,
    workbook_ready: bool,
    validation: dict[str, Any],
    auto_error: str = "",
) -> dict[str, str]:
    """Return the visible status for one standardized Inputs-rail field.

    Status is intentionally field-local. An input that passes its own checks
    remains green even when Finalize is blocked by a different input. Yellow is
    reserved for mapped inputs that cannot yet be fully cross-validated because
    a companion input is still missing.
    """
    key = str(field.get("key") or "")
    label = str(field.get("label") or key or "Input")
    guidance = str(field.get("help") or "").strip()
    errors = [
        str(value).strip()
        for value in (validation.get("errors") or [])
        if str(value).strip()
    ]
    warnings = [
        str(value).strip()
        for value in (validation.get("warnings") or [])
        if str(value).strip()
    ]
    missing = {str(value) for value in (validation.get("missing") or []) if str(value)}

    if not workbook_ready:
        return {
            "level": "error",
            "symbol": "❗",
            "message": f"{label}: upload a workbook before this input can be validated.",
        }

    if not mapping:
        message = f"{label} has not been assigned yet."
        if auto_error:
            message += " Auto-detect issue: " + _compact_input_validation_message(auto_error, limit=210)
        elif guidance:
            message += " Expected selection: " + guidance
        return {"level": "error", "symbol": "❗", "message": message}

    local_errors = _validation_errors_for_input(key, errors)
    if local_errors:
        return {
            "level": "error",
            "symbol": "❗",
            "message": _compact_input_validation_message(local_errors[0]),
        }

    local_warnings = _validation_errors_for_input(key, warnings)
    if local_warnings:
        return {
            "level": "warning",
            "symbol": "⚠️",
            "message": _compact_input_validation_message(local_warnings[0]),
        }

    dependencies = {
        "taxonomy": ("fte_values",),
        "fte_values": ("taxonomy",),
        "revenue_values": ("company_headers",),
        "geocost_table": ("geocost_data_row",),
        "geocost_data_row": ("geocost_table",),
    }.get(key, ())
    missing_dependencies = [dep for dep in dependencies if dep in missing]
    if missing_dependencies:
        names = [
            STANDARD_INPUT_BY_KEY.get(dep, {}).get("label", dep)
            for dep in missing_dependencies
        ]
        return {
            "level": "warning",
            "symbol": "⚠️",
            "message": (
                f"{label} is mapped, but it cannot be fully cross-validated until "
                + ", ".join(names)
                + " is mapped."
            ),
        }

    sheet_name = str(mapping.get("sheet_name") or "").strip()
    range_a1 = str(mapping.get("range_a1") or "").strip()
    location = "!".join(part for part in (sheet_name, range_a1) if part)
    ignored_names = [
        " ".join(str(value or "").split()).strip()
        for value in (validation.get("ignored_taxonomy_functions") or [])
        if " ".join(str(value or "").split()).strip()
    ]
    message = f"{label} appears valid" + (f" at {location}." if location else ".")
    if key in {"taxonomy", "fte_values"} and ignored_names:
        message += (
            " Excluded unmatched Taxonomy function"
            + ("s" if len(ignored_names) != 1 else "")
            + ": "
            + ", ".join(ignored_names[:6])
            + ("…" if len(ignored_names) > 6 else "")
            + ". No slide will be generated for the excluded function"
            + ("s" if len(ignored_names) != 1 else "")
            + "."
        )
    return {
        "level": "ok",
        "symbol": "✅",
        "message": message,
    }

def _standard_input_button_help(
    *,
    field: dict[str, str],
    mapping: dict[str, Any] | None,
    workbook_ready: bool,
    validation: dict[str, Any],
    auto_error: str = "",
) -> str:
    """Build a useful hover tooltip for an Inputs-rail button.

    The tooltip intentionally reports both the field-local state and, when this
    field itself is healthy, the first blocker that is preventing Finalize from
    becoming available elsewhere in the standardized input set.
    """
    key = str(field.get("key") or "")
    label = str(field.get("label") or key or "Input")
    guidance = str(field.get("help") or "").strip()
    if not workbook_ready:
        return f"{label}: upload a workbook before selecting or validating this input."

    errors = [str(value).strip() for value in (validation.get("errors") or []) if str(value).strip()]
    missing = [str(value) for value in (validation.get("missing") or []) if str(value)]
    local_errors = _validation_errors_for_input(key, errors)

    if not mapping:
        parts = [f"{label} is required but is not currently mapped."]
        if auto_error:
            parts.append(f"Auto-detect issue: {auto_error}")
        if guidance:
            parts.append(guidance)
        return "\n\n".join(parts)

    sheet_name = str(mapping.get("sheet_name") or "").strip()
    range_a1 = str(mapping.get("range_a1") or "").strip()
    location = "!".join(part for part in (sheet_name, range_a1) if part)
    origin = str(mapping.get("mapping_origin") or "manual").strip().upper()
    parts = [f"{label}: {origin} mapping" + (f" at {location}." if location else ".")]

    if local_errors:
        parts.append("Issue: " + " ".join(local_errors[:3]))
    elif bool(validation.get("ready")):
        parts.append("This input currently passes validation. Click to review or replace the range.")
    else:
        other_missing = [
            STANDARD_INPUT_BY_KEY[item]["label"]
            for item in missing
            if item != key and item in STANDARD_INPUT_BY_KEY
        ]
        if other_missing:
            parts.append("This input is mapped. Finalize is still blocked because these inputs are missing: " + ", ".join(other_missing) + ".")
        elif errors:
            parts.append("This input has no direct validation error. Finalize is currently blocked by: " + errors[0])
        else:
            parts.append("This input is mapped. Final validation is still in progress.")

    if guidance:
        parts.append("Expected selection: " + guidance)
    return "\n\n".join(parts)


def _inject_input_hover_tooltips(tooltips: dict[str, str]) -> None:
    """Install/update the persistent delayed diagnostic tooltip controller."""
    if not tooltips:
        return
    payload = json.dumps(
        {str(key): str(value or "").strip() for key, value in tooltips.items()},
        ensure_ascii=False,
    )
    markup = r"""
    <!doctype html>
    <html><head><style>html,body{margin:0;padding:0;width:0;height:0;overflow:hidden}</style></head>
    <body><script>
    (() => {
      const win = window.parent;
      const doc = win.document;
      const INPUT_MAP = __TC_TOOLTIP_MAP__;
      const DELAY = 2000;

      function ensureManager() {
        let state = win.__tcPersistentDiagnosticTooltip;
        if (!state) {
          state = {
            inputMap: {},
            finalizeText: '',
            timer: null,
            hoverTarget: null,
            hoverKind: '',
            bubbleId: 'tc-persistent-diagnostic-tooltip'
          };
          win.__tcPersistentDiagnosticTooltip = state;
        }

        function ensureBubble() {
          let bubble = doc.getElementById(state.bubbleId);
          if (bubble) return bubble;
          bubble = doc.createElement('div');
          bubble.id = state.bubbleId;
          bubble.setAttribute('role', 'tooltip');
          Object.assign(bubble.style, {
            position: 'fixed', zIndex: '2147483647', maxWidth: '440px', minWidth: '240px',
            padding: '13px 15px', borderRadius: '12px', background: '#15171a', color: '#ffffff',
            border: '1px solid rgba(255,255,255,.16)', boxShadow: '0 18px 42px rgba(0,0,0,.40)',
            fontFamily: 'Arial, Helvetica, sans-serif', fontSize: '12px', fontWeight: '650',
            lineHeight: '1.48', whiteSpace: 'pre-line', overflowWrap: 'anywhere', pointerEvents: 'none',
            opacity: '0', visibility: 'hidden', transform: 'translateY(4px)',
            transition: 'opacity 120ms ease, transform 120ms ease, visibility 120ms ease'
          });
          const accent = doc.createElement('div');
          Object.assign(accent.style, {
            position: 'absolute', left: '12px', right: '12px', top: '0', height: '3px',
            background: '#cc0000', borderRadius: '0 0 3px 3px'
          });
          const body = doc.createElement('div');
          body.className = 'tc-persistent-diagnostic-tooltip-body';
          const arrow = doc.createElement('div');
          arrow.className = 'tc-persistent-diagnostic-tooltip-arrow';
          Object.assign(arrow.style, {
            position: 'absolute', width: '10px', height: '10px', background: '#15171a',
            transform: 'rotate(45deg)', border: '1px solid rgba(255,255,255,.10)'
          });
          bubble.append(accent, body, arrow);
          doc.body.appendChild(bubble);
          return bubble;
        }

        state.hide = function() {
          if (state.timer) {
            win.clearTimeout(state.timer);
            state.timer = null;
          }
          state.hoverTarget = null;
          state.hoverKind = '';
          const bubble = doc.getElementById(state.bubbleId);
          if (!bubble) return;
          bubble.style.opacity = '0';
          bubble.style.visibility = 'hidden';
          bubble.style.transform = 'translateY(4px)';
        };

        state.resolve = function(node) {
          if (!node || !node.closest) return null;
          const card = node.closest('[class*="st-key-sidebar_l1_input_"]');
          if (card && !card.classList.contains('st-key-sidebar_l1_input_list')) {
            const cls = Array.from(card.classList).find(
              (name) => name.startsWith('st-key-sidebar_l1_input_') && name !== 'st-key-sidebar_l1_input_list'
            );
            const key = cls ? cls.slice('st-key-sidebar_l1_input_'.length) : '';
            const text = String((state.inputMap || {})[key] || '').trim();
            if (text) return { target: card, text, kind: 'input' };
          }
          const finalize = node.closest('.st-key-workbook_review_action');
          if (finalize) {
            const text = String(state.finalizeText || '').trim();
            if (text) return { target: finalize, text, kind: 'finalize' };
          }
          return null;
        };

        state.show = function(item) {
          if (!item || !item.target || !item.target.isConnected) return;
          const bubble = ensureBubble();
          const body = bubble.querySelector('.tc-persistent-diagnostic-tooltip-body');
          if (body) body.textContent = item.text;
          bubble.style.display = 'block';
          bubble.style.visibility = 'hidden';
          bubble.style.opacity = '0';

          const rect = item.target.getBoundingClientRect();
          const br = bubble.getBoundingClientRect();
          const gap = 12, pad = 10;
          let left, top, side;
          if (item.kind === 'finalize') {
            left = rect.right - br.width;
            top = rect.bottom + gap;
            side = 'top';
            if (top + br.height > win.innerHeight - pad) {
              top = rect.top - br.height - gap;
              side = 'bottom';
            }
          } else {
            left = rect.right + gap;
            top = rect.top + (rect.height - br.height) / 2;
            side = 'left';
            if (left + br.width > win.innerWidth - pad) {
              left = rect.left - br.width - gap;
              side = 'right';
            }
          }
          left = Math.max(pad, Math.min(left, win.innerWidth - br.width - pad));
          top = Math.max(pad, Math.min(top, win.innerHeight - br.height - pad));
          bubble.style.left = Math.round(left) + 'px';
          bubble.style.top = Math.round(top) + 'px';

          const arrow = bubble.querySelector('.tc-persistent-diagnostic-tooltip-arrow');
          if (arrow) {
            arrow.style.left = ''; arrow.style.right = ''; arrow.style.top = ''; arrow.style.bottom = '';
            if (side === 'left') {
              arrow.style.left = '-6px';
              arrow.style.top = Math.max(14, Math.min(br.height - 24, rect.top + rect.height / 2 - top - 5)) + 'px';
            } else if (side === 'right') {
              arrow.style.right = '-6px';
              arrow.style.top = Math.max(14, Math.min(br.height - 24, rect.top + rect.height / 2 - top - 5)) + 'px';
            } else if (side === 'top') {
              arrow.style.top = '-6px'; arrow.style.right = '24px';
            } else {
              arrow.style.bottom = '-6px'; arrow.style.right = '24px';
            }
          }
          bubble.style.visibility = 'visible';
          bubble.style.opacity = '1';
          bubble.style.transform = 'translateY(0)';
        };

        if (!state.listenersInstalled) {
          state.listenersInstalled = true;
          doc.addEventListener('mouseover', (event) => {
            const item = state.resolve(event.target);
            if (!item) return;
            if (state.hoverTarget === item.target && state.hoverKind === item.kind) return;
            state.hide();
            state.hoverTarget = item.target;
            state.hoverKind = item.kind;
            state.timer = win.setTimeout(() => {
              state.timer = null;
              const current = state.resolve(item.target);
              if (current && current.target === item.target) state.show(current);
            }, DELAY);
          }, true);

          doc.addEventListener('mouseout', (event) => {
            if (!state.hoverTarget) return;
            const related = event.relatedTarget;
            if (related && state.hoverTarget.contains && state.hoverTarget.contains(related)) return;
            const leaving = event.target && event.target.closest
              ? event.target.closest('[class*="st-key-sidebar_l1_input_"], .st-key-workbook_review_action')
              : null;
            if (leaving === state.hoverTarget || !related) state.hide();
          }, true);
          doc.addEventListener('mousedown', () => state.hide(), true);
          win.addEventListener('scroll', () => state.hide(), true);
          win.addEventListener('resize', () => state.hide(), true);
        }
        return state;
      }

      const manager = ensureManager();
      manager.inputMap = Object.assign({}, INPUT_MAP || {});
    })();
    </script></body></html>
    """.replace('__TC_TOOLTIP_MAP__', payload)
    components.html(markup, height=0, scrolling=False)


WorkbookInspector = Callable[[Path], UploadInspection]


# -----------------------------------------------------------------------------
# Paths, upload persistence, and validation
# -----------------------------------------------------------------------------


def windows_downloads_folder(repository_root: Path) -> Path:
    """Return the current Windows user's Downloads folder."""
    candidates: list[Path] = []

    if os.name == "nt":
        try:
            import winreg

            registry_path = (
                r"Software\Microsoft\Windows\CurrentVersion\Explorer"
                r"\User Shell Folders"
            )
            downloads_guid = "{374DE290-123F-4565-9164-39C4925E467B}"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
                raw_value, _ = winreg.QueryValueEx(key, downloads_guid)
            candidates.append(Path(os.path.expandvars(str(raw_value))).expanduser())
        except Exception:
            pass

    candidates.extend(
        [
            Path.home() / "Downloads",
            Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Downloads",
        ]
    )

    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            return candidate.resolve()
        except OSError:
            continue

    fallback = repository_root / "UI_Output"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback.resolve()


def runtime_folder() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir())
    folder = base / "ThinkCellSlideMaker" / "StreamlitUI"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def safe_filename(value: str, fallback: str) -> str:
    name = Path(str(value or "")).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name or fallback


def normalize_output_name(value: str) -> str:
    name = safe_filename(value, "Generated_ThinkCell_Deck.pptx")
    if Path(name).suffix.lower() not in PRESENTATION_SUFFIXES:
        name = f"{Path(name).stem or 'Generated_ThinkCell_Deck'}.pptx"
    return name


def validate_workbook_package(path: Path) -> None:
    """Perform a fast structural validation of an OOXML workbook.

    ``ZipFile.testzip()`` expands every member in the workbook. Large client
    workbooks can contain sizeable cached objects, images, or model parts, so a
    full CRC sweep blocks the Streamlit script for minutes before the preview
    can render. The slide readers and openpyxl will still validate the specific
    worksheet parts they consume. At upload time we only need to confirm that
    this is a readable Excel OOXML package with the required workbook parts.
    """
    if path.suffix.lower() not in SUPPORTED_WORKBOOK_SUFFIXES:
        raise ValueError("Upload an .xlsx or .xlsm workbook.")
    if not path.is_file() or path.stat().st_size <= 0:
        raise ValueError("The uploaded workbook is empty or unavailable.")

    required_parts = (
        "[Content_Types].xml",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
    )
    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = set(archive.namelist())
            missing = [part for part in required_parts if part not in names]
            if missing:
                raise ValueError(
                    "The workbook package is missing required Excel parts: "
                    + ", ".join(missing)
                )
            # Read only the small structural parts. This catches truncated or
            # unreadable central-directory entries without decompressing the
            # entire workbook.
            for part in required_parts:
                archive.read(part)
    except (zipfile.BadZipFile, KeyError, OSError) as exc:
        raise ValueError("The uploaded file is not a readable Excel workbook.") from exc


def uploaded_file_cache_token(uploaded_file: Any) -> tuple[str, int]:
    """Return a stable, inexpensive identity for the current uploader value.

    The identity deliberately does not depend on Streamlit's transient
    ``file_id``. A browser refresh may allocate a new file ID for the same
    upload, which previously forced another full hash/copy/validation cycle.
    A name + size + head/tail sample is stable across reruns and refreshes while
    remaining inexpensive even for very large workbooks.
    """
    original_name = safe_filename(
        getattr(uploaded_file, "name", "client_workbook.xlsx"),
        "client_workbook.xlsx",
    )
    raw_size = getattr(uploaded_file, "size", None)
    try:
        view = memoryview(uploaded_file.getbuffer())
        size = len(view)
        sample_size = min(128 * 1024, size)
        digest = hashlib.sha256()
        digest.update(original_name.encode("utf-8", errors="replace"))
        digest.update(str(size).encode("ascii"))
        if sample_size:
            digest.update(view[:sample_size])
            if size > sample_size:
                digest.update(view[-sample_size:])
        return f"sample:{digest.hexdigest()}", size
    except Exception:
        try:
            size = int(raw_size or 0)
        except Exception:
            size = 0
        return f"metadata:{original_name}:{size}", size


def _stream_uploaded_file(
    uploaded_file: Any,
    destination: Path,
    *,
    expected_size: int = 0,
    load_progress: WorkbookLoadProgress | None = None,
) -> tuple[str, int]:
    """Write the upload once while calculating its SHA-256 in the same pass."""
    digest = hashlib.sha256()
    total = 0
    expected_size = max(0, int(expected_size or 0))

    def report_copy_progress() -> None:
        if load_progress is None or not load_progress.active:
            return
        if expected_size > 0:
            fraction = min(1.0, total / expected_size)
            percent = 6 + int(round(fraction * 28))
            detail = (
                f"Copied {total / (1024 * 1024):.1f} of "
                f"{expected_size / (1024 * 1024):.1f} MB into the secure runtime folder."
            )
        else:
            percent = min(34, 8 + int(total / (4 * 1024 * 1024)))
            detail = f"Copied {total / (1024 * 1024):.1f} MB into the secure runtime folder."
        load_progress.update(
            percent,
            "Uploading workbook",
            detail,
            phase="upload",
        )

    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    with destination.open("wb") as target:
        reader = getattr(uploaded_file, "read", None)
        if callable(reader):
            while True:
                chunk = reader(4 * 1024 * 1024)
                if not chunk:
                    break
                if isinstance(chunk, memoryview):
                    chunk = chunk.tobytes()
                target.write(chunk)
                digest.update(chunk)
                total += len(chunk)
                report_copy_progress()
        else:
            view = memoryview(uploaded_file.getbuffer())
            for offset in range(0, len(view), 4 * 1024 * 1024):
                chunk = view[offset : offset + 4 * 1024 * 1024]
                target.write(chunk)
                digest.update(chunk)
                total += len(chunk)
                report_copy_progress()
        target.flush()
        try:
            os.fsync(target.fileno())
        except OSError:
            pass
    return digest.hexdigest(), total


def save_uploaded_workbook(
    uploaded_file: Any,
    load_progress: WorkbookLoadProgress | None = None,
) -> tuple[Path, bool]:
    """Persist an uploaded workbook once and return ``(path, is_new_upload)``.

    Every material step is logged to the Streamlit shell and the persistent UI
    diagnostics file. A repeated rerun performs only a token/size/path check.
    """
    upload_dir = runtime_folder() / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)

    original_name = safe_filename(
        getattr(uploaded_file, "name", "client_workbook.xlsx"),
        "client_workbook.xlsx",
    )
    suffix = Path(original_name).suffix.lower()
    if suffix not in SUPPORTED_WORKBOOK_SUFFIXES:
        raise ValueError("Upload an .xlsx or .xlsm workbook.")

    with ui_stage(
        "upload.identify",
        warn_after=5,
        stack_after=20,
        file_name=original_name,
        reported_size=getattr(uploaded_file, "size", None),
    ):
        upload_token, upload_size = uploaded_file_cache_token(uploaded_file)

    previous_token = str(st.session_state.get("uploaded_widget_token") or "")
    previous_size = int(st.session_state.get("uploaded_file_size") or 0)
    previous_path_raw = st.session_state.get("uploaded_workbook_path")
    previous_path = Path(previous_path_raw) if previous_path_raw else None

    if (
        upload_token == previous_token
        and previous_path is not None
        and previous_path.is_file()
        and previous_path.stat().st_size == upload_size
        and (not previous_size or previous_size == upload_size)
    ):
        ui_log(
            "UPLOAD REUSED",
            file_name=original_name,
            size_bytes=upload_size,
            path=str(previous_path.resolve()),
        )
        return previous_path.resolve(), False

    if load_progress is not None:
        load_progress.begin(original_name, upload_size)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    partial = upload_dir / f".{stamp}_{os.getpid()}_{original_name}.part"
    destination: Path | None = None
    try:
        with ui_stage(
            "upload.copy_and_hash",
            warn_after=10,
            stack_after=45,
            file_name=original_name,
            expected_size_bytes=upload_size,
            partial_path=str(partial),
        ):
            signature, written_size = _stream_uploaded_file(
                uploaded_file,
                partial,
                expected_size=upload_size,
                load_progress=load_progress,
            )
        if written_size <= 0:
            raise ValueError("The uploaded workbook is empty.")
        if upload_size and written_size != upload_size:
            raise ValueError(
                "The uploaded workbook copy is incomplete: "
                f"expected {upload_size:,} bytes, wrote {written_size:,} bytes."
            )

        destination = upload_dir / f"{stamp}_{signature[:8]}_{original_name}"
        with ui_stage(
            "upload.commit",
            warn_after=5,
            stack_after=20,
            destination=str(destination),
        ):
            if load_progress is not None:
                load_progress.update(
                    38,
                    "Securing workbook copy",
                    "Finishing the local copy and locking in the uploaded file.",
                    phase="upload",
                )
            os.replace(partial, destination)

        with ui_stage(
            "upload.validate_ooxml",
            warn_after=8,
            stack_after=30,
            path=str(destination),
            size_bytes=written_size,
        ):
            if load_progress is not None:
                load_progress.update(
                    44,
                    "Verifying workbook",
                    "Checking the Excel package structure before reading any worksheets.",
                    phase="verify",
                )
            validate_workbook_package(destination)
    except Exception:
        partial.unlink(missing_ok=True)
        if destination is not None:
            destination.unlink(missing_ok=True)
        if load_progress is not None:
            load_progress.fail("The uploaded Excel package could not be validated.")
        raise

    st.session_state["uploaded_widget_token"] = upload_token
    st.session_state["uploaded_file_size"] = written_size
    st.session_state["uploaded_signature"] = signature
    st.session_state["uploaded_workbook_path"] = str(destination.resolve())
    st.session_state["uploaded_workbook_name"] = original_name
    ui_log(
        "UPLOAD READY",
        file_name=original_name,
        size_bytes=written_size,
        signature=signature[:16],
        path=str(destination.resolve()),
    )
    if load_progress is not None:
        load_progress.update(
            50,
            "Workbook verified",
            "The uploaded file is safe to inspect. Reading workbook structure next.",
            phase="verify",
        )
    return destination.resolve(), True


def validate_project_layout(paths: AppPaths) -> list[str]:
    problems: list[str] = []
    if not paths.ui_dir.is_dir():
        problems.append(f"UI folder not found: {paths.ui_dir}")
    if not paths.slide_maker_dir.is_dir():
        problems.append(f"Slide Maker folder not found: {paths.slide_maker_dir}")
    if not paths.backend_cli.is_file():
        problems.append(f"Backend CLI not found: {paths.backend_cli}")
    if not paths.python_exe.is_file():
        problems.append(f"Python executable not found: {paths.python_exe}")
    return problems


def initialize_control_panel_state() -> None:
    defaults = {
        "output_name": f"ThinkCell_Deck_{datetime.now():%Y%m%d_%H%M%S}.pptx",
        "workflow_mode": [],
        "input_section": "",
        "only_l0_filter": "",
        "keep_temp": False,
        "debug_logging": False,
        "uploaded_widget_token": "",
        "uploaded_file_size": 0,
        "uploaded_signature": "",
        "uploaded_workbook_path": None,
        "uploaded_workbook_name": "",
        "preview_workbook_path": None,
        "preview_source_signature": "",
        "preview_metadata": {},
        "preview_error": "",
        "preview_sheet_name": "",
        "preview_source_role": "aura",
        "preview_input_navigation_token": "",
        "selected_aura_sheet_override": "",
        "resolved_aura_sheet": "",
        "selected_taxonomy_sheet_override": "",
        "resolved_taxonomy_sheet": "",
        "resolved_geocost_sheet": "",
        "taxonomy_sheet_profile": {},
        "selected_geocost_sheet_override": "",
        "selected_geocost_scope_override": "",
        "selected_geocost_global_sheet_override": "",
        "selected_geocost_us_sheet_override": "",
        "workbook_range_selections": {},
        "active_range_selection": None,
        "slide_input_mappings": {},
        "aura_ignored_taxonomy_functions": {},
        "active_input_by_section": {},
        "l1_mapping_validation": {},
        "l1_auto_detection": {},
        "geocost_auto_detection": {},
        "geocost_input_profile": {},
        "geocost_mapping_validation": {},
        "l1_master_label_profile": {},
        "l1_function_lookup_profile": {},
        "l1_sort_order": L1_SORT_ORDER_DEFAULT,
        "l1_segment_sort_direction": L1_SEGMENT_SORT_DIRECTION_DEFAULT,
        "l1_segment_profile": {},
        "l1_segment_profile_signature": "",
        "auto_mapping_rerun_signature": "",
        "segment_profile_rerun_signature": "",
        "l1_segment_name_overrides": {},
        "l1_segment_rename_open": {},
        "l1_segment_assignment_overrides": {},
        "l1_segment_assignment_editor_open": {},
        "workbook_review_finalized": False,
        "canvas_stage": "review",
        "input_sidebar_autostart_token": 0,
        "input_sidebar_session_nonce": f"{os.getpid()}-{datetime.now():%Y%m%d%H%M%S%f}",
        "sort_sidebar_open": False,
        "sort_sidebar_autostart_token": 0,
        "sort_sidebar_session_nonce": f"{os.getpid()}-{datetime.now():%Y%m%d%H%M%S%f}",
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)

    # UI default v2: the slide-family multiselect starts empty.  Reset the
    # previous implicit "select everything" default once per browser session
    # after this upgrade; subsequent user selections are preserved normally.
    workflow_default_version = int(
        st.session_state.get("workflow_empty_default_version") or 0
    )
    if workflow_default_version < 2:
        st.session_state["workflow_mode"] = []
        st.session_state["input_section"] = ""
        st.session_state["workflow_empty_default_version"] = 2

    # Migrate the previous single-select state (for example ``"complete"`` or
    # ``"l1"``) into the four-value multiselect without requiring users to
    # clear browser state.
    raw_workflow_state = st.session_state.get("workflow_mode")
    normalized_sections = normalize_deck_sections(raw_workflow_state)
    # An explicitly empty list is a valid transient multiselect state and is
    # handled by the disabled Generate button. Only legacy strings need a
    # migration fallback.
    if isinstance(raw_workflow_state, str) and not normalized_sections:
        normalized_sections = DECK_SECTION_ORDER
    st.session_state["workflow_mode"] = list(normalized_sections)


def _clear_uploaded_workbook_state() -> None:
    st.session_state["uploaded_widget_token"] = ""
    st.session_state["uploaded_file_size"] = 0
    st.session_state["uploaded_signature"] = ""
    st.session_state["uploaded_workbook_path"] = None
    st.session_state["uploaded_workbook_name"] = ""
    st.session_state["preview_workbook_path"] = None
    st.session_state["preview_source_signature"] = ""
    st.session_state["preview_metadata"] = {}
    st.session_state["preview_error"] = ""
    st.session_state["preview_sheet_name"] = ""
    st.session_state["preview_source_role"] = "aura"
    st.session_state["preview_input_navigation_token"] = ""
    st.session_state["selected_aura_sheet_override"] = ""
    st.session_state["resolved_aura_sheet"] = ""
    st.session_state["selected_taxonomy_sheet_override"] = ""
    st.session_state["resolved_taxonomy_sheet"] = ""
    st.session_state["resolved_geocost_sheet"] = ""
    st.session_state["taxonomy_sheet_profile"] = {}
    st.session_state["selected_geocost_sheet_override"] = ""
    st.session_state["selected_geocost_scope_override"] = ""
    st.session_state["selected_geocost_global_sheet_override"] = ""
    st.session_state["selected_geocost_us_sheet_override"] = ""
    st.session_state["workbook_range_selections"] = {}
    st.session_state["active_range_selection"] = None
    st.session_state["slide_input_mappings"] = {}
    st.session_state["active_input_by_section"] = {}
    st.session_state["l1_mapping_validation"] = {}
    st.session_state["l1_auto_detection"] = {}
    st.session_state["geocost_auto_detection"] = {}
    st.session_state["geocost_input_profile"] = {}
    st.session_state["geocost_mapping_validation"] = {}
    st.session_state["l1_master_label_profile"] = {}
    st.session_state["l1_function_lookup_profile"] = {}
    st.session_state["l1_segment_profile"] = {}
    st.session_state["l1_segment_profile_signature"] = ""
    st.session_state["auto_mapping_rerun_signature"] = ""
    st.session_state["segment_profile_rerun_signature"] = ""
    st.session_state["l1_segment_name_overrides"] = {}
    st.session_state["l1_segment_rename_open"] = {}
    st.session_state["l1_segment_assignment_overrides"] = {}
    st.session_state["l1_segment_assignment_editor_open"] = {}
    st.session_state["l1_sort_order"] = L1_SORT_ORDER_DEFAULT
    st.session_state["l1_segment_sort_direction"] = L1_SEGMENT_SORT_DIRECTION_DEFAULT
    st.session_state["sort_sidebar_open"] = False
    st.session_state["sort_sidebar_autostart_token"] = 0
    st.session_state["workbook_review_finalized"] = False
    st.session_state["canvas_stage"] = "review"


# -----------------------------------------------------------------------------
# Control-panel rendering
# -----------------------------------------------------------------------------


def _input_sections_for_deck_sections(sections: Any) -> tuple[str, ...]:
    """Return the single standardized Inputs rail when any slide is selected."""
    return ("inputs",) if normalize_deck_sections(sections) else ()


def _synchronize_input_section(sections: Any) -> tuple[str, ...]:
    """Keep the internal standardized Inputs-rail state synchronized."""
    options = _input_sections_for_deck_sections(sections)
    st.session_state["input_section"] = "inputs" if options else ""
    return options


def _return_to_workbook_review() -> None:
    """A source-configuration change must be reviewed before generation."""
    st.session_state["workbook_review_finalized"] = False
    st.session_state["canvas_stage"] = "review"


def _request_l1_input_sidebar_autostart() -> None:
    """Ask the browser controller to collapse the main rail and open Inputs once."""
    current = int(st.session_state.get("input_sidebar_autostart_token") or 0)
    st.session_state["input_sidebar_autostart_token"] = current + 1


def _request_sort_sidebar_autostart() -> None:
    """Open and expand the dedicated company-order rail on the next rerun."""
    st.session_state["sort_sidebar_open"] = True
    current = int(st.session_state.get("sort_sidebar_autostart_token") or 0)
    st.session_state["sort_sidebar_autostart_token"] = current + 1


def _close_sort_sidebar() -> None:
    """Hide the dedicated company-order rail completely."""
    st.session_state["sort_sidebar_open"] = False


def _on_workflow_change() -> None:
    selected_sections = normalize_deck_sections(
        st.session_state.get("workflow_mode")
    )
    # Preserve canonical order in Streamlit's multi-select chips and state.
    st.session_state["workflow_mode"] = list(selected_sections)
    _synchronize_input_section(selected_sections)
    if selected_sections:
        _request_l1_input_sidebar_autostart()
    # The company-order rail is deliberately user-triggered from the bottom of
    # the Inputs rail. Changing deck scope closes it so it never opens by
    # surprise or shows settings for a no-longer-selected slide family.
    st.session_state["sort_sidebar_open"] = False
    _return_to_workbook_review()


def _on_input_section_change() -> None:
    # Retained for backward compatibility with stale browser sessions. The
    # standardized Inputs rail no longer has a separate family selector.
    if normalize_deck_sections(st.session_state.get("workflow_mode")):
        _request_l1_input_sidebar_autostart()
    _return_to_workbook_review()


def _section_heading(number: int, title: str, caption: str = "") -> None:
    caption_html = (
        f'<div class="tc-panel-section-caption">{html.escape(caption)}</div>'
        if caption
        else ""
    )
    st.markdown(
        f"""
        <div class="tc-panel-section-heading">
          <span class="tc-panel-section-number">{number:02d}</span>
          <div>
            <div class="tc-panel-section-title">{html.escape(title)}</div>
            {caption_html}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _panel_divider() -> None:
    st.markdown('<div class="tc-panel-divider"></div>', unsafe_allow_html=True)


def _field_label(
    title: str,
    hint: str = "",
    *,
    spaced: bool = False,
) -> None:
    """Render a stable sidebar label independent of Streamlit's widget DOM."""
    classes = "tc-panel-field-label"
    if spaced:
        classes += " tc-panel-field-label-spaced"
    hint_html = (
        f'<span class="tc-panel-field-label-hint">{html.escape(hint)}</span>'
        if hint
        else ""
    )
    markup = (
        f'<div class="{classes}">'
        f'<span class="tc-panel-field-label-text">{html.escape(title)}</span>'
        f'{hint_html}'
        '</div>'
    )
    st.markdown(markup, unsafe_allow_html=True)


def _inject_final_light_control_text_css() -> None:
    """Force dark text on the sidebar's light BaseWeb controls.

    This is intentionally rendered after the widgets. Streamlit versions differ
    in whether a selectbox keeps the ``stSelectbox`` wrapper, so the CSS targets
    both the explicit widget keys and the underlying BaseWeb elements.
    """
    st.markdown(
        """
        <style data-tc-ui-style="true">
          [data-testid="stSidebar"] .st-key-workflow_mode div[data-baseweb="select"],
          [data-testid="stSidebar"] .st-key-workflow_mode div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] .st-key-workflow_mode div[role="combobox"],
          [data-testid="stSidebar"] .st-key-input_section div[data-baseweb="select"],
          [data-testid="stSidebar"] .st-key-input_section div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] .st-key-input_section div[role="combobox"],
          [data-testid="stSidebar"] div[data-baseweb="select"],
          [data-testid="stSidebar"] div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] div[role="combobox"] {
            background: #ffffff !important;
            background-color: #ffffff !important;
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] .st-key-workflow_mode div[data-baseweb="select"] *,
          [data-testid="stSidebar"] .st-key-workflow_mode div[role="combobox"] *,
          [data-testid="stSidebar"] .st-key-input_section div[data-baseweb="select"] *,
          [data-testid="stSidebar"] .st-key-input_section div[role="combobox"] *,
          [data-testid="stSidebar"] div[data-baseweb="select"] *,
          [data-testid="stSidebar"] div[role="combobox"] * {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            text-shadow: none !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] .st-key-workflow_mode svg,
          [data-testid="stSidebar"] .st-key-input_section svg,
          [data-testid="stSidebar"] div[data-baseweb="select"] svg,
          [data-testid="stSidebar"] div[role="combobox"] svg {
            color: #343a42 !important;
            fill: #343a42 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="input"],
          [data-testid="stSidebar"] div[data-baseweb="base-input"],
          [data-testid="stSidebar"] div[data-baseweb="input"] > div,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] > div {
            background: #ffffff !important;
            background-color: #ffffff !important;
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="input"] input,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] input,
          [data-testid="stSidebar"] div[data-baseweb="input"] input:disabled,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] input:disabled {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="input"] input::placeholder,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] input::placeholder {
            color: #7a828d !important;
            -webkit-text-fill-color: #7a828d !important;
            opacity: 1 !important;
          }

          /* Slide-family multiselect: Bain-red selected chips, white bold type. */
          [data-testid="stSidebar"] .st-key-workflow_mode [data-baseweb="tag"] {
            background: #cc0000 !important;
            background-color: #cc0000 !important;
            border-color: #cc0000 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-weight: 700 !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] .st-key-workflow_mode [data-baseweb="tag"] * {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-weight: 700 !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
          }

          /* Keep the clear/menu affordances visually integrated with the white selector. */
          [data-testid="stSidebar"] .st-key-workflow_mode div[data-baseweb="select"] button:not([data-baseweb="tag"] *),
          [data-testid="stSidebar"] .st-key-workflow_mode div[data-baseweb="select"] [role="button"]:not([data-baseweb="tag"] *) {
            background: transparent !important;
            background-color: transparent !important;
            border: 0 !important;
            box-shadow: none !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )



def _inject_rendered_control_text_repair(
    *,
    input_sidebar_active: bool,
    segment_sidebar_active: bool,
    segment_sidebar_requested_open: bool,
) -> None:
    """Repair sidebar control colors in the rendered parent DOM.

    Streamlit/BaseWeb class names and nesting vary between releases. CSS alone
    has proved unreliable on the installed build, so this zero-height component
    applies the final colors directly to the actual rendered controls and keeps
    them correct after reruns or dropdown portal creation.
    """
    autostart_counter = int(st.session_state.get("input_sidebar_autostart_token") or 0)
    session_nonce = str(st.session_state.get("input_sidebar_session_nonce") or "session")
    autostart_token = (
        f"{session_nonce}:{autostart_counter}"
        if autostart_counter > 0
        else ""
    )
    sort_autostart_counter = int(st.session_state.get("sort_sidebar_autostart_token") or 0)
    sort_session_nonce = str(st.session_state.get("sort_sidebar_session_nonce") or "sort-session")
    sort_autostart_token = (
        f"{sort_session_nonce}:{sort_autostart_counter}"
        if sort_autostart_counter > 0
        else ""
    )
    component_markup = r"""
        <!doctype html>
        <html>
          <head>
            <style>
              html, body { margin: 0; padding: 0; width: 0; height: 0; overflow: hidden; }
            </style>
          </head>
          <body>
            <script>
              (() => {
                const parentWindow = window.parent;
                const doc = parentWindow.document;
                const INK = '#111418';
                const ARROW = '#343a42';
                const SURFACE = '#ffffff';
                const BORDER = '#d5d9df';
                const SIDEBAR_STATE_KEY = 'thinkcell.slide-maker.sidebar.compact.v2';
                const SIDEBAR_CLASS = 'tc-sidebar-is-collapsed';
                const INPUT_SIDEBAR_STATE_KEY = 'thinkcell.slide-maker.input-sidebar.compact.v1';
                const INPUT_SIDEBAR_AUTOSTART_KEY = 'thinkcell.slide-maker.input-sidebar.autostart.v1';
                const INPUT_SIDEBAR_OPEN_CLASS = 'tc-input-sidebar-is-open';
                const INPUT_SIDEBAR_COLLAPSED_CLASS = 'tc-input-sidebar-is-collapsed';
                const INPUT_SIDEBAR_ACTIVE = __TC_INPUT_SIDEBAR_ACTIVE__;
                const INPUT_SIDEBAR_AUTOSTART_TOKEN = '__TC_INPUT_SIDEBAR_AUTOSTART_TOKEN__';
                const SEGMENT_SIDEBAR_STATE_KEY = 'thinkcell.slide-maker.sort-sidebar.compact.v77';
                const SEGMENT_SIDEBAR_AUTOSTART_KEY = 'thinkcell.slide-maker.sort-sidebar.autostart.v77';
                const SEGMENT_SIDEBAR_OPEN_CLASS = 'tc-segment-sidebar-is-open';
                const SEGMENT_SIDEBAR_COLLAPSED_CLASS = 'tc-segment-sidebar-is-collapsed';
                const SEGMENT_SIDEBAR_ACTIVE = __TC_SEGMENT_SIDEBAR_ACTIVE__;
                const SEGMENT_SIDEBAR_REQUESTED_OPEN = __TC_SEGMENT_SIDEBAR_REQUESTED_OPEN__;
                const SEGMENT_SIDEBAR_AUTOSTART_TOKEN = '__TC_SEGMENT_SIDEBAR_AUTOSTART_TOKEN__';

                function important(node, property, value) {
                  if (!node || !node.style) return;
                  node.style.setProperty(property, value, 'important');
                }

                function forceText(node, color = INK) {
                  important(node, 'color', color);
                  important(node, '-webkit-text-fill-color', color);
                  important(node, 'text-shadow', 'none');
                  important(node, '-webkit-text-stroke', '0 transparent');
                  important(node, 'opacity', '1');
                  important(node, 'filter', 'none');
                }

                function repairRoot(root) {
                  if (!root) return;

                  important(root, '--text-color', INK);
                  important(root, '--secondary-text-color', INK);
                  important(root, '--disabled-text-color', INK);
                  important(root, 'opacity', '1');

                  let ancestor = root;
                  const sidebar = doc.querySelector('[data-testid="stSidebar"]');
                  while (ancestor && ancestor !== sidebar) {
                    important(ancestor, 'opacity', '1');
                    ancestor = ancestor.parentElement;
                  }

                  const surfaces = root.querySelectorAll([
                    'div[data-baseweb="select"]',
                    'div[role="combobox"]',
                    'button[aria-haspopup="listbox"]',
                    'button[aria-expanded]',
                    'div[data-baseweb="input"]',
                    'div[data-baseweb="base-input"]'
                  ].join(','));

                  surfaces.forEach((surface) => {
                    important(surface, 'background', SURFACE);
                    important(surface, 'background-color', SURFACE);
                    important(surface, 'background-image', 'none');
                    important(surface, 'border-color', BORDER);
                    important(surface, 'color-scheme', 'light');
                    important(surface, 'opacity', '1');
                    forceText(surface);
                  });

                  [root, ...root.querySelectorAll('*')].forEach((node) => {
                    if (!node || ['STYLE', 'SCRIPT'].includes(node.tagName)) return;
                    forceText(node);
                  });

                  root.querySelectorAll('input, textarea').forEach((input) => {
                    important(input, 'background-color', 'transparent');
                    important(input, 'caret-color', INK);
                    forceText(input);
                  });

                  root.querySelectorAll('svg').forEach((svg) => {
                    important(svg, 'color', ARROW);
                    important(svg, 'fill', ARROW);
                    important(svg, 'opacity', '1');
                  });
                  root.querySelectorAll('svg *').forEach((shape) => {
                    important(shape, 'color', ARROW);
                    important(shape, 'fill', ARROW);
                    important(shape, 'stroke', ARROW);
                    important(shape, 'opacity', '1');
                  });
                }

                function repairWorkflowMultiselect(sidebar) {
                  if (!sidebar) return;
                  const deckLabels = new Set(['Aura Summary', 'Aura L0', 'Aura L1', 'Geo Cost']);
                  const BAIN = '#cc0000';
                  const WHITE = '#ffffff';
                  const DECK_SURFACE = '#111214';
                  const DECK_BORDER = '#34383f';
                  const ICON = '#6f7782';

                  const candidates = [...sidebar.querySelectorAll('[data-testid="stMultiSelect"]')];
                  let root = candidates.find((candidate) => {
                    const text = String(candidate.innerText || candidate.textContent || '');
                    return [...deckLabels].some((label) => text.includes(label));
                  });
                  if (!root) {
                    const fieldRoot = sidebar.querySelector('.st-key-field_workflow_mode');
                    root = fieldRoot && fieldRoot.querySelector('[data-testid="stMultiSelect"]');
                  }
                  if (!root) return;

                  const select = root.querySelector('div[data-baseweb="select"]') || root.querySelector('div[role="combobox"]') || root;
                  important(select, 'position', 'relative');
                  important(select, 'background', DECK_SURFACE);
                  important(select, 'background-color', DECK_SURFACE);
                  important(select, 'background-image', 'none');
                  important(select, 'border-color', DECK_BORDER);
                  important(select, 'box-shadow', 'none');
                  important(select, 'color-scheme', 'dark');
                  important(select, 'padding-right', '3.6rem');

                  // Match only the multiselect widget itself to the output-name
                  // control. Do not resize/repaint the surrounding field wrapper.
                  const outputRef = sidebar.querySelector(
                    '.st-key-field_output_name div[data-baseweb="input"], ' +
                    '.st-key-field_output_name div[data-baseweb="base-input"]'
                  );
                  const outputRect = outputRef ? outputRef.getBoundingClientRect() : null;
                  if (outputRect && outputRect.width > 40) {
                    important(root, 'width', `${outputRect.width}px`);
                    important(root, 'max-width', '100%');
                    important(select, 'width', '100%');
                    important(select, 'max-width', '100%');
                    important(select, 'box-sizing', 'border-box');
                  }

                  // Paint only the selector shell dark. Never paint the shared
                  // value-container red; each selected item gets its own chip.
                  [...select.children].forEach((child) => {
                    if (child.classList && child.classList.contains('tc-workflow-clean-icons')) return;
                    important(child, 'background', 'transparent');
                    important(child, 'background-color', 'transparent');
                    important(child, 'background-image', 'none');
                  });

                  const leaves = [...root.querySelectorAll('*')].filter((node) => {
                    if (node.children.length) return false;
                    return deckLabels.has(String(node.textContent || '').trim());
                  });

                  const chips = [];
                  const labelLeafCount = (node) => {
                    if (!node) return 0;
                    let count = deckLabels.has(String(node.textContent || '').trim()) && node.children.length === 0 ? 1 : 0;
                    node.querySelectorAll('*').forEach((child) => {
                      if (child.children.length) return;
                      if (deckLabels.has(String(child.textContent || '').trim())) count += 1;
                    });
                    return count;
                  };

                  leaves.forEach((leaf) => {
                    let chip = null;
                    const nativeTag = leaf.closest('[data-baseweb="tag"]');
                    if (nativeTag && labelLeafCount(nativeTag) === 1) chip = nativeTag;

                    if (!chip) {
                      let cursor = leaf.parentElement;
                      let fallback = leaf.parentElement || leaf;
                      let depth = 0;
                      while (cursor && cursor !== select && cursor !== root && depth < 7) {
                        const rect = cursor.getBoundingClientRect();
                        const oneLabelOnly = labelLeafCount(cursor) === 1;
                        if (oneLabelOnly && rect.height >= 18 && rect.height <= 52 && rect.width >= 35 && rect.width <= 180) {
                          fallback = cursor;
                          if (cursor.querySelector('svg, button, [role="button"]')) {
                            chip = cursor;
                            break;
                          }
                        }
                        cursor = cursor.parentElement;
                        depth += 1;
                      }
                      if (!chip) chip = fallback;
                    }

                    if (!chip || chips.includes(chip)) return;
                    chips.push(chip);
                    important(chip, 'display', 'inline-flex');
                    important(chip, 'align-items', 'center');
                    important(chip, 'flex', '0 0 auto');
                    important(chip, 'width', 'auto');
                    important(chip, 'min-width', '0');
                    important(chip, 'max-width', '11rem');
                    important(chip, 'height', '2.18rem');
                    important(chip, 'margin', '0.14rem 0.28rem 0.14rem 0');
                    important(chip, 'padding', '0 0.52rem');
                    important(chip, 'background', BAIN);
                    important(chip, 'background-color', BAIN);
                    important(chip, 'background-image', 'none');
                    important(chip, 'border', `1px solid ${BAIN}`);
                    important(chip, 'border-radius', '0.46rem');
                    important(chip, 'box-shadow', 'none');
                    important(chip, 'font-weight', '800');
                    [chip, ...chip.querySelectorAll('*')].forEach((node) => {
                      forceText(node, WHITE);
                      important(node, 'font-weight', '800');
                    });
                    chip.querySelectorAll('svg, svg *').forEach((node) => {
                      important(node, 'color', WHITE);
                      important(node, 'fill', WHITE);
                      important(node, 'stroke', WHITE);
                      important(node, 'opacity', '1');
                    });
                  });

                  // Normalize the common value row so multiple selected slides
                  // remain separate compact chips rather than one wide red slab.
                  if (chips.length) {
                    let commonParent = chips[0].parentElement;
                    while (commonParent && commonParent !== select && !chips.every((chip) => commonParent.contains(chip))) {
                      commonParent = commonParent.parentElement;
                    }
                    if (commonParent && commonParent !== select) {
                      important(commonParent, 'display', 'flex');
                      important(commonParent, 'align-items', 'center');
                      important(commonParent, 'flex-wrap', 'wrap');
                      important(commonParent, 'gap', '0.08rem');
                      important(commonParent, 'width', 'auto');
                      important(commonParent, 'max-width', 'calc(100% - 3.6rem)');
                      important(commonParent, 'background', 'transparent');
                      important(commonParent, 'background-color', 'transparent');
                      important(commonParent, 'background-image', 'none');
                    }
                  }

                  // Native control artwork differs by Streamlit/BaseWeb release.
                  // Keep the real controls clickable but make their surface clean.
                  select.querySelectorAll('button, [role="button"]').forEach((control) => {
                    const inChip = chips.some((chip) => chip.contains(control));
                    if (inChip) return;
                    important(control, 'background', 'transparent');
                    important(control, 'background-color', 'transparent');
                    important(control, 'background-image', 'none');
                    important(control, 'border', '0');
                    important(control, 'box-shadow', 'none');
                  });
                  select.querySelectorAll('svg').forEach((svg) => {
                    const inChip = chips.some((chip) => chip.contains(svg));
                    if (inChip) return;
                    important(svg, 'opacity', '0');
                    important(svg, 'color', 'transparent');
                    important(svg, 'fill', 'transparent');
                    important(svg, 'stroke', 'transparent');
                  });

                  let overlay = select.querySelector(':scope > .tc-workflow-clean-icons');
                  if (!overlay) {
                    overlay = doc.createElement('div');
                    overlay.className = 'tc-workflow-clean-icons';
                    select.appendChild(overlay);
                  }
                  const hasSelection = chips.length > 0;

                  // Empty state only: stretch the existing value surface to the
                  // reference field width while preserving Streamlit's widget
                  // hierarchy. This is intentionally NOT applied to the wrapper
                  // or to selected-chip layouts.
                  if (!hasSelection) {
                    let emptyValue = null;
                    const placeholderLeaf = [...select.querySelectorAll('*')].find((node) => {
                      if (node.children.length) return false;
                      return String(node.textContent || '').trim() === 'Select one or more slide types';
                    });
                    let probe = placeholderLeaf || select.querySelector('input');
                    while (probe && probe.parentElement && probe.parentElement !== select) {
                      probe = probe.parentElement;
                    }
                    if (probe && probe.parentElement === select) emptyValue = probe;

                    if (emptyValue) {
                      important(select, 'padding-right', '0');
                      important(emptyValue, 'display', 'flex');
                      important(emptyValue, 'align-items', 'center');
                      important(emptyValue, 'flex', '1 1 auto');
                      important(emptyValue, 'width', '100%');
                      important(emptyValue, 'max-width', '100%');
                      important(emptyValue, 'min-width', '0');
                      important(emptyValue, 'box-sizing', 'border-box');
                      important(emptyValue, 'background', WHITE);
                      important(emptyValue, 'background-color', WHITE);
                      important(emptyValue, 'border', '1px solid #d5d9df');
                      important(emptyValue, 'padding-right', '2.7rem');

                      let radius = '8px';
                      if (outputRef) {
                        const refStyle = parentWindow.getComputedStyle(outputRef);
                        if (refStyle && refStyle.borderRadius) radius = refStyle.borderRadius;
                      }
                      important(emptyValue, 'border-radius', radius);

                      if (outputRect && outputRect.height > 20) {
                        important(emptyValue, 'height', `${outputRect.height}px`);
                        important(emptyValue, 'min-height', `${outputRect.height}px`);
                      }
                      [emptyValue, ...emptyValue.querySelectorAll('*')].forEach((node) => {
                        if (node.tagName && node.tagName.toLowerCase() === 'svg') return;
                        important(node, 'color', '#111418');
                        important(node, '-webkit-text-fill-color', '#111418');
                      });
                    }
                  }

                  overlay.innerHTML = `${hasSelection ? '<span class="tc-workflow-clean-clear">×</span>' : ''}`;
                  important(overlay, 'position', 'absolute');
                  important(overlay, 'right', '10px');
                  important(overlay, 'top', '50%');
                  important(overlay, 'transform', 'translateY(-50%)');
                  important(overlay, 'display', 'flex');
                  important(overlay, 'align-items', 'center');
                  important(overlay, 'gap', '10px');
                  important(overlay, 'pointer-events', 'none');
                  important(overlay, 'z-index', '8');
                  overlay.querySelectorAll('span').forEach((span) => {
                    forceText(span, ICON);
                    important(span, 'font-size', '16px');
                    important(span, 'line-height', '1');
                    important(span, 'font-weight', '800');
                    important(span, 'background', 'transparent');
                    important(span, 'border', '0');
                    important(span, 'box-shadow', 'none');
                  });
                }

                function repairPortalOptions() {
                  doc.querySelectorAll([
                    'div[data-baseweb="popover"] [role="listbox"]',
                    'div[data-baseweb="popover"] [role="option"]',
                    'div[data-baseweb="popover"] [role="option"] *'
                  ].join(',')).forEach((node) => {
                    important(node, 'background-color', SURFACE);
                    forceText(node);
                  });
                }


                function readSidebarCollapsedState() {
                  try {
                    const stored = parentWindow.localStorage.getItem(SIDEBAR_STATE_KEY);
                    if (stored === '1') return true;
                    if (stored === '0') return false;
                  } catch (_) {}
                  return doc.documentElement.classList.contains(SIDEBAR_CLASS);
                }

                function writeSidebarCollapsedState(collapsed) {
                  try {
                    parentWindow.localStorage.setItem(
                      SIDEBAR_STATE_KEY,
                      collapsed ? '1' : '0'
                    );
                  } catch (_) {}
                }

                function sidebarToggleIcon(collapsed) {
                  const path = collapsed
                    ? 'M2.5 3.5L7 8l-4.5 4.5M7 3.5L11.5 8 7 12.5'
                    : 'M11.5 3.5L7 8l4.5 4.5M7 3.5L2.5 8 7 12.5';
                  return `
                    <svg class="tc-sidebar-toggle-icon" viewBox="0 0 14 16" aria-hidden="true">
                      <path d="${path}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"></path>
                    </svg>
                  `;
                }

                function syncSidebarToggle(collapsed) {
                  const button = doc.getElementById('tc-sidebar-toggle');
                  if (button) {
                    const label = collapsed ? 'Expand control panel' : 'Collapse control panel';
                    button.dataset.collapsed = collapsed ? 'true' : 'false';
                    button.setAttribute('aria-label', label);
                    button.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
                    if (button.dataset.iconState !== button.dataset.collapsed) {
                      button.innerHTML = sidebarToggleIcon(collapsed);
                      button.dataset.iconState = button.dataset.collapsed;
                    }
                  }

                  const miniBrand = doc.getElementById('tc-sidebar-mini-brand');
                  if (miniBrand) {
                    miniBrand.setAttribute('aria-hidden', collapsed ? 'false' : 'true');
                  }
                }

                function announceSidebarResize() {
                  parentWindow.requestAnimationFrame(() => {
                    try {
                      parentWindow.dispatchEvent(new parentWindow.Event('resize'));
                    } catch (_) {}
                    parentWindow.setTimeout(() => {
                      try {
                        parentWindow.dispatchEvent(new parentWindow.Event('resize'));
                      } catch (_) {}
                    }, 240);
                  });
                }

                function setSidebarCollapsed(collapsed, persist = true) {
                  const root = doc.documentElement;
                  const body = doc.body;
                  const next = Boolean(collapsed);
                  const changed = root.classList.contains(SIDEBAR_CLASS) !== next;

                  root.classList.toggle(SIDEBAR_CLASS, next);
                  if (body) body.classList.toggle(SIDEBAR_CLASS, next);
                  parentWindow.__tcSidebarCollapsed = next;
                  if (persist) writeSidebarCollapsedState(next);
                  syncSidebarToggle(next);
                  if (changed) announceSidebarResize();
                }

                function miniSidebarIcon(kind) {
                  const icons = {
                    upload: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M12 15V4m0 0L8.5 7.5M12 4l3.5 3.5M5 14.5v3A2.5 2.5 0 0 0 7.5 20h9a2.5 2.5 0 0 0 2.5-2.5v-3" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>
                      </svg>`,
                    slides: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <rect x="5" y="5" width="14" height="10" rx="2" fill="none" stroke="currentColor" stroke-width="1.8"/>
                        <path d="M8 19h8M12 15v4M8.5 9h7M8.5 12h4.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                      </svg>`,
                    generate: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M8 5.5v13l11-6.5L8 5.5Z" fill="currentColor"/>
                      </svg>`
                  };
                  return icons[kind] || '';
                }

                function ensureMiniSidebarActions(sidebar) {
                  if (!sidebar) return null;
                  let rail = doc.getElementById('tc-sidebar-mini-actions');
                  if (!rail || rail.parentElement !== sidebar) {
                    if (rail) rail.remove();
                    rail = doc.createElement('div');
                    rail.id = 'tc-sidebar-mini-actions';
                    rail.setAttribute('aria-label', 'Compact Slide Maker controls');
                    rail.innerHTML = `
                      <button type="button" class="tc-sidebar-mini-action" data-action="upload" aria-label="Upload source workbook" title="Upload workbook">
                        ${miniSidebarIcon('upload')}
                        <span class="tc-sidebar-mini-action-label">Upload</span>
                      </button>
                      <button type="button" class="tc-sidebar-mini-action" data-action="slides" aria-label="Choose slide types" title="Choose slide types">
                        ${miniSidebarIcon('slides')}
                        <span class="tc-sidebar-mini-action-label">Slides</span>
                      </button>
                      <button type="button" class="tc-sidebar-mini-action is-generate" data-action="generate" aria-label="Generate PowerPoint" title="Generate PowerPoint">
                        ${miniSidebarIcon('generate')}
                        <span class="tc-sidebar-mini-action-label">Build</span>
                      </button>`;
                    sidebar.appendChild(rail);
                  }

                  const uploadButton = rail.querySelector('[data-action="upload"]');
                  if (uploadButton) {
                    uploadButton.onclick = (event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      const input = sidebar.querySelector('[data-testid="stFileUploader"] input[type="file"], input[type="file"]');
                      if (input) input.click();
                      else {
                        setSidebarCollapsed(false, true);
                        parentWindow.setTimeout(() => {
                          const retry = sidebar.querySelector('[data-testid="stFileUploader"] input[type="file"], input[type="file"]');
                          if (retry) retry.click();
                        }, 180);
                      }
                    };
                  }

                  const slidesButton = rail.querySelector('[data-action="slides"]');
                  if (slidesButton) {
                    slidesButton.onclick = (event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      setSidebarCollapsed(false, true);
                      parentWindow.setTimeout(() => {
                        const workflowRoot = sidebar.querySelector('.st-key-workflow_mode');
                        const combobox = workflowRoot && workflowRoot.querySelector('[role="combobox"], div[data-baseweb="select"]');
                        if (combobox) {
                          combobox.scrollIntoView({ block: 'center', behavior: 'smooth' });
                          combobox.click();
                        }
                      }, 180);
                    };
                  }

                  const generateButton = rail.querySelector('[data-action="generate"]');
                  if (generateButton) {
                    generateButton.onclick = (event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      if (generateButton.disabled) return;
                      const realButton = sidebar.querySelector('.st-key-generate_powerpoint button, button[kind="primary"]');
                      if (realButton && !realButton.disabled && realButton.getAttribute('aria-disabled') !== 'true') {
                        realButton.click();
                      }
                    };
                  }
                  return rail;
                }

                function syncMiniSidebarActions(sidebar) {
                  const rail = ensureMiniSidebarActions(sidebar);
                  if (!rail) return;
                  const realGenerate = sidebar.querySelector('.st-key-generate_powerpoint button, button[kind="primary"]');
                  const miniGenerate = rail.querySelector('[data-action="generate"]');
                  if (miniGenerate) {
                    const disabled = !realGenerate || realGenerate.disabled || realGenerate.getAttribute('aria-disabled') === 'true';
                    miniGenerate.disabled = disabled;
                    miniGenerate.setAttribute('aria-disabled', disabled ? 'true' : 'false');
                    miniGenerate.title = disabled ? 'Generate PowerPoint (finish setup first)' : 'Generate PowerPoint';
                  }
                }

                function ensureSidebarToggle(sidebar) {
                  if (!sidebar) return;
                  if (!sidebar.id) sidebar.id = 'tc-control-panel';

                  let miniBrand = doc.getElementById('tc-sidebar-mini-brand');
                  if (!miniBrand || miniBrand.parentElement !== sidebar) {
                    if (miniBrand) miniBrand.remove();
                    miniBrand = doc.createElement('div');
                    miniBrand.id = 'tc-sidebar-mini-brand';
                    miniBrand.innerHTML = '<span>TC</span>';
                    miniBrand.setAttribute('aria-label', 'ThinkCell Slide Maker');
                    sidebar.appendChild(miniBrand);
                  }

                  ensureMiniSidebarActions(sidebar);

                  let button = doc.getElementById('tc-sidebar-toggle');
                  if (!button || button.parentElement !== sidebar) {
                    if (button) button.remove();
                    button = doc.createElement('button');
                    button.id = 'tc-sidebar-toggle';
                    button.type = 'button';
                    button.setAttribute('aria-controls', 'tc-control-panel');
                    sidebar.appendChild(button);
                  }

                  button.onclick = (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    setSidebarCollapsed(!readSidebarCollapsedState(), true);
                  };

                  // A fresh page/app load always begins with the primary rail
                  // expanded. The browser-global flag survives Streamlit reruns
                  // but resets on a real page reload/new app session.
                  if (!parentWindow.__tcSidebarStartupExpandedV1) {
                    parentWindow.__tcSidebarStartupExpandedV1 = true;
                    writeSidebarCollapsedState(false);
                    setSidebarCollapsed(false, false);
                  } else {
                    setSidebarCollapsed(readSidebarCollapsedState(), false);
                  }
                  syncMiniSidebarActions(sidebar);
                }

                function readInputSidebarCollapsedState() {
                  try {
                    const stored = parentWindow.localStorage.getItem(INPUT_SIDEBAR_STATE_KEY);
                    if (stored === '1') return true;
                    if (stored === '0') return false;
                  } catch (_) {}
                  return doc.documentElement.classList.contains(INPUT_SIDEBAR_COLLAPSED_CLASS);
                }

                function writeInputSidebarCollapsedState(collapsed) {
                  try {
                    parentWindow.localStorage.setItem(
                      INPUT_SIDEBAR_STATE_KEY,
                      collapsed ? '1' : '0'
                    );
                  } catch (_) {}
                }

                function inputSidebarToggleIcon(collapsed) {
                  const path = collapsed
                    ? 'M2.5 3.5L7 8l-4.5 4.5M7 3.5L11.5 8 7 12.5'
                    : 'M11.5 3.5L7 8l4.5 4.5M7 3.5L2.5 8 7 12.5';
                  return `
                    <svg class="tc-input-sidebar-toggle-icon" viewBox="0 0 14 16" aria-hidden="true">
                      <path d="${path}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"></path>
                    </svg>
                  `;
                }

                function inputSidebarMiniIcon(kind) {
                  const icons = {
                    taxonomy: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M12 5v4M6.5 9h11M6.5 9v4M17.5 9v4M12 9v4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                        <rect x="3.5" y="13" width="6" height="5.5" rx="1.4" fill="none" stroke="currentColor" stroke-width="1.7"/>
                        <rect x="9" y="13" width="6" height="5.5" rx="1.4" fill="none" stroke="currentColor" stroke-width="1.7"/>
                        <rect x="14.5" y="13" width="6" height="5.5" rx="1.4" fill="none" stroke="currentColor" stroke-width="1.7"/>
                      </svg>`,
                    company: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <circle cx="9" cy="9" r="3" fill="none" stroke="currentColor" stroke-width="1.8"/>
                        <circle cx="16.5" cy="10" r="2.35" fill="none" stroke="currentColor" stroke-width="1.6"/>
                        <path d="M3.8 19c.6-3 2.4-4.6 5.2-4.6s4.7 1.6 5.3 4.6M14.2 15.1c2.8-.5 5 .7 5.8 3.5" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                      </svg>`,
                    revenue: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M7 17V12M12 17V8M17 17V5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/>
                        <path d="M5 19h14" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                      </svg>`,
                    table: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <rect x="4" y="5" width="16" height="14" rx="2" fill="none" stroke="currentColor" stroke-width="1.8"/>
                        <path d="M4 10h16M9 5v14" fill="none" stroke="currentColor" stroke-width="1.7"/>
                      </svg>`,
                    quartile: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M5 18h14M7 15l3-4 3 2 4-6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
                        <circle cx="7" cy="15" r="1.2" fill="currentColor"/><circle cx="10" cy="11" r="1.2" fill="currentColor"/><circle cx="13" cy="13" r="1.2" fill="currentColor"/><circle cx="17" cy="7" r="1.2" fill="currentColor"/>
                      </svg>`,
                    globe: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <circle cx="12" cy="12" r="8" fill="none" stroke="currentColor" stroke-width="1.8"/>
                        <path d="M4.5 12h15M12 4c2 2.2 3 4.9 3 8s-1 5.8-3 8M12 4c-2 2.2-3 4.9-3 8s1 5.8 3 8" fill="none" stroke="currentColor" stroke-width="1.5"/>
                      </svg>`,
                    row: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <rect x="4" y="7" width="16" height="10" rx="2" fill="none" stroke="currentColor" stroke-width="1.8"/>
                        <path d="M8 7v10M12 7v10M16 7v10" fill="none" stroke="currentColor" stroke-width="1.5"/>
                      </svg>`,
                    sort: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M8 5v14M8 19l-3-3M8 19l3-3M16 19V5M16 5l-3 3M16 5l3 3" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>
                      </svg>`
                  };
                  return icons[kind] || icons.table;
                }

                function inputSidebarCardKey(card) {
                  if (!card || !card.classList) return '';
                  for (const className of card.classList) {
                    const prefix = 'st-key-sidebar_l1_input_';
                    if (className.startsWith(prefix) && className !== prefix + 'list') {
                      return className.slice(prefix.length);
                    }
                  }
                  return '';
                }

                function inputSidebarMiniMeta(key) {
                  const meta = {
                    taxonomy: { icon: 'taxonomy', label: 'Tax', name: 'Taxonomy' },
                    company_headers: { icon: 'company', label: 'Co', name: 'Company names' },
                    revenue_values: { icon: 'revenue', label: 'Rev', name: 'Revenue row' },
                    fte_values: { icon: 'table', label: 'FTE', name: 'Function + FTE / $B table' },
                    l0_quartile_table: { icon: 'quartile', label: 'L0', name: 'L0 quartiles' },
                    l1_quartile_table: { icon: 'quartile', label: 'L1', name: 'L1 quartiles' },
                    geocost_table: { icon: 'globe', label: 'Geo', name: 'Geo Cost table' },
                    geocost_data_row: { icon: 'row', label: 'Row', name: 'Geo Cost data row' }
                  };
                  return meta[key] || { icon: 'table', label: 'In', name: 'Input' };
                }

                function inputSidebarCards(panel) {
                  if (!panel) return [];
                  const seen = new Set();
                  const rows = [];
                  panel.querySelectorAll('[class*="st-key-sidebar_l1_input_"]').forEach((card) => {
                    if (card.classList.contains('st-key-sidebar_l1_input_list')) return;
                    const key = inputSidebarCardKey(card);
                    const button = card.querySelector('button');
                    if (!key || !button || seen.has(key)) return;
                    seen.add(key);
                    rows.push({ key, card, button });
                  });
                  return rows;
                }

                function ensureMiniInputSidebarActions(panel) {
                  if (!panel) return null;
                  let rail = doc.getElementById('tc-input-sidebar-mini-actions');
                  if (!rail || rail.parentElement !== panel) {
                    if (rail) rail.remove();
                    rail = doc.createElement('div');
                    rail.id = 'tc-input-sidebar-mini-actions';
                    rail.setAttribute('aria-label', 'Compact Inputs controls');
                    panel.appendChild(rail);
                  }

                  const rows = inputSidebarCards(panel);
                  const sortButton = panel.querySelector('.st-key-open_sort_sidebar_v77 button, .st-key-input_sort_launcher button');
                  const signature = rows.map((row) => row.key).join('|') + (sortButton ? '|sort' : '');

                  if (rail.dataset.signature !== signature) {
                    rail.dataset.signature = signature;
                    rail.innerHTML = '';
                    rows.forEach((row) => {
                      const meta = inputSidebarMiniMeta(row.key);
                      const mini = doc.createElement('button');
                      mini.type = 'button';
                      mini.className = 'tc-input-sidebar-mini-action';
                      mini.dataset.inputKey = row.key;
                      mini.innerHTML = `${inputSidebarMiniIcon(meta.icon)}<span class="tc-input-sidebar-mini-action-label">${meta.label}</span><span class="tc-input-sidebar-mini-status" aria-hidden="true"></span>`;
                      mini.setAttribute('aria-label', meta.name);
                      rail.appendChild(mini);
                    });
                    if (sortButton) {
                      const divider = doc.createElement('div');
                      divider.className = 'tc-input-sidebar-mini-divider';
                      divider.setAttribute('aria-hidden', 'true');
                      rail.appendChild(divider);
                      const sortMini = doc.createElement('button');
                      sortMini.type = 'button';
                      sortMini.className = 'tc-input-sidebar-mini-action is-sort';
                      sortMini.dataset.action = 'sort';
                      sortMini.innerHTML = `${inputSidebarMiniIcon('sort')}<span class="tc-input-sidebar-mini-action-label">Sort</span>`;
                      sortMini.setAttribute('aria-label', 'Company sort');
                      rail.appendChild(sortMini);
                    }
                  }

                  rows.forEach((row) => {
                    const mini = rail.querySelector(`[data-input-key="${row.key}"]`);
                    if (!mini) return;
                    const meta = inputSidebarMiniMeta(row.key);
                    const text = String(row.button.innerText || row.button.textContent || '').trim();
                    const signatureText = [
                      row.button.getAttribute('kind') || '',
                      row.button.getAttribute('data-testid') || '',
                      String(row.button.className || '')
                    ].join(' ').toLowerCase();
                    const mapped = /^\s*✓/.test(text);
                    const active = signatureText.includes('primary');
                    const disabled = row.button.disabled || row.button.getAttribute('aria-disabled') === 'true';
                    mini.classList.toggle('is-mapped', mapped);
                    mini.classList.toggle('is-active', active);
                    mini.disabled = disabled;
                    mini.setAttribute('aria-disabled', disabled ? 'true' : 'false');
                    mini.setAttribute('aria-label', `${meta.name}${mapped ? ' — mapped' : ' — select range'}${active ? ' — active' : ''}`);
                    mini.onclick = (event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      if (disabled) return;
                      row.button.click();
                    };
                  });

                  const sortMini = rail.querySelector('[data-action="sort"]');
                  if (sortMini) {
                    const liveSortButton = panel.querySelector('.st-key-open_sort_sidebar_v77 button, .st-key-input_sort_launcher button');
                    const disabled = !liveSortButton || liveSortButton.disabled || liveSortButton.getAttribute('aria-disabled') === 'true';
                    sortMini.disabled = disabled;
                    sortMini.setAttribute('aria-disabled', disabled ? 'true' : 'false');
                    sortMini.onclick = (event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      if (!disabled && liveSortButton) liveSortButton.click();
                    };
                  }
                  return rail;
                }

                function syncMiniInputSidebarActions(panel) {
                  const rail = ensureMiniInputSidebarActions(panel);
                  if (!rail) return;
                  rail.setAttribute('aria-hidden', readInputSidebarCollapsedState() ? 'false' : 'true');
                }

                function syncInputSidebarToggle(collapsed) {
                  const button = doc.getElementById('tc-input-sidebar-toggle');
                  if (button) {
                    const label = collapsed ? 'Expand Inputs panel' : 'Collapse Inputs panel';
                    button.dataset.collapsed = collapsed ? 'true' : 'false';
                    button.setAttribute('aria-label', label);
                    button.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
                    button.removeAttribute('title');
                    if (button.dataset.iconState !== button.dataset.collapsed) {
                      button.innerHTML = inputSidebarToggleIcon(collapsed);
                      button.dataset.iconState = button.dataset.collapsed;
                    }
                  }
                  const mini = doc.getElementById('tc-input-sidebar-mini');
                  if (mini) mini.setAttribute('aria-hidden', collapsed ? 'false' : 'true');
                  const actions = doc.getElementById('tc-input-sidebar-mini-actions');
                  if (actions) actions.setAttribute('aria-hidden', collapsed ? 'false' : 'true');
                }

                function setInputSidebarCollapsed(collapsed, persist = true) {
                  const root = doc.documentElement;
                  const body = doc.body;
                  const next = Boolean(collapsed);
                  const changed = root.classList.contains(INPUT_SIDEBAR_COLLAPSED_CLASS) !== next;
                  root.classList.toggle(INPUT_SIDEBAR_COLLAPSED_CLASS, next);
                  if (body) body.classList.toggle(INPUT_SIDEBAR_COLLAPSED_CLASS, next);
                  parentWindow.__tcInputSidebarCollapsed = next;
                  if (persist) writeInputSidebarCollapsedState(next);
                  syncInputSidebarToggle(next);
                  if (changed) announceSidebarResize();
                }

                function ensureInputSidebarToggle(panel) {
                  if (!panel) return;
                  if (!panel.id) panel.id = 'tc-input-sidebar';

                  let mini = doc.getElementById('tc-input-sidebar-mini');
                  if (!mini || mini.parentElement !== panel) {
                    if (mini) mini.remove();
                    mini = doc.createElement('div');
                    mini.id = 'tc-input-sidebar-mini';
                    mini.innerHTML = '<span class="tc-input-sidebar-mini-title">IN</span>';
                    mini.setAttribute('aria-label', 'Slide inputs');
                    panel.appendChild(mini);
                  }

                  ensureMiniInputSidebarActions(panel);

                  let button = doc.getElementById('tc-input-sidebar-toggle');
                  if (!button || button.parentElement !== panel) {
                    if (button) button.remove();
                    button = doc.createElement('button');
                    button.id = 'tc-input-sidebar-toggle';
                    button.type = 'button';
                    button.setAttribute('aria-controls', 'tc-input-sidebar');
                    panel.appendChild(button);
                  }
                  button.removeAttribute('title');
                  button.onclick = (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    setInputSidebarCollapsed(!readInputSidebarCollapsedState(), true);
                  };
                  setInputSidebarCollapsed(readInputSidebarCollapsedState(), false);
                }

                function applyInputSidebarState() {
                  const root = doc.documentElement;
                  const body = doc.body;
                  const panel = doc.querySelector('.st-key-l1_input_sidebar');
                  root.classList.toggle(INPUT_SIDEBAR_OPEN_CLASS, INPUT_SIDEBAR_ACTIVE);
                  if (body) body.classList.toggle(INPUT_SIDEBAR_OPEN_CLASS, INPUT_SIDEBAR_ACTIVE);

                  if (!INPUT_SIDEBAR_ACTIVE) {
                    root.classList.remove(INPUT_SIDEBAR_COLLAPSED_CLASS);
                    if (body) body.classList.remove(INPUT_SIDEBAR_COLLAPSED_CLASS);
                    const staleToggle = doc.getElementById('tc-input-sidebar-toggle');
                    if (staleToggle) staleToggle.remove();
                    const staleMini = doc.getElementById('tc-input-sidebar-mini');
                    if (staleMini) staleMini.remove();
                    const staleActions = doc.getElementById('tc-input-sidebar-mini-actions');
                    if (staleActions) staleActions.remove();
                    announceSidebarResize();
                    return;
                  }

                  if (panel) ensureInputSidebarToggle(panel);

                  let processedToken = '';
                  try {
                    processedToken = parentWindow.localStorage.getItem(INPUT_SIDEBAR_AUTOSTART_KEY) || '';
                  } catch (_) {}
                  if (INPUT_SIDEBAR_AUTOSTART_TOKEN && processedToken !== INPUT_SIDEBAR_AUTOSTART_TOKEN) {
                    setSidebarCollapsed(true, true);
                    setInputSidebarCollapsed(false, true);
                    try {
                      parentWindow.localStorage.setItem(
                        INPUT_SIDEBAR_AUTOSTART_KEY,
                        INPUT_SIDEBAR_AUTOSTART_TOKEN
                      );
                    } catch (_) {}
                  } else {
                    setInputSidebarCollapsed(readInputSidebarCollapsedState(), false);
                  }
                }


                function readSegmentSidebarCollapsedState() {
                  try {
                    const stored = parentWindow.localStorage.getItem(SEGMENT_SIDEBAR_STATE_KEY);
                    if (stored === '1') return true;
                    if (stored === '0') return false;
                  } catch (_) {}
                  return doc.documentElement.classList.contains(SEGMENT_SIDEBAR_COLLAPSED_CLASS);
                }

                function writeSegmentSidebarCollapsedState(collapsed) {
                  try {
                    parentWindow.localStorage.setItem(
                      SEGMENT_SIDEBAR_STATE_KEY,
                      collapsed ? '1' : '0'
                    );
                  } catch (_) {}
                }

                function segmentSidebarToggleIcon(collapsed) {
                  const path = collapsed
                    ? 'M2.5 3.5L7 8l-4.5 4.5M7 3.5L11.5 8 7 12.5'
                    : 'M11.5 3.5L7 8l4.5 4.5M7 3.5L2.5 8 7 12.5';
                  return `
                    <svg class="tc-segment-sidebar-toggle-icon" viewBox="0 0 14 16" aria-hidden="true">
                      <path d="${path}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"></path>
                    </svg>
                  `;
                }

                function syncSegmentSidebarToggle(collapsed) {
                  const button = doc.getElementById('tc-segment-sidebar-toggle');
                  if (button) {
                    const label = collapsed ? 'Expand company sort panel' : 'Collapse company sort panel';
                    button.dataset.collapsed = collapsed ? 'true' : 'false';
                    button.setAttribute('aria-label', label);
                    button.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
                    button.removeAttribute('title');
                    if (button.dataset.iconState !== button.dataset.collapsed) {
                      button.innerHTML = segmentSidebarToggleIcon(collapsed);
                      button.dataset.iconState = button.dataset.collapsed;
                    }
                  }
                  const mini = doc.getElementById('tc-segment-sidebar-mini');
                  if (mini) mini.setAttribute('aria-hidden', collapsed ? 'false' : 'true');
                  const actions = doc.getElementById('tc-segment-sidebar-mini-actions');
                  if (actions) actions.setAttribute('aria-hidden', collapsed ? 'false' : 'true');
                }

                function setSegmentSidebarCollapsed(collapsed, persist = true) {
                  const root = doc.documentElement;
                  const body = doc.body;
                  const next = Boolean(collapsed);
                  const changed = root.classList.contains(SEGMENT_SIDEBAR_COLLAPSED_CLASS) !== next;
                  root.classList.toggle(SEGMENT_SIDEBAR_COLLAPSED_CLASS, next);
                  if (body) body.classList.toggle(SEGMENT_SIDEBAR_COLLAPSED_CLASS, next);
                  parentWindow.__tcSegmentSidebarCollapsed = next;
                  if (persist) writeSegmentSidebarCollapsedState(next);
                  syncSegmentSidebarToggle(next);
                  if (changed) announceSidebarResize();
                }

                function segmentSidebarMiniIcon(kind) {
                  const icons = {
                    revenue_ascending: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M7 18V6M7 6l-3 3M7 6l3 3" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>
                        <path d="M12 17h7M12 12h5M12 7h3" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                      </svg>`,
                    revenue_descending: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <path d="M7 6v12M7 18l-3-3M7 18l3-3" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/>
                        <path d="M12 7h7M12 12h5M12 17h3" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
                      </svg>`,
                    sheet_order: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <rect x="5" y="4.5" width="14" height="15" rx="2" fill="none" stroke="currentColor" stroke-width="1.8"/>
                        <path d="M8 8h8M8 12h8M8 16h8" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"/>
                      </svg>`,
                    segment: `
                      <svg viewBox="0 0 24 24" aria-hidden="true">
                        <rect x="4" y="5" width="7" height="5.5" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.7"/>
                        <rect x="13" y="5" width="7" height="5.5" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.7"/>
                        <rect x="4" y="13.5" width="7" height="5.5" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.7"/>
                        <rect x="13" y="13.5" width="7" height="5.5" rx="1.5" fill="none" stroke="currentColor" stroke-width="1.7"/>
                      </svg>`
                  };
                  return icons[kind] || icons.sheet_order;
                }

                function segmentSidebarModeMeta(key) {
                  const meta = {
                    revenue_ascending: { label: 'Up', name: 'Increasing revenue' },
                    revenue_descending: { label: 'Dn', name: 'Decreasing revenue' },
                    sheet_order: { label: 'XL', name: 'Sheet order' },
                    segment: { label: 'Seg', name: 'Sort by segment' }
                  };
                  return meta[key] || { label: 'Sort', name: 'Company sort' };
                }

                function segmentSidebarModeButtons(panel) {
                  if (!panel) return [];
                  const keys = ['revenue_ascending', 'revenue_descending', 'sheet_order', 'segment'];
                  return keys.map((key) => {
                    const owner = panel.querySelector(`.st-key-sort_sidebar_mode_${key}`);
                    const button = owner ? owner.querySelector('button') : null;
                    return button ? { key, owner, button } : null;
                  }).filter(Boolean);
                }

                function ensureMiniSegmentSidebarActions(panel) {
                  if (!panel) return null;
                  let rail = doc.getElementById('tc-segment-sidebar-mini-actions');
                  if (!rail || rail.parentElement !== panel) {
                    if (rail) rail.remove();
                    rail = doc.createElement('div');
                    rail.id = 'tc-segment-sidebar-mini-actions';
                    rail.setAttribute('aria-label', 'Compact company sort controls');
                    panel.appendChild(rail);
                  }

                  const rows = segmentSidebarModeButtons(panel);
                  const signature = rows.map((row) => row.key).join('|');
                  if (rail.dataset.signature !== signature) {
                    rail.dataset.signature = signature;
                    rail.innerHTML = '';
                    rows.forEach((row) => {
                      const meta = segmentSidebarModeMeta(row.key);
                      const mini = doc.createElement('button');
                      mini.type = 'button';
                      mini.className = 'tc-segment-sidebar-mini-action';
                      mini.dataset.sortKey = row.key;
                      mini.innerHTML = `${segmentSidebarMiniIcon(row.key)}<span class="tc-segment-sidebar-mini-action-label">${meta.label}</span><span class="tc-segment-sidebar-mini-status" aria-hidden="true"></span>`;
                      mini.setAttribute('aria-label', meta.name);
                      rail.appendChild(mini);
                    });
                  }

                  rows.forEach((row) => {
                    const mini = rail.querySelector(`[data-sort-key="${row.key}"]`);
                    if (!mini) return;
                    const meta = segmentSidebarModeMeta(row.key);
                    const disabled = row.button.disabled || row.button.getAttribute('aria-disabled') === 'true';
                    const selected = (row.button.getAttribute('kind') || '').toLowerCase() === 'primary'
                      || /^\s*✓/.test(String(row.button.innerText || row.button.textContent || ''));
                    mini.classList.toggle('is-selected', selected);
                    mini.disabled = disabled;
                    mini.setAttribute('aria-disabled', disabled ? 'true' : 'false');
                    mini.setAttribute('aria-label', `${meta.name}${selected ? ' — selected' : ''}`);
                    mini.onclick = (event) => {
                      event.preventDefault();
                      event.stopPropagation();
                      if (!disabled) row.button.click();
                    };
                  });
                  return rail;
                }

                function syncMiniSegmentSidebarActions(panel) {
                  const rail = ensureMiniSegmentSidebarActions(panel);
                  if (!rail) return;
                  rail.setAttribute('aria-hidden', readSegmentSidebarCollapsedState() ? 'false' : 'true');
                }

                function ensureSegmentSidebarToggle(panel) {
                  if (!panel) return;
                  if (!panel.id) panel.id = 'tc-segment-sidebar';

                  let mini = doc.getElementById('tc-segment-sidebar-mini');
                  if (!mini || mini.parentElement !== panel) {
                    if (mini) mini.remove();
                    mini = doc.createElement('div');
                    mini.id = 'tc-segment-sidebar-mini';
                    mini.innerHTML = '<span class="tc-segment-sidebar-mini-title">SO</span>';
                    mini.setAttribute('aria-label', 'Company sort');
                    panel.appendChild(mini);
                  }

                  ensureMiniSegmentSidebarActions(panel);

                  let button = doc.getElementById('tc-segment-sidebar-toggle');
                  if (!button || button.parentElement !== panel) {
                    if (button) button.remove();
                    button = doc.createElement('button');
                    button.id = 'tc-segment-sidebar-toggle';
                    button.type = 'button';
                    button.setAttribute('aria-controls', 'tc-segment-sidebar');
                    panel.appendChild(button);
                  }
                  button.removeAttribute('title');
                  button.onclick = (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    setSegmentSidebarCollapsed(!readSegmentSidebarCollapsedState(), true);
                  };
                  setSegmentSidebarCollapsed(readSegmentSidebarCollapsedState(), false);
                }

                function applySegmentSidebarState() {
                  const root = doc.documentElement;
                  const body = doc.body;
                  const panel = doc.querySelector('.st-key-l1_segment_sidebar');
                  root.classList.toggle(SEGMENT_SIDEBAR_OPEN_CLASS, SEGMENT_SIDEBAR_ACTIVE);
                  if (body) body.classList.toggle(SEGMENT_SIDEBAR_OPEN_CLASS, SEGMENT_SIDEBAR_ACTIVE);

                  if (!SEGMENT_SIDEBAR_ACTIVE) {
                    root.classList.remove(SEGMENT_SIDEBAR_COLLAPSED_CLASS);
                    if (body) body.classList.remove(SEGMENT_SIDEBAR_COLLAPSED_CLASS);
                    const staleToggle = doc.getElementById('tc-segment-sidebar-toggle');
                    if (staleToggle) staleToggle.remove();
                    const staleMini = doc.getElementById('tc-segment-sidebar-mini');
                    if (staleMini) staleMini.remove();
                    const staleActions = doc.getElementById('tc-segment-sidebar-mini-actions');
                    if (staleActions) staleActions.remove();
                    announceSidebarResize();
                    return;
                  }

                  if (panel) ensureSegmentSidebarToggle(panel);

                  // The panel is emitted only after the Inputs launcher is used.
                  // From that point onward, retain the proven v59 independent
                  // collapse/expand behavior.
                  let processedToken = '';
                  try {
                    processedToken = parentWindow.localStorage.getItem(SEGMENT_SIDEBAR_AUTOSTART_KEY) || '';
                  } catch (_) {}
                  if (SEGMENT_SIDEBAR_AUTOSTART_TOKEN && processedToken !== SEGMENT_SIDEBAR_AUTOSTART_TOKEN) {
                    setSegmentSidebarCollapsed(false, true);
                    try {
                      parentWindow.localStorage.setItem(
                        SEGMENT_SIDEBAR_AUTOSTART_KEY,
                        SEGMENT_SIDEBAR_AUTOSTART_TOKEN
                      );
                    } catch (_) {}
                  } else {
                    setSegmentSidebarCollapsed(readSegmentSidebarCollapsedState(), false);
                  }
                }

                const TOOLTIP_DELAY_MS = 2000;

                function ensureDelayedTooltip() {
                  let bubble = doc.getElementById('tc-delayed-tooltip');
                  if (bubble) return bubble;

                  bubble = doc.createElement('div');
                  bubble.id = 'tc-delayed-tooltip';
                  bubble.setAttribute('role', 'tooltip');
                  bubble.setAttribute('aria-hidden', 'true');
                  Object.assign(bubble.style, {
                    position: 'fixed',
                    left: '0px',
                    top: '0px',
                    maxWidth: '420px',
                    minWidth: '220px',
                    padding: '12px 14px',
                    borderRadius: '12px',
                    border: '1px solid rgba(255,255,255,0.16)',
                    background: '#15171a',
                    color: '#ffffff',
                    boxShadow: '0 14px 34px rgba(0,0,0,0.34)',
                    fontFamily: 'Arial, Helvetica, sans-serif',
                    fontSize: '12px',
                    fontWeight: '650',
                    lineHeight: '1.45',
                    whiteSpace: 'pre-line',
                    overflowWrap: 'anywhere',
                    pointerEvents: 'none',
                    opacity: '0',
                    visibility: 'hidden',
                    transform: 'translateY(4px)',
                    transition: 'opacity 120ms ease, transform 120ms ease, visibility 120ms ease',
                    zIndex: '2147483647'
                  });

                  const accent = doc.createElement('div');
                  accent.className = 'tc-delayed-tooltip-accent';
                  Object.assign(accent.style, {
                    position: 'absolute',
                    left: '12px',
                    right: '12px',
                    top: '0',
                    height: '3px',
                    borderRadius: '0 0 3px 3px',
                    background: '#cc0000'
                  });

                  const body = doc.createElement('div');
                  body.className = 'tc-delayed-tooltip-body';

                  const arrow = doc.createElement('div');
                  arrow.className = 'tc-delayed-tooltip-arrow';
                  Object.assign(arrow.style, {
                    position: 'absolute',
                    width: '10px',
                    height: '10px',
                    background: '#15171a',
                    transform: 'rotate(45deg)',
                    border: '1px solid rgba(255,255,255,0.10)'
                  });

                  bubble.appendChild(accent);
                  bubble.appendChild(body);
                  bubble.appendChild(arrow);
                  doc.body.appendChild(bubble);
                  return bubble;
                }

                function hideDelayedTooltip() {
                  if (parentWindow.__tcTooltipTimer) {
                    parentWindow.clearTimeout(parentWindow.__tcTooltipTimer);
                    parentWindow.__tcTooltipTimer = null;
                  }
                  const bubble = doc.getElementById('tc-delayed-tooltip');
                  if (!bubble) return;
                  bubble.style.opacity = '0';
                  bubble.style.visibility = 'hidden';
                  bubble.style.transform = 'translateY(4px)';
                  bubble.setAttribute('aria-hidden', 'true');
                }

                function positionDelayedTooltip(bubble, target, placement) {
                  const rect = target.getBoundingClientRect();
                  const arrow = bubble.querySelector('.tc-delayed-tooltip-arrow');
                  const gap = 12;
                  const pad = 10;

                  // Measure after content is installed but before revealing it.
                  bubble.style.visibility = 'hidden';
                  bubble.style.opacity = '0';
                  bubble.style.display = 'block';
                  const bubbleRect = bubble.getBoundingClientRect();
                  let left = rect.right + gap;
                  let top = rect.top + (rect.height - bubbleRect.height) / 2;
                  let side = 'left';

                  if (placement === 'below') {
                    left = rect.right - bubbleRect.width;
                    top = rect.bottom + gap;
                    side = 'top';
                    if (top + bubbleRect.height > parentWindow.innerHeight - pad) {
                      top = rect.top - bubbleRect.height - gap;
                      side = 'bottom';
                    }
                  } else if (left + bubbleRect.width > parentWindow.innerWidth - pad) {
                    left = rect.left - bubbleRect.width - gap;
                    side = 'right';
                  }

                  left = Math.max(pad, Math.min(left, parentWindow.innerWidth - bubbleRect.width - pad));
                  top = Math.max(pad, Math.min(top, parentWindow.innerHeight - bubbleRect.height - pad));
                  bubble.style.left = `${Math.round(left)}px`;
                  bubble.style.top = `${Math.round(top)}px`;

                  if (arrow) {
                    arrow.style.left = '';
                    arrow.style.right = '';
                    arrow.style.top = '';
                    arrow.style.bottom = '';
                    if (side === 'left') {
                      arrow.style.left = '-6px';
                      arrow.style.top = `${Math.max(14, Math.min(bubbleRect.height - 24, rect.top + rect.height / 2 - top - 5))}px`;
                    } else if (side === 'right') {
                      arrow.style.right = '-6px';
                      arrow.style.top = `${Math.max(14, Math.min(bubbleRect.height - 24, rect.top + rect.height / 2 - top - 5))}px`;
                    } else if (side === 'top') {
                      arrow.style.top = '-6px';
                      arrow.style.right = '24px';
                    } else {
                      arrow.style.bottom = '-6px';
                      arrow.style.right = '24px';
                    }
                  }
                }

                function showDelayedTooltip(target) {
                  const text = String(target.dataset.tcTooltipText || '').trim();
                  if (!text) return;
                  const bubble = ensureDelayedTooltip();
                  const body = bubble.querySelector('.tc-delayed-tooltip-body');
                  if (body) body.textContent = text;
                  positionDelayedTooltip(bubble, target, target.dataset.tcTooltipPlacement || 'right');
                  bubble.style.visibility = 'visible';
                  bubble.style.opacity = '1';
                  bubble.style.transform = 'translateY(0)';
                  bubble.setAttribute('aria-hidden', 'false');
                }

                function bindDelayedTooltip(target, text, placement = 'right') {
                  if (!target) return;
                  target.dataset.tcTooltipText = String(text || '').trim();
                  target.dataset.tcTooltipPlacement = placement;
                  target.removeAttribute('title');
                  target.removeAttribute('aria-describedby');

                  if (target.dataset.tcTooltipWired === 'true') return;
                  target.dataset.tcTooltipWired = 'true';
                  target.addEventListener('mouseenter', () => {
                    hideDelayedTooltip();
                    if (!String(target.dataset.tcTooltipText || '').trim()) return;
                    parentWindow.__tcTooltipTimer = parentWindow.setTimeout(
                      () => showDelayedTooltip(target),
                      TOOLTIP_DELAY_MS
                    );
                  });
                  target.addEventListener('mouseleave', hideDelayedTooltip);
                  target.addEventListener('mousedown', hideDelayedTooltip);
                  target.addEventListener('click', hideDelayedTooltip);
                }

                function wireInputCards(panel) {
                  if (!panel) return;
                  panel.querySelectorAll('[class*="st-key-sidebar_l1_input_"]').forEach((card) => {
                    if (card.classList.contains('st-key-sidebar_l1_input_list')) return;

                    const helpSource = card.querySelector('.tc-input-help-source');
                    const helpText = helpSource
                      ? String(helpSource.textContent || '').trim()
                      : '';
                    const visibleButton = card.querySelector('button');
                    bindDelayedTooltip(card, helpText, 'right');
                    if (visibleButton) {
                      visibleButton.removeAttribute('title');
                      visibleButton.removeAttribute('aria-describedby');
                    }

                    if (card.dataset.tcCardWired === 'true') return;
                    card.dataset.tcCardWired = 'true';
                    card.addEventListener('click', (event) => {
                      if (event.target.closest('button')) return;
                      const button = card.querySelector('button');
                      if (!button || button.disabled) return;
                      event.preventDefault();
                      button.click();
                    });
                  });
                }

                function wireFinalizeTooltip() {
                  const holder = doc.querySelector('.st-key-workbook_review_action');
                  if (!holder) return;
                  const helpSource = holder.querySelector('.tc-finalize-help-source');
                  const helpText = helpSource
                    ? String(helpSource.textContent || '').trim()
                    : '';
                  bindDelayedTooltip(holder, helpText, 'below');
                  const button = holder.querySelector('button');
                  if (button) {
                    button.removeAttribute('title');
                    button.removeAttribute('aria-describedby');
                  }
                }

                function repairSortOrderButtons(panel) {
                  if (!panel) return;
                  const section = panel.querySelector('.st-key-sidebar_l1_sort_section');
                  if (!section) return;

                  // Streamlit inserts an element-container wrapper around each
                  // button. Styling only div.stButton changes the button surface
                  // but leaves those wrappers flush against each other. Apply the
                  // spacing to the keyed element containers themselves so the
                  // visual gap is stable across Streamlit releases and reruns.
                  const buttons = Array.from(section.querySelectorAll('button')).filter((button) => {
                    return button.id !== 'tc-sidebar-toggle' && button.id !== 'tc-input-sidebar-toggle';
                  });

                  // Keep the heading/copy in normal document flow and reserve
                  // clear space before the first option. This prevents the first
                  // button from covering the explanatory sentence after reruns.
                  const sortCopy = section.querySelector('.tc-input-sort-copy');
                  if (sortCopy) {
                    const copyOwner = sortCopy.closest('[data-testid="stElementContainer"]');
                    if (copyOwner) {
                      important(copyOwner, 'display', 'block');
                      important(copyOwner, 'height', 'auto');
                      important(copyOwner, 'min-height', '0');
                      important(copyOwner, 'overflow', 'visible');
                      important(copyOwner, 'margin', '0');
                      important(copyOwner, 'padding', '0');
                    }
                  }

                  section.querySelectorAll('[data-testid="stVerticalBlock"]').forEach((block) => {
                    important(block, 'gap', '0');
                  });

                  buttons.forEach((button, index) => {
                    let owner = button.closest('[class*="st-key-sidebar_l1_sort_"]');
                    if (owner && owner.classList.contains('st-key-sidebar_l1_sort_section')) {
                      owner = null;
                    }
                    owner = owner
                      || button.closest('[data-testid="stElementContainer"]')
                      || button.closest('div.stButton');

                    if (owner) {
                      important(owner, 'display', 'block');
                      important(owner, 'width', '100%');
                      important(owner, 'padding', '0');
                      important(owner, 'overflow', 'visible');
                      const topGap = index === 0 ? '16px' : '0';
                      const bottomGap = index < buttons.length - 1 ? '12px' : '0';
                      important(owner, 'margin', `${topGap} 0 ${bottomGap} 0`);
                    }

                    const buttonWrapper = button.closest('div.stButton');
                    if (buttonWrapper) {
                      important(buttonWrapper, 'display', 'block');
                      important(buttonWrapper, 'width', '100%');
                      important(buttonWrapper, 'margin', '0');
                      important(buttonWrapper, 'padding', '0');
                    }

                    important(button, 'width', '100%');
                    important(button, 'min-height', '46px');
                    important(button, 'margin', '0');
                    important(button, 'padding', '10px 14px');
                    important(button, 'border-radius', '12px');
                    important(button, 'box-shadow', button.getAttribute('kind') === 'primary'
                      ? '0 6px 16px rgba(35,0,0,0.24)'
                      : '0 4px 12px rgba(68,0,0,0.14)');
                  });
                }

                function repair() {
                  const sidebar = doc.querySelector('[data-testid="stSidebar"]');
                  if (!sidebar) return;
                  ensureSidebarToggle(sidebar);
                  syncMiniSidebarActions(sidebar);
                  applyInputSidebarState();
                  applySegmentSidebarState();
                  const segmentPanel = doc.querySelector('.st-key-l1_segment_sidebar');
                  syncMiniSegmentSidebarActions(segmentPanel);
                  const inputPanel = doc.querySelector('.st-key-l1_input_sidebar');
                  // Tooltip wiring is handled by dedicated zero-height injectors
                  // using serialized Python validation data. Do not inspect or
                  // mutate hidden content in the visible Streamlit DOM here.
                  syncMiniInputSidebarActions(inputPanel);
                  repairSortOrderButtons(inputPanel);

                  // The Sort order selector lives in the adjacent L1 rail,
                  // outside Streamlit's native sidebar. Repair its actual
                  // BaseWeb nodes as well so the selected value remains dark
                  // and readable on the white control surface across releases.
                  if (inputPanel) {
                    const inputPanelRoots = new Set([
                      ...inputPanel.querySelectorAll('.st-key-l1_sort_order'),
                      ...inputPanel.querySelectorAll('[data-testid="stSelectbox"]')
                    ]);
                    inputPanelRoots.forEach(repairRoot);
                  }

                  const roots = new Set([
                    ...sidebar.querySelectorAll('.st-key-workflow_mode'),
                    ...sidebar.querySelectorAll('.st-key-input_section'),
                    ...sidebar.querySelectorAll('[data-testid="stSelectbox"]'),
                    ...sidebar.querySelectorAll('[data-testid="stTextInput"]')
                  ]);
                  roots.forEach(repairRoot);
                  // repairRoot intentionally normalizes light controls to dark text.
                  // Re-apply the workflow multiselect's selected-chip theme afterwards.
                  repairWorkflowMultiselect(sidebar);

                  sidebar.querySelectorAll('[data-testid="stFileUploaderDropzone"] button').forEach((button) => {
                    forceText(button);
                    button.querySelectorAll('*').forEach((node) => forceText(node));
                  });

                  // Streamlit's sidebar theme can repaint secondary buttons as
                  // white-on-white. Target the rendered L1 mapping buttons
                  // directly instead of depending on release-specific wrappers.
                  const l1InputLabels = new Set([
                    'taxonomy', 'company names', 'revenue values', 'revenue row',
                    'fte / $b matrix', 'function + fte / $b table',
                    'l0 quartile table detector', 'l0 quartile table',
                    'l1 quartile table detector', 'l1 quartile table'
                  ]);
                  doc.querySelectorAll('[data-testid="stSidebar"] button, .st-key-l1_input_sidebar button').forEach((button) => {
                    if (button.id === 'tc-sidebar-toggle' || button.id === 'tc-input-sidebar-toggle') return;
                    if (button.closest('#tc-sidebar-mini-actions, #tc-input-sidebar-mini-actions')) return;
                    if (button.closest('.st-key-l1_sort_order, .st-key-sidebar_l1_sort_section, .st-key-input_sort_launcher')) return;
                    const owner = button.closest(
                      '[class*="st-key-sidebar_activate_l1_input_"], ' +
                      '[class*="st-key-sidebar_l1_input_"], ' +
                      '.st-key-l1_input_sidebar'
                    );
                    const visibleLabel = String(button.innerText || button.textContent || '')
                      .replace(/^\s*✓\s*/, '')
                      .trim()
                      .toLowerCase();
                    if (!owner && !l1InputLabels.has(visibleLabel)) return;
                    const buttonSignature = [
                      button.getAttribute('kind') || '',
                      button.getAttribute('data-testid') || '',
                      String(button.className || '')
                    ].join(' ').toLowerCase();
                    const isPrimary = buttonSignature.includes('primary');
                    const inInputPanel = Boolean(button.closest('.st-key-l1_input_sidebar'));
                    const textColor = isPrimary ? '#ffffff' : INK;
                    if (inInputPanel) {
                      const cardSurface = isPrimary ? '#101214' : '#ffffff';
                      const cardBorder = isPrimary ? '#101214' : 'rgba(255,255,255,0.82)';
                      important(button, 'background', cardSurface);
                      important(button, 'background-color', cardSurface);
                      important(button, 'background-image', 'none');
                      important(button, 'border-color', cardBorder);
                      important(
                        button,
                        'box-shadow',
                        isPrimary
                          ? '0 7px 18px rgba(35,0,0,0.25)'
                          : '0 5px 14px rgba(68,0,0,0.18)'
                      );
                    } else if (!isPrimary) {
                      important(button, 'background', SURFACE);
                      important(button, 'background-color', SURFACE);
                      important(button, 'background-image', 'none');
                      important(button, 'border-color', BORDER);
                    }
                    if (inInputPanel) {
                      // The earlier CSS-only typography patch was being defeated
                      // by this DOM repair routine, which wrote gray inline
                      // !important text after the stylesheet loaded. Apply the
                      // final L1 card typography here, on the actual rendered DOM.
                      const detailColor = isPrimary ? '#ff9a9a' : '#cc0000';
                      const titleColor = isPrimary ? '#ffffff' : '#15171a';
                      const card = button.closest('[class*="st-key-sidebar_l1_input_"]');

                      if (card && !card.classList.contains('st-key-sidebar_l1_input_list')) {
                        important(card, 'margin', '0 0 12px 0');
                        important(card, 'padding', '0');
                        important(card, 'overflow', 'visible');
                      }

                      important(button, 'min-height', '72px');
                      important(button, 'padding', '12px 14px');
                      important(button, 'display', 'flex');
                      important(button, 'align-items', 'center');
                      important(button, 'justify-content', 'center');

                      forceText(button, detailColor);
                      button.querySelectorAll('*').forEach((node) => forceText(node, detailColor));

                      button.querySelectorAll('p').forEach((paragraph) => {
                        important(paragraph, 'width', '100%');
                        important(paragraph, 'margin', '0');
                        important(paragraph, 'padding', '0');
                        important(paragraph, 'color', detailColor);
                        important(paragraph, '-webkit-text-fill-color', detailColor);
                        important(paragraph, 'font-size', '12.5px');
                        important(paragraph, 'line-height', '1.55');
                        important(paragraph, 'font-weight', '800');
                        important(paragraph, 'letter-spacing', '0');
                        important(paragraph, 'text-align', 'center');
                        important(paragraph, 'white-space', 'normal');
                        important(paragraph, 'overflow-wrap', 'anywhere');
                      });

                      button.querySelectorAll('strong').forEach((title) => {
                        forceText(title, titleColor);
                        important(title, 'display', 'inline');
                        important(title, 'font-size', '16px');
                        important(title, 'line-height', '1.28');
                        important(title, 'font-weight', '850');
                        important(title, 'letter-spacing', '-0.01em');
                      });
                    } else {
                      forceText(button, textColor);
                      button.querySelectorAll('*').forEach((node) => forceText(node, textColor));
                    }
                    if (inInputPanel) {
                      button.removeAttribute('title');
                      button.removeAttribute('data-tooltip');
                      button.removeAttribute('aria-describedby');
                    }
                  });

                  repairPortalOptions();
                }

                if (parentWindow.__tcControlColorRepairObserver) {
                  try { parentWindow.__tcControlColorRepairObserver.disconnect(); } catch (_) {}
                }
                if (parentWindow.__tcControlColorRepairTimer) {
                  try { parentWindow.clearInterval(parentWindow.__tcControlColorRepairTimer); } catch (_) {}
                }

                let scheduled = false;
                const scheduleRepair = () => {
                  if (scheduled) return;
                  scheduled = true;
                  parentWindow.requestAnimationFrame(() => {
                    scheduled = false;
                    repair();
                  });
                };

                repair();
                parentWindow.setTimeout(repair, 50);
                parentWindow.setTimeout(repair, 250);
                parentWindow.setTimeout(repair, 750);

                const observer = new parentWindow.MutationObserver(scheduleRepair);
                observer.observe(doc.body, {
                  subtree: true,
                  childList: true,
                  attributes: true,
                  attributeFilter: ['class', 'style', 'aria-disabled', 'aria-expanded']
                });
                parentWindow.__tcControlColorRepairObserver = observer;
                parentWindow.__tcControlColorRepairTimer = parentWindow.setInterval(repair, 1000);
              })();
            </script>
          </body>
        </html>
        """
    component_markup = component_markup.replace(
        "__TC_INPUT_SIDEBAR_ACTIVE__",
        "true" if input_sidebar_active else "false",
    ).replace(
        "__TC_INPUT_SIDEBAR_AUTOSTART_TOKEN__",
        str(autostart_token),
    ).replace(
        "__TC_SEGMENT_SIDEBAR_ACTIVE__",
        "true" if segment_sidebar_active else "false",
    ).replace(
        "__TC_SEGMENT_SIDEBAR_REQUESTED_OPEN__",
        "true" if segment_sidebar_requested_open else "false",
    ).replace(
        "__TC_SEGMENT_SIDEBAR_AUTOSTART_TOKEN__",
        sort_autostart_token,
    )
    components.html(
        component_markup,
        height=0,
        scrolling=False,
    )


def _restore_preview_path() -> Path | None:
    raw = st.session_state.get("preview_workbook_path")
    if not raw:
        return None
    candidate = Path(str(raw))
    return candidate.resolve() if candidate.is_file() else None


def _markdown_escape_button_text(value: object) -> str:
    """Escape dynamic workbook text before placing it in a button label."""
    return re.sub(r"([\\`*_{}\[\]()#+.!|>\-])", r"\\\1", str(value or ""))


def _render_universal_sort_controls(
    selected_sections: tuple[str, ...],
    *,
    compact: bool = False,
) -> None:
    """Render universal company-order choices as direct sidebar buttons.

    These buttons intentionally are *not* wrapped in a nested ``st.container``.
    On the installed Streamlit build that wrapper could stretch to the remaining
    rail height and push every segment mapping/editor block below the viewport.
    Direct keyed buttons keep normal document flow and preserve the stable v59
    segment display/editor beneath the selected mode.
    """
    if not set(selected_sections).intersection(SORTABLE_DECK_SECTIONS):
        return

    current_sort_order = str(
        st.session_state.get("l1_sort_order") or L1_SORT_ORDER_DEFAULT
    ).strip()
    if current_sort_order not in L1_SORT_ORDER_LABELS:
        current_sort_order = L1_SORT_ORDER_DEFAULT
        st.session_state["l1_sort_order"] = current_sort_order

    segment_profile = dict(st.session_state.get("l1_segment_profile") or {})
    segments_detected = bool(segment_profile.get("detected"))
    segment_count = int(segment_profile.get("segment_count") or 0)
    segment_company_count = int(segment_profile.get("company_count") or 0)
    segment_error = str(segment_profile.get("error") or "").strip()
    if current_sort_order == "segment" and not segments_detected:
        current_sort_order = L1_SORT_ORDER_DEFAULT
        st.session_state["l1_sort_order"] = current_sort_order

    for sort_key, sort_label in L1_SORT_ORDER_LABELS.items():
        selected = sort_key == current_sort_order
        if sort_key == "segment":
            if segments_detected:
                detail = (
                    f"{segment_count} segment{'s' if segment_count != 1 else ''} detected"
                    f" · {segment_company_count} companies"
                )
            elif segment_error:
                detail = "Segments not detected"
            else:
                detail = "Detecting segment row..."
            title = f"✓ {sort_label}" if selected else sort_label
            display_label = (
                f"**{_markdown_escape_button_text(title)}**  \n"
                f"{_markdown_escape_button_text(detail)}"
            )
            disabled = not segments_detected
        else:
            display_label = f"✓ {sort_label}" if selected else sort_label
            disabled = False

        if st.button(
            display_label,
            key=f"sort_sidebar_mode_{sort_key}",
            use_container_width=True,
            type="primary" if selected else "secondary",
            disabled=disabled,
        ):
            st.session_state["l1_sort_order"] = sort_key
            # Segment order has a dedicated detail workspace. Keep the sort
            # selection itself separate from whether that workspace is open so
            # Back can return to Company Sort without changing the chosen order.
            st.session_state["l1_segment_workspace_open"] = (sort_key == "segment")
            st.session_state["sort_sidebar_open"] = True
            _return_to_workbook_review()
            _rerun_streamlit()
            return

        if compact:
            # The CSS uses the keyed button wrapper for spacing. Keeping this
            # branch explicit documents that no separate spacer DOM is needed.
            pass

def _render_sort_sidebar_launcher(selected_sections: tuple[str, ...]) -> None:
    """Render an always-visible launcher at the bottom of the Inputs rail."""
    if not set(selected_sections).intersection(SORTABLE_DECK_SECTIONS):
        return

    current_sort_order = str(
        st.session_state.get("l1_sort_order") or L1_SORT_ORDER_DEFAULT
    ).strip()
    current_label = L1_SORT_ORDER_LABELS.get(
        current_sort_order,
        L1_SORT_ORDER_LABELS[L1_SORT_ORDER_DEFAULT],
    )
    panel_open = bool(st.session_state.get("sort_sidebar_open"))
    detail = current_label + (" · open" if panel_open else " · open settings")
    button_label = (
        f"**{_markdown_escape_button_text('Company sort')}**  \n"
        f"{_markdown_escape_button_text(detail)}"
    )

    st.markdown(
        '<div class="tc-input-sort-launch-divider" aria-hidden="true"></div>',
        unsafe_allow_html=True,
    )
    if st.button(
        button_label,
        key="open_sort_sidebar_v77",
        use_container_width=True,
        type="secondary",
        help=(
            "Open Company Sort settings. Company ordering is shared by Aura L0, Aura L1, "
            "and Geo Cost when those sections are selected. Current mode: " + current_label + "."
        ),
    ):
        _request_sort_sidebar_autostart()
        _rerun_streamlit()


def _render_aura_input_list(

    *,
    source_signature: str,
    workbook_ready: bool,
    selected_sections: tuple[str, ...],
) -> None:
    """Render the one standardized Inputs list for all selected slide families."""
    required_keys = standardized_input_keys_for_sections(selected_sections)
    mappings = get_standard_input_mappings(source_signature)

    st.markdown(
        '<div class="tc-input-sidebar-header">Inputs</div>',
        unsafe_allow_html=True,
    )

    if not required_keys:
        if set(selected_sections).intersection(SORTABLE_DECK_SECTIONS):
            message = (
                "No mapped slide inputs are required for this selection yet. "
                "Company order still applies to every selected bar-chart section."
            )
        else:
            message = "Summary-specific mapped inputs will be added in a later phase."
        st.markdown(
            '<div class="tc-sidebar-l1-status">'
            + html.escape(message)
            + '</div>',
            unsafe_allow_html=True,
        )
        return

    active_key = active_l1_input_key(source_signature) if source_signature else required_keys[0]
    if active_key not in required_keys:
        active_key = next(
            (key for key in required_keys if key not in mappings),
            required_keys[0],
        )
        if source_signature:
            set_active_l1_input_key(source_signature, active_key)

    validation = (
        validate_standard_input_mappings(
            mappings,
            selected_sections=selected_sections,
            source_signature=source_signature,
            taxonomy_profile=dict(st.session_state.get("taxonomy_sheet_profile") or {}),
            function_lookup_profile=dict(
                st.session_state.get("l1_function_lookup_profile") or {}
            ),
            master_label_profile=dict(
                st.session_state.get("l1_master_label_profile") or {}
            ),
            geocost_profile=dict(
                st.session_state.get("geocost_input_profile") or {}
            ),
            ignored_taxonomy_functions=get_ignored_aura_taxonomy_functions(source_signature),
        )
        if workbook_ready and source_signature
        else {
            "ready": False,
            "mapped_count": sum(1 for key in required_keys if mappings.get(key)),
            "required_count": len(required_keys),
            "errors": [],
        }
    )
    mapped_count = int(validation["mapped_count"])
    required_count = int(validation["required_count"])

    # Derive the company count directly from the resolved Company names mapping.
    # This is intentionally metadata/range based only: it never reopens Excel.
    company_mapping = dict(mappings.get("company_headers") or {})
    try:
        company_row_count = int(company_mapping.get("row_count") or 0)
        company_column_count = int(company_mapping.get("column_count") or 0)
    except (TypeError, ValueError):
        company_row_count = 0
        company_column_count = 0

    if company_row_count <= 0 or company_column_count <= 0:
        company_range = str(company_mapping.get("range_a1") or "").strip()
        if company_range:
            # Be tolerant of persisted values that accidentally include a sheet
            # qualifier or absolute-reference markers.
            bare_range = company_range.rsplit("!", 1)[-1].replace("$", "").strip()
            try:
                from openpyxl.utils.cell import range_boundaries

                min_col, min_row, max_col, max_row = range_boundaries(bare_range)
                company_row_count = max_row - min_row + 1
                company_column_count = max_col - min_col + 1
            except Exception:
                company_row_count = 0
                company_column_count = 0

    chart_company_count = (
        company_column_count
        if company_row_count == 1 and company_column_count > 0
        else 0
    )
    st.session_state["chart_company_count"] = chart_company_count

    validation_errors = [
        str(value).strip()
        for value in (validation.get("errors") or [])
        if str(value).strip()
    ]
    aura_diagnostics = dict(st.session_state.get("l1_auto_detection") or {})
    geocost_diagnostics = dict(st.session_state.get("geocost_auto_detection") or {})

    def _auto_error_for_input(input_key: str) -> str:
        if mappings.get(input_key):
            return ""
        if input_key in GEOCOST_INPUT_KEYS:
            return str(geocost_diagnostics.get("error") or "").strip()
        diagnostic_key = (
            "taxonomy_error"
            if input_key == "taxonomy"
            else (
                "l0_quartile_error"
                if input_key == "l0_quartile_table"
                else "aura_error"
            )
        )
        return str(aura_diagnostics.get(diagnostic_key) or "").strip()

    clicked_key = ""
    with st.container(key="sidebar_l1_input_list"):
        for field in STANDARD_INPUT_FIELDS:
            key = field["key"]
            if key not in required_keys:
                continue
            mapping = mappings.get(key)
            input_status = standard_input_status(
                field=field,
                mapping=(dict(mapping) if isinstance(mapping, dict) else None),
                workbook_ready=workbook_ready,
                validation=validation,
                auto_error=_auto_error_for_input(key),
            )
            if mapping:
                range_text = str(mapping.get("range_a1") or "")
                sheet_name = str(mapping.get("sheet_name") or "")
                qualified_range = f"{sheet_name}!{range_text}" if sheet_name else range_text
                origin = str(mapping.get("mapping_origin") or "").strip().lower()
                origin_label = "AUTO" if origin == "auto" else "MANUAL"
                title_text = f"{input_status['symbol']} {field['label']}"
                detail_text = f"{origin_label} · {qualified_range}"
            else:
                title_text = f"{input_status['symbol']} {field['label']}"
                detail_text = "Select range"

            button_label = (
                f"**{_markdown_escape_button_text(title_text)}**  \n"
                f"{_markdown_escape_button_text(detail_text)}"
            )
            with st.container(key=f"sidebar_l1_input_{key}"):
                if st.button(
                    button_label,
                    key=f"sidebar_activate_l1_input_{key}",
                    use_container_width=True,
                    type="primary" if key == active_key else "secondary",
                    disabled=not workbook_ready,
                ):
                    clicked_key = key

        # Compact information line below the mapped inputs. This is intentionally
        # not a button/card; it simply reports the resolved company count.
        if company_mapping and chart_company_count > 0:
            with st.container(key="sidebar_company_count_info"):
                st.markdown(
                    (
                        '<div class="tc-company-count-inline">'
                        '<span class="tc-company-count-dot" aria-hidden="true"></span>'
                        f'<span><strong>{chart_company_count}</strong> companies detected</span>'
                        '</div>'
                    ),
                    unsafe_allow_html=True,
                )

    if clicked_key and source_signature:
        if clicked_key != active_key:
            set_active_l1_input_key(source_signature, clicked_key)
        if clicked_key == "taxonomy":
            st.session_state["preview_source_role"] = "taxonomy"
        elif clicked_key in GEOCOST_INPUT_KEYS:
            st.session_state["preview_source_role"] = "geocost"
        else:
            st.session_state["preview_source_role"] = "aura"
        rerun = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)
        if callable(rerun):
            rerun()

    # Field status is rendered directly on each Inputs card. Detailed feedback
    # for the active input is shown in the workbook notification strip.

    # Company ordering is configured in a dedicated third rail. The Inputs
    # rail only exposes this launcher so mappings remain compact and the sort
    # controls can never be pushed off-screen by additional input cards.
    _render_sort_sidebar_launcher(selected_sections)


# Backward-compatible internal name retained for older callers.
def _render_l1_input_list(
    *,
    source_signature: str,
    workbook_ready: bool,
    selected_sections: tuple[str, ...] = ("l1",),
) -> None:
    _render_aura_input_list(
        source_signature=source_signature,
        workbook_ready=workbook_ready,
        selected_sections=selected_sections,
    )


def render_input_sidebar(controls: ControlPanelState) -> None:
    """Render the single standardized Inputs rail for any deck selection."""
    if not controls.selected_sections:
        return

    with st.container(key="l1_input_sidebar"):
        st.markdown(
            '<div class="tc-input-sidebar-accent" aria-hidden="true"></div>',
            unsafe_allow_html=True,
        )
        _render_aura_input_list(
            source_signature=str(st.session_state.get("uploaded_signature") or ""),
            workbook_ready=(
                controls.workbook_path is not None
                and controls.workbook_error is None
            ),
            selected_sections=controls.selected_sections,
        )





def _segment_numeric(value: Any) -> float | None:
    """Parse a revenue value used only for the segment-order preview."""
    if value is None:
        return None
    if isinstance(value, str):
        cleaned = (
            value.replace("$", "")
            .replace(",", "")
            .replace("\u00a0", "")
            .strip()
        )
        if cleaned.lower() in {"", "-", "na", "n/a", "nan", "none"}:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    try:
        return float(value)
    except Exception:
        return None


def _segment_assignments_from_profile_or_workbook(
    profile: dict[str, Any],
) -> list[dict[str, Any]]:
    """Return the detected Segment -> Company rows with a workbook fallback.

    The segment detector normally stores a complete ``assignments`` list in
    session state. Older/stale Streamlit state can retain the detected counts
    and segment spans while dropping that list, which previously left the third
    sidebar visibly blank. Rebuild the rows from the committed Company names and
    Revenue mappings when that happens so the UI never silently hides the
    detected allocation.
    """
    direct = [
        dict(item)
        for item in (profile.get("assignments") or [])
        if isinstance(item, dict) and str(item.get("company") or "").strip()
    ]
    if direct:
        return direct

    source_signature = str(st.session_state.get("uploaded_signature") or "")
    mappings = get_l1_input_mappings(source_signature)
    company_mapping = dict(mappings.get("company_headers") or {})
    revenue_mapping = dict(mappings.get("revenue_values") or {})
    workbook_raw = (
        st.session_state.get("preview_workbook_path")
        or st.session_state.get("uploaded_workbook_path")
    )
    if not workbook_raw:
        return []

    workbook_path = Path(str(workbook_raw))
    if not workbook_path.is_file():
        return []

    sheet_name = str(
        company_mapping.get("sheet_name")
        or profile.get("sheet_name")
        or ""
    ).strip()
    company_range = str(
        company_mapping.get("range_a1")
        or profile.get("company_range")
        or ""
    ).strip()
    if not sheet_name or not company_range:
        return []

    try:
        from openpyxl import load_workbook
        from openpyxl.utils.cell import range_boundaries

        min_col, min_row, max_col, max_row = range_boundaries(company_range)
        if min_row != max_row:
            return []

        try:
            workbook = load_workbook(
                workbook_path,
                read_only=True,
                data_only=True,
                keep_links=False,
            )
        except TypeError:
            workbook = load_workbook(
                workbook_path,
                read_only=True,
                data_only=True,
            )

        try:
            if sheet_name not in workbook.sheetnames:
                return []
            worksheet = workbook[sheet_name]
            companies = [
                str(worksheet.cell(min_row, column).value or "").strip()
                for column in range(min_col, max_col + 1)
            ]
            if any(not company for company in companies):
                return []

            revenues: list[float | None] = [None] * len(companies)
            revenue_sheet = str(revenue_mapping.get("sheet_name") or "").strip()
            revenue_range = str(revenue_mapping.get("range_a1") or "").strip()
            if revenue_sheet in workbook.sheetnames and revenue_range:
                rev_min_col, rev_min_row, rev_max_col, rev_max_row = range_boundaries(
                    revenue_range
                )
                if rev_min_row == rev_max_row:
                    raw_values = [
                        workbook[revenue_sheet].cell(rev_min_row, column).value
                        for column in range(rev_min_col, rev_max_col + 1)
                    ]
                    if len(raw_values) == len(companies) + 1:
                        raw_values = raw_values[1:]
                    if len(raw_values) == len(companies):
                        revenues = [_segment_numeric(value) for value in raw_values]

            group_by_column: dict[int, dict[str, Any]] = {}
            for group in profile.get("segments") or []:
                if not isinstance(group, dict):
                    continue
                try:
                    start_column = int(group.get("start_column") or 0)
                    end_column = int(group.get("end_column") or 0)
                except Exception:
                    continue
                for column in range(start_column, end_column + 1):
                    group_by_column.setdefault(column, group)

            rebuilt: list[dict[str, Any]] = []
            for index, (column, company) in enumerate(
                zip(range(min_col, max_col + 1), companies)
            ):
                group = group_by_column.get(column) or {}
                segment = "Client" if index == 0 else (
                    str(group.get("segment") or "Unassigned").strip()
                    or "Unassigned"
                )
                rebuilt.append(
                    {
                        "company_index": index,
                        "company": company,
                        "segment": segment,
                        "segment_color": str(group.get("segment_color") or ""),
                        "source_column": column,
                        "revenue": revenues[index],
                    }
                )
            return rebuilt
        finally:
            workbook.close()
    except Exception:
        return []


def _ordered_segment_assignments(
    profile: dict[str, Any],
    direction: str,
) -> list[dict[str, Any]]:
    """Return client first, then source-order segments with revenue-sorted peers."""
    assignments = _segment_assignments_from_profile_or_workbook(profile)
    if not assignments:
        return []
    assignments.sort(key=lambda item: int(item.get("company_index") or 0))
    client = assignments[0]
    peers = assignments[1:]
    segment_order = [
        str(value)
        for value in (profile.get("segment_order") or [])
        if str(value).strip()
    ]
    if not segment_order:
        seen: set[str] = set()
        for item in peers:
            segment = str(item.get("segment") or "Unassigned").strip() or "Unassigned"
            if segment not in seen:
                seen.add(segment)
                segment_order.append(segment)

    descending = direction != "revenue_ascending"

    def peer_key(item: dict[str, Any]) -> tuple[int, float, int]:
        value = _segment_numeric(item.get("revenue"))
        source_index = int(item.get("company_index") or 0)
        if value is None:
            return (1, 0.0, source_index)
        return (0, -value if descending else value, source_index)

    ordered = [client]
    for segment in segment_order:
        segment_items = [
            item
            for item in peers
            if (str(item.get("segment") or "Unassigned").strip() or "Unassigned")
            == segment
        ]
        ordered.extend(sorted(segment_items, key=peer_key))

    included = {int(item.get("company_index") or 0) for item in ordered}
    remainder = [
        item
        for item in peers
        if int(item.get("company_index") or 0) not in included
    ]
    ordered.extend(sorted(remainder, key=peer_key))
    return ordered


def _render_segment_mapping_component(ordered: list[dict[str, Any]]) -> None:
    """Render the detected allocation as a stable list capped at 12 visible rows."""

    def readable_text_color(hex_color: str) -> str:
        try:
            red = int(hex_color[0:2], 16)
            green = int(hex_color[2:4], 16)
            blue = int(hex_color[4:6], 16)
        except Exception:
            return "#ffffff"
        luminance = (0.299 * red) + (0.587 * green) + (0.114 * blue)
        return "#171717" if luminance >= 158 else "#ffffff"

    row_markup: list[str] = []
    for row_index, item in enumerate(ordered):
        segment = str(item.get("segment") or "Unassigned").strip() or "Unassigned"
        company = str(item.get("company") or "").strip()
        color = str(item.get("segment_color") or "").strip().lstrip("#")

        if row_index == 0:
            background = "#CC0000"
            foreground = "#FFFFFF"
        elif re.fullmatch(r"[0-9A-Fa-f]{6}", color):
            background = f"#{color.upper()}"
            foreground = readable_text_color(color.upper())
        else:
            background = "#5C5C5C"
            foreground = "#FFFFFF"

        row_markup.append(
            '<div class="mapping-item' + (' client' if row_index == 0 else '') + '">'
            '<div class="segment-label" style="background:' + background + ';color:' + foreground + ';">'
            + html.escape(segment)
            + '</div>'
            '<div class="company-label">'
            + html.escape(company)
            + '</div>'
            '</div>'
        )

    visible_row_limit = 12
    visible_rows = min(max(1, len(ordered)), visible_row_limit)
    row_pitch_px = 41
    list_height_px = visible_rows * row_pitch_px
    list_overflow = "auto" if len(ordered) > visible_row_limit else "visible"

    mapping_component = f"""<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<style>
  :root {{ color-scheme: light; }}
  * {{ box-sizing: border-box; }}
  html, body {{
    width: 100%;
    height: 100%;
    margin: 0;
    padding: 0;
    overflow: hidden;
    background: transparent;
    font-family: Arial, Helvetica, sans-serif;
  }}
  .mapping-shell {{
    width: 100%;
    margin: 0;
    padding: 12px 10px 12px;
    border: 1px solid rgba(255,255,255,0.22);
    border-radius: 12px;
    background: rgba(28,28,28,0.20);
  }}
  .audit-title-row {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    margin: 0 0 11px;
  }}
  .audit-title {{
    min-width: 0;
    color: #ffffff;
    font-size: 13px;
    line-height: 1.28;
    font-weight: 850;
    letter-spacing: 0.005em;
  }}
  .audit-count {{
    flex: 0 0 auto;
    padding: 5px 9px;
    border: 1px solid rgba(255,255,255,0.34);
    border-radius: 999px;
    background: rgba(20,20,20,0.42);
    color: #ffffff;
    font-size: 10.5px;
    line-height: 1;
    font-weight: 850;
  }}
  .mapping-list {{
    display: flex;
    flex-direction: column;
    gap: 7px;
    width: 100%;
    height: {list_height_px}px;
    max-height: {list_height_px}px;
    padding: 0 5px 0 0;
    overflow-x: hidden;
    overflow-y: {list_overflow};
    scrollbar-width: thin;
    scrollbar-color: #9B9B9B rgba(0,0,0,0.10);
  }}
  .mapping-list::-webkit-scrollbar {{ width: 8px; }}
  .mapping-list::-webkit-scrollbar-track {{
    border-radius: 999px;
    background: rgba(0,0,0,0.10);
  }}
  .mapping-list::-webkit-scrollbar-thumb {{
    border: 2px solid transparent;
    border-radius: 999px;
    background: #9B9B9B;
    background-clip: padding-box;
  }}
  .mapping-item {{
    display: grid;
    grid-template-columns: minmax(76px, 0.72fr) minmax(0, 1.65fr);
    align-items: center;
    gap: 10px;
    flex: 0 0 34px;
    width: 100%;
    min-height: 34px;
    padding: 5px 10px 5px 7px;
    border: 1px solid #D0D0D0;
    border-radius: 10px;
    background: #F7F7F7;
    box-shadow: 0 2px 6px rgba(0,0,0,0.10);
  }}
  .mapping-item.client {{
    background: #E7E7E7;
    border-color: #C5C5C5;
  }}
  .segment-label {{
    min-width: 0;
    padding: 6px 8px;
    border-radius: 7px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    text-align: center;
    font-size: 10.5px;
    line-height: 1.15;
    font-weight: 850;
    letter-spacing: 0.02em;
  }}
  .company-label {{
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    color: #1F1F1F;
    font-size: 11.5px;
    line-height: 1.2;
    font-weight: 740;
  }}
</style>
</head>
<body>
  <div class="mapping-shell">
    <div class="audit-title-row">
      <div class="audit-title">Segment–Company mapping</div>
      <div class="audit-count">{len(ordered)}</div>
    </div>
    <div class="mapping-list">
      {''.join(row_markup)}
    </div>
  </div>
</body>
</html>"""
    # The list has its own scrollbar only when more than 12 assignments exist.
    # Keeping the iframe itself fixed prevents the rest of the third sidebar
    # from being pushed below the viewport.
    components.html(
        mapping_component,
        height=max(150, 78 + list_height_px),
        scrolling=False,
    )

def _rerun_streamlit() -> None:
    rerun = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)
    if callable(rerun):
        rerun()




_SEGMENT_ASSIGNMENT_COMPONENT = None
_SEGMENT_ASSIGNMENT_COMPONENT_BUILD = "thinkcell-segment-assignment-editor-v2026-09-17-r1"
_SEGMENT_ASSIGNMENT_COMPONENT_HTML = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>ThinkCell segment assignment editor</title>
  <style>
    :root { color-scheme: light; }
    * { box-sizing: border-box; }
    html, body {
      width: 100%;
      margin: 0;
      padding: 0;
      overflow: hidden;
      background: transparent;
      font-family: "Segoe UI", Inter, Arial, sans-serif;
      color: #17191c;
    }
    #root { width: 100%; }
    .editor-shell {
      width: 100%;
      padding: 10px;
      border: 1px solid rgba(255,255,255,0.25);
      border-radius: 12px;
      background: rgba(24,24,24,0.18);
      box-shadow: inset 0 1px 0 rgba(255,255,255,0.06);
    }
    .assignment-list {
      width: 100%;
      display: flex;
      flex-direction: column;
      gap: 8px;
      padding: 0 7px 0 0;
      overflow-x: hidden;
      overflow-y: auto;
      scrollbar-width: thin;
      scrollbar-color: #b8b8b8 rgba(0,0,0,0.13);
    }
    .assignment-list::-webkit-scrollbar { width: 8px; }
    .assignment-list::-webkit-scrollbar-track {
      border-radius: 999px;
      background: rgba(0,0,0,0.13);
    }
    .assignment-list::-webkit-scrollbar-thumb {
      border: 2px solid transparent;
      border-radius: 999px;
      background: #b8b8b8;
      background-clip: padding-box;
    }
    .assignment-row {
      flex: 0 0 44px;
      display: grid;
      grid-template-columns: minmax(0, 0.42fr) minmax(0, 0.58fr);
      align-items: stretch;
      gap: 8px;
      width: 100%;
      min-height: 44px;
    }
    .segment-input,
    .fixed-segment,
    .company-card {
      width: 100%;
      min-width: 0;
      height: 44px;
      border-radius: 9px;
      font-size: 12px;
      line-height: 1.2;
      font-weight: 800;
    }
    .segment-input {
      appearance: none;
      padding: 0 12px;
      border: 1px solid #c8c8c8;
      outline: none;
      background: #ffffff;
      color: #17191c;
      box-shadow: 0 1px 2px rgba(0,0,0,0.08);
    }
    .segment-input:focus {
      border-color: #cc0000;
      box-shadow: 0 0 0 2px rgba(204,0,0,0.16);
    }
    .fixed-segment,
    .company-card {
      display: flex;
      align-items: center;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      padding: 0 12px;
    }
    .fixed-segment {
      justify-content: center;
      border: 1px solid #e34848;
      background: #cc0000;
      color: #ffffff;
    }
    .company-card {
      border: 1px solid #c8c8c8;
      background: #f7f7f7;
      color: #1e1e1e;
      box-shadow: 0 1px 2px rgba(0,0,0,0.06);
    }
    .action-row {
      display: grid;
      grid-template-columns: minmax(0, 1.22fr) minmax(0, 1fr) minmax(0, 0.78fr);
      gap: 9px;
      width: 100%;
      margin-top: 12px;
    }
    .action-button {
      min-width: 0;
      min-height: 42px;
      padding: 8px 9px;
      border-radius: 9px;
      border: 1px solid rgba(255,255,255,0.46);
      background: #ffffff;
      color: #333333;
      font-size: 11.5px;
      line-height: 1.2;
      font-weight: 820;
      cursor: pointer;
      white-space: normal;
    }
    .action-button:hover { filter: brightness(0.97); }
    .action-button.primary {
      border-color: #cc0000;
      background: #cc0000;
      color: #ffffff;
    }
    .editor-status {
      display: none;
      margin-top: 9px;
      padding: 8px 10px;
      border-left: 3px solid #ff8b91;
      border-radius: 7px;
      background: rgba(84,28,32,0.82);
      color: #ffe6e8;
      font-size: 11px;
      line-height: 1.35;
      font-weight: 700;
    }
    .editor-status.is-visible { display: block; }
  </style>
</head>
<body>
  <div id="root"></div>
  <script>
    (() => {
      const root = document.getElementById("root");
      let args = {};

      function sendMessage(type, data = {}) {
        window.parent.postMessage(
          Object.assign({ isStreamlitMessage: true, type }, data),
          "*"
        );
      }

      function setFrameHeight(height) {
        sendMessage("streamlit:setFrameHeight", { height });
      }

      function setComponentValue(value) {
        sendMessage("streamlit:setComponentValue", {
          value,
          dataType: "json",
        });
      }

      function emit(action, assignments = {}) {
        setComponentValue({
          event_id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
          action,
          assignments,
        });
      }

      function showStatus(message) {
        const status = root.querySelector(".editor-status");
        if (!status) return;
        status.textContent = String(message || "");
        status.classList.toggle("is-visible", Boolean(message));
      }

      function render() {
        const rows = Array.isArray(args.rows) ? args.rows : [];
        const visibleLimit = Math.max(1, Math.min(12, Number(args.visible_limit || 12)));
        root.innerHTML = "";

        const shell = document.createElement("div");
        shell.className = "editor-shell";

        const list = document.createElement("div");
        list.className = "assignment-list";
        const visibleRows = Math.max(1, Math.min(rows.length || 1, visibleLimit));
        const rowHeight = 44;
        const rowGap = 8;
        const listHeight = (visibleRows * rowHeight) + (Math.max(0, visibleRows - 1) * rowGap);
        list.style.height = `${listHeight}px`;
        list.style.maxHeight = `${listHeight}px`;

        rows.forEach((row, index) => {
          const item = document.createElement("div");
          item.className = "assignment-row";

          const companyKey = String(row.company_key || "");
          const company = String(row.company || "");
          const segment = String(row.segment || "");
          const isClient = Boolean(row.is_client) || index === 0;

          if (isClient) {
            const fixed = document.createElement("div");
            fixed.className = "fixed-segment";
            fixed.textContent = "Client";
            item.appendChild(fixed);
          } else {
            const input = document.createElement("input");
            input.type = "text";
            input.className = "segment-input";
            input.value = segment;
            input.maxLength = 40;
            input.autocomplete = "off";
            input.spellcheck = false;
            input.dataset.companyKey = companyKey;
            input.setAttribute("aria-label", `Segment for ${company}`);
            input.addEventListener("input", () => showStatus(""));
            item.appendChild(input);
          }

          const companyCard = document.createElement("div");
          companyCard.className = "company-card";
          companyCard.textContent = company;
          companyCard.title = company;
          item.appendChild(companyCard);
          list.appendChild(item);
        });

        const actions = document.createElement("div");
        actions.className = "action-row";

        const save = document.createElement("button");
        save.type = "button";
        save.className = "action-button primary";
        save.textContent = "Save assignments";
        save.addEventListener("click", () => {
          const assignments = {};
          const invalid = [];
          list.querySelectorAll(".segment-input").forEach((input) => {
            const companyKey = String(input.dataset.companyKey || "");
            const value = String(input.value || "").replace(/\s+/g, " ").trim();
            if (!value || ["client", "unassigned"].includes(value.toLowerCase())) {
              invalid.push(companyKey);
            }
            assignments[companyKey] = value;
          });
          if (invalid.length) {
            showStatus("Every peer needs a segment. 'Client' and 'Unassigned' are reserved.");
            return;
          }
          showStatus("");
          emit("save", assignments);
        });

        const reset = document.createElement("button");
        reset.type = "button";
        reset.className = "action-button";
        reset.textContent = "Reset detected";
        reset.addEventListener("click", () => emit("reset"));

        const close = document.createElement("button");
        close.type = "button";
        close.className = "action-button";
        close.textContent = "Close";
        close.addEventListener("click", () => emit("close"));

        actions.appendChild(save);
        actions.appendChild(reset);
        actions.appendChild(close);

        const status = document.createElement("div");
        status.className = "editor-status";
        status.setAttribute("role", "alert");

        shell.appendChild(list);
        shell.appendChild(actions);
        shell.appendChild(status);
        root.appendChild(shell);

        const frameHeight = listHeight + 10 + 12 + 42 + 10 + 2;
        setFrameHeight(Math.max(160, frameHeight));
      }

      window.addEventListener("message", (event) => {
        if (!event.data || event.data.type !== "streamlit:render") return;
        args = event.data.args || {};
        render();
      });

      sendMessage("streamlit:componentReady", { apiVersion: 1 });
      setFrameHeight(690);
    })();
  </script>
</body>
</html>
"""


def _get_segment_assignment_component():
    """Return the isolated, fixed-height per-company assignment editor."""
    global _SEGMENT_ASSIGNMENT_COMPONENT
    if _SEGMENT_ASSIGNMENT_COMPONENT is not None:
        return _SEGMENT_ASSIGNMENT_COMPONENT

    component_dir = (
        Path(tempfile.gettempdir())
        / "ThinkCellSlideMaker"
        / _SEGMENT_ASSIGNMENT_COMPONENT_BUILD
    )
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        current = index_path.read_text(encoding="utf-8")
    except OSError:
        current = ""
    if current != _SEGMENT_ASSIGNMENT_COMPONENT_HTML:
        temporary_path = index_path.with_suffix(".html.writing")
        temporary_path.write_text(
            _SEGMENT_ASSIGNMENT_COMPONENT_HTML,
            encoding="utf-8",
        )
        temporary_path.replace(index_path)

    _SEGMENT_ASSIGNMENT_COMPONENT = components.declare_component(
        "thinkcell_segment_assignment_editor_v1",
        path=str(component_dir),
    )
    return _SEGMENT_ASSIGNMENT_COMPONENT


def _segment_assignment_editor_open_key(source_signature: str) -> str:
    return source_signature or "__no_workbook__"


def _segment_assignment_editor_is_open(source_signature: str) -> bool:
    root = st.session_state.get("l1_segment_assignment_editor_open") or {}
    if not isinstance(root, dict):
        return False
    return bool(root.get(_segment_assignment_editor_open_key(source_signature), False))


def _set_segment_assignment_editor_open(
    source_signature: str,
    is_open: bool,
) -> None:
    root = dict(st.session_state.get("l1_segment_assignment_editor_open") or {})
    root[_segment_assignment_editor_open_key(source_signature)] = bool(is_open)
    st.session_state["l1_segment_assignment_editor_open"] = root


def _render_segment_assignment_editor_stable(
    *,
    profile: dict[str, Any],
    source_signature: str,
) -> None:
    """Edit peer-company segments in one isolated 12-row scrolling component."""
    assignments = _segment_assignments_from_profile_or_workbook(profile)
    assignments = sorted(
        assignments,
        key=lambda item: int(item.get("company_index") or 0),
    )
    if not assignments:
        return

    is_open = _segment_assignment_editor_is_open(source_signature)
    st.markdown(
        '<div class="tc-segment-section-divider"></div>'
        '<div class="tc-segment-assignment-heading-block">'
        '<div class="tc-segment-section-heading">Company assignments</div>'
        '<div class="tc-segment-section-copy">Review the detected allocation above, then edit any peer segment here when needed. Matching names are grouped together; the client remains fixed.</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    if not is_open:
        manual_count = len(profile.get("segment_assignment_overrides") or {})
        toggle_label = "Edit individual assignments"
        if manual_count:
            toggle_label += f" · {manual_count} changed"
        if st.button(
            toggle_label,
            key="l1_segment_assignment_toggle",
            use_container_width=True,
            type="secondary",
        ):
            _set_segment_assignment_editor_open(source_signature, True)
            _rerun_streamlit()
        return

    st.markdown(
        '<div class="tc-segment-assignment-editor-note">'
        'Change only the left-hand Segment value. The Company value on the right is read-only.'
        '</div>',
        unsafe_allow_html=True,
    )

    current_overrides = get_l1_segment_assignment_overrides(source_signature)
    rows: list[dict[str, Any]] = []
    for index, item in enumerate(assignments):
        company = str(item.get("company") or "").strip()
        detected = str(
            item.get("detected_segment")
            or item.get("segment")
            or "Unassigned"
        ).strip() or "Unassigned"
        company_key = _segment_company_key(company)
        current_value = (
            "Client"
            if index == 0
            else (
                current_overrides.get(company_key)
                or str(item.get("segment") or detected).strip()
                or detected
            )
        )
        rows.append(
            {
                "company_key": company_key,
                "company": company,
                "segment": current_value,
                "detected_segment": detected,
                "is_client": index == 0,
            }
        )

    component = _get_segment_assignment_component()
    event = component(
        rows=rows,
        visible_limit=12,
        key=f"l1_segment_assignment_editor_{source_signature[:12] or 'workbook'}",
        default=None,
    )
    if not isinstance(event, dict):
        return

    event_id = str(event.get("event_id") or "")
    event_state_key = (
        "l1_segment_assignment_editor_last_event_"
        + (source_signature[:16] or "workbook")
    )
    if not event_id or st.session_state.get(event_state_key) == event_id:
        return
    st.session_state[event_state_key] = event_id

    action = str(event.get("action") or "").strip().lower()
    peer_assignments = assignments[1:]

    if action == "save":
        raw_assignments = event.get("assignments") or {}
        if not isinstance(raw_assignments, dict):
            st.error("The segment editor returned an invalid assignment payload.")
            return

        cleaned = {
            str(company_key): re.sub(r"\s+", " ", str(value or "")).strip()
            for company_key, value in raw_assignments.items()
        }
        blank_companies = [
            str(item.get("company") or "").strip()
            for item in peer_assignments
            if not cleaned.get(_segment_company_key(item.get("company")), "")
        ]
        reserved_companies = [
            str(item.get("company") or "").strip()
            for item in peer_assignments
            if cleaned.get(_segment_company_key(item.get("company")), "").casefold()
            in {"client", "unassigned"}
        ]
        if blank_companies:
            st.error("Every peer needs a segment: " + ", ".join(blank_companies))
            return
        if reserved_companies:
            st.error(
                "Client and Unassigned are reserved. Update: "
                + ", ".join(reserved_companies)
            )
            return

        overrides: dict[str, str] = {}
        for item in peer_assignments:
            company = str(item.get("company") or "").strip()
            company_key = _segment_company_key(company)
            detected = str(
                item.get("detected_segment")
                or item.get("segment")
                or "Unassigned"
            ).strip() or "Unassigned"
            selected = cleaned[company_key]
            if selected != detected:
                overrides[company_key] = selected

        set_l1_segment_assignment_overrides(source_signature, overrides)
        st.session_state["l1_segment_profile"] = (
            apply_l1_segment_assignment_overrides(profile, overrides)
        )
        st.session_state["l1_segment_profile_signature"] = ""
        _set_segment_assignment_editor_open(source_signature, False)
        _rerun_streamlit()
        return

    if action == "reset":
        set_l1_segment_assignment_overrides(source_signature, {})
        st.session_state["l1_segment_profile"] = (
            apply_l1_segment_assignment_overrides(profile, {})
        )
        st.session_state["l1_segment_profile_signature"] = ""
        _set_segment_assignment_editor_open(source_signature, False)
        _rerun_streamlit()
        return

    if action == "close":
        _set_segment_assignment_editor_open(source_signature, False)
        _rerun_streamlit()


def render_segment_sidebar(controls: ControlPanelState) -> None:
    """Render the dedicated universal company-sort rail.

    In Segment mode the proven v59 mapping and assignment editor are rendered
    *before* the sort-mode picker. This makes selecting ``Sort by segment``
    immediately open the requested display/edit workspace instead of leaving a
    blank rail beneath a stretching mode-selector wrapper.
    """
    sortable_selected = bool(
        set(controls.selected_sections).intersection(SORTABLE_DECK_SECTIONS)
    )
    if not sortable_selected:
        return
    if not bool(st.session_state.get("sort_sidebar_open")):
        return

    current_sort_order = str(
        st.session_state.get("l1_sort_order") or L1_SORT_ORDER_DEFAULT
    ).strip()
    if current_sort_order not in L1_SORT_ORDER_LABELS:
        current_sort_order = L1_SORT_ORDER_DEFAULT
        st.session_state["l1_sort_order"] = current_sort_order

    with st.container(key="l1_segment_sidebar"):
        raw_workspace_state = st.session_state.get("l1_segment_workspace_open")
        segment_workspace_open = bool(
            current_sort_order == "segment"
            and (True if raw_workspace_state is None else raw_workspace_state)
        )
        sidebar_heading = "Segment Order" if segment_workspace_open else "Company Sort"

        if not segment_workspace_open:
            # Keep the Company Sort title, metadata, and intro card in one DOM
            # block. The sidebar intentionally uses gap: 0 for Streamlit wrappers,
            # so separate st.markdown() calls can visually collapse together.
            # A local CSS grid gives this top stack deterministic vertical rhythm.
            st.markdown(
                '<div class="tc-sort-home-topstack">'
                f'<div class="tc-segment-sidebar-header">{html.escape(sidebar_heading)}</div>'
                '<div class="tc-segment-sidebar-summary">'
                '<strong>Universal company order</strong>'
                '<span aria-hidden="true">·</span>'
                'client fixed first'
                '</div>'
                '<div class="tc-segment-section-card tc-segment-mode-card">'
                '<div class="tc-segment-section-heading">Choose bar order</div>'
                '<div class="tc-segment-section-copy">Applies to Aura L0, Aura L1, and Geo Cost. Only peers from the second bar onward are reordered.</div>'
                '</div>'
                '</div>',
                unsafe_allow_html=True,
            )
            _render_universal_sort_controls(controls.selected_sections)

            description_by_mode = {
                "revenue_ascending": (
                    "Increasing revenue",
                    "Peers are ordered from lowest to highest revenue after the client.",
                ),
                "revenue_descending": (
                    "Decreasing revenue",
                    "Peers are ordered from highest to lowest revenue after the client.",
                ),
                "sheet_order": (
                    "Sheet order",
                    "Peers retain their original left-to-right order from Excel.",
                ),
                "segment": (
                    "Segment order",
                    "Sort by segment remains active. Select it above to reopen the segment mapping and edit workspace.",
                ),
            }
            heading, copy = description_by_mode.get(
                current_sort_order,
                description_by_mode[L1_SORT_ORDER_DEFAULT],
            )
            st.markdown(
                '<div class="tc-segment-mode-summary">'
                f'<div class="tc-segment-mode-summary-title">{html.escape(heading)}</div>'
                f'<div class="tc-segment-mode-summary-copy">{html.escape(copy)}</div>'
                '</div>',
                unsafe_allow_html=True,
            )
            return

        st.markdown(
            f'<div class="tc-segment-sidebar-header">{html.escape(sidebar_heading)}</div>',
            unsafe_allow_html=True,
        )

        source_signature = str(st.session_state.get("uploaded_signature") or "")
        if st.button(
            "← Sort options",
            key="l1_segment_workspace_back",
            help="Back to Company Sort options",
            type="secondary",
        ):
            st.session_state["l1_segment_workspace_open"] = False
            _set_segment_assignment_editor_open(source_signature, False)
            _rerun_streamlit()
            return

        profile = dict(st.session_state.get("l1_segment_profile") or {})
        direction = str(
            st.session_state.get("l1_segment_sort_direction")
            or L1_SEGMENT_SORT_DIRECTION_DEFAULT
        ).strip()
        if direction not in L1_SEGMENT_SORT_DIRECTION_LABELS:
            direction = L1_SEGMENT_SORT_DIRECTION_DEFAULT
            st.session_state["l1_segment_sort_direction"] = direction

        if not bool(profile.get("detected")):
            message = str(profile.get("error") or "Segments have not been detected yet.")
            st.markdown(
                '<div class="tc-segment-sidebar-error">'
                + html.escape(message)
                + '</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                '<div class="tc-segment-section-divider"></div>'
                '<div class="tc-segment-section-heading">Choose another order</div>',
                unsafe_allow_html=True,
            )
            _render_universal_sort_controls(controls.selected_sections, compact=True)
            return

        segment_count = int(profile.get("segment_count") or 0)
        company_count = int(profile.get("company_count") or 0)
        st.markdown(
            '<div class="tc-segment-detected-banner">'
            f'<span><strong>{segment_count}</strong> segment{"s" if segment_count != 1 else ""}</span>'
            '<span aria-hidden="true">·</span>'
            f'<span><strong>{company_count}</strong> companies</span>'
            '<span aria-hidden="true">·</span>'
            '<span>client fixed first</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        ordered = _ordered_segment_assignments(profile, direction)
        if not ordered:
            st.markdown(
                '<div class="tc-segment-sidebar-error">'
                'Segments were detected, but the company allocation could not be reconstructed. Re-upload the workbook or remap Company names.'
                '</div>',
                unsafe_allow_html=True,
            )
        else:
            editor_open = _segment_assignment_editor_is_open(source_signature)
            if not editor_open:
                _render_segment_mapping_component(ordered)
            _render_segment_assignment_editor_stable(
                profile=profile,
                source_signature=source_signature,
            )

            st.markdown(
                '<div class="tc-segment-section-divider"></div>'
                '<div class="tc-segment-section-card tc-segment-direction-card">'
                '<div class="tc-segment-section-heading">Within each segment</div>'
                '<div class="tc-segment-section-copy">Client remains first; peers are sorted only inside their assigned segment.</div>'
                '</div>',
                unsafe_allow_html=True,
            )
            clicked_direction = ""
            with st.container(key="l1_segment_direction"):
                columns = st.columns(2, gap="small")
                for column, (key, label) in zip(
                    columns,
                    L1_SEGMENT_SORT_DIRECTION_LABELS.items(),
                ):
                    with column:
                        selected = key == direction
                        display = f"✓ {label}" if selected else label
                        if st.button(
                            display,
                            key=f"l1_segment_direction_{key}",
                            type="primary" if selected else "secondary",
                            use_container_width=True,
                        ):
                            clicked_direction = key

            if clicked_direction:
                st.session_state["l1_segment_sort_direction"] = clicked_direction
                _return_to_workbook_review()
                _rerun_streamlit()
                return

        # Keep all universal choices available, but place them after the
        # Segment display/editor so the chosen mode behaves like a workspace.
        st.markdown(
            '<div class="tc-segment-section-divider"></div>'
            '<div class="tc-segment-section-card tc-segment-change-mode-card">'
            '<div class="tc-segment-section-heading">Change company order</div>'
            '<div class="tc-segment-section-copy">Choose another universal order, or keep Sort by segment selected.</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        _render_universal_sort_controls(controls.selected_sections, compact=True)

def render_control_panel(

    paths: AppPaths,
    layout_problems: list[str],
    inspect_workbook: WorkbookInspector,
    load_progress: WorkbookLoadProgress | None = None,
) -> ControlPanelState:
    """Render a clean, compact 25% control rail."""
    downloads_dir = windows_downloads_folder(paths.repository_root)

    workbook_path: Path | None = None
    preview_workbook_path: Path | None = _restore_preview_path()
    preview_metadata = dict(st.session_state.get("preview_metadata") or {})
    preview_error = str(st.session_state.get("preview_error") or "") or None
    workbook_error: str | None = None
    is_new_upload = False

    st.markdown(
        """
        <div class="tc-panel-brand">
          <div class="tc-panel-brand-mark" aria-hidden="true">TC</div>
          <div class="tc-panel-brand-copy">
            <div class="tc-panel-brand-title">Slide Maker</div>
            <div class="tc-panel-brand-subtitle">THINKCELL AUTOMATION</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.container(key="panel_source"):
        with st.container(key="field_source_workbook"):
            uploaded_file = st.file_uploader(
                "Source workbook",
                type=["xlsx", "xlsm"],
                label_visibility="visible",
                key="source_workbook_uploader",
            )

        if uploaded_file is not None:
            try:
                workbook_path, is_new_upload = save_uploaded_workbook(
                    uploaded_file,
                    load_progress=load_progress,
                )
                signature = str(st.session_state.get("uploaded_signature") or "")
                cached_preview_path = _restore_preview_path()
                cached_signature = str(
                    st.session_state.get("preview_source_signature") or ""
                )
                cached_metadata = dict(
                    st.session_state.get("preview_metadata") or {}
                )
                can_reuse_inspection = bool(
                    not is_new_upload
                    and cached_preview_path is not None
                    and cached_preview_path.is_file()
                    and cached_signature == signature
                    and cached_metadata
                )
                if can_reuse_inspection:
                    inspection = UploadInspection(
                        preview_path=cached_preview_path,
                        metadata=cached_metadata,
                        preview_error=(
                            str(st.session_state.get("preview_error") or "") or None
                        ),
                        relevant_tab_count=int(
                            cached_metadata.get("relevant_tab_count") or 0
                        ),
                    )
                else:
                    if load_progress is not None and load_progress.active:
                        load_progress.update(
                            56,
                            "Reading workbook structure",
                            "Finding Aura Summary, Taxonomy, and Geo Cost worksheets.",
                            phase="detect",
                        )
                    inspection = inspect_workbook(workbook_path)
                    inspection.metadata["relevant_tab_count"] = int(
                        inspection.relevant_tab_count
                    )
            except Exception as exc:
                workbook_error = str(exc)
                if load_progress is not None:
                    load_progress.fail(workbook_error)
                st.error(workbook_error)
            else:
                preview_workbook_path = inspection.preview_path
                preview_metadata = inspection.metadata
                preview_error = inspection.preview_error

                st.session_state["preview_workbook_path"] = str(inspection.preview_path)
                st.session_state["preview_source_signature"] = signature
                st.session_state["preview_metadata"] = inspection.metadata
                st.session_state["preview_error"] = inspection.preview_error or ""

                if load_progress is not None and load_progress.active:
                    load_progress.update(
                        62,
                        "Workbook structure ready",
                        "Source worksheets were identified. Detecting slide inputs next.",
                        phase="detect",
                    )

                if inspection.relevant_tab_count:
                    source_label = (
                        f"{inspection.relevant_tab_count} source tab"
                        f"{'s' if inspection.relevant_tab_count != 1 else ''} detected"
                    )
                else:
                    source_label = "Workbook ready for source selection"
                st.markdown(
                    f"""
                    <div class="tc-upload-status" title="{html.escape(uploaded_file.name)}">
                      <span class="tc-upload-status-dot" aria-hidden="true"></span>
                      <span class="tc-upload-status-copy">{html.escape(source_label)}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        else:
            if st.session_state.get("uploaded_workbook_path"):
                _clear_uploaded_workbook_state()
            preview_workbook_path = None
            preview_metadata = {}
            preview_error = None

    with st.container(key="panel_build"):
        with st.container(key="field_workflow_mode"):
            selected_section_values = st.multiselect(
                "Deck to generate",
                options=list(DECK_SECTION_ORDER),
                format_func=lambda value: DECK_SECTION_LABELS[value],
                label_visibility="visible",
                key="workflow_mode",
                on_change=_on_workflow_change,
                placeholder="Select one or more slide types",
            )

        selected_sections = normalize_deck_sections(selected_section_values)
        workflow_mode = workflow_mode_from_sections(selected_sections)
        _synchronize_input_section(selected_sections)
        input_section = "inputs" if selected_sections else ""

        only_l0 = ""
        if selected_sections == ("l0",):
            with st.container(key="field_only_l0"):
                only_l0 = st.text_input(
                    "L0 section filter (optional)",
                    placeholder="G&A, S&M, Ops or OOS",
                    label_visibility="visible",
                    key="only_l0_filter",
                )

    with st.container(key="panel_output"):
        with st.container(key="field_output_name"):
            output_name = st.text_input(
                "PowerPoint filename",
                key="output_name",
                label_visibility="visible",
                placeholder="Generated_ThinkCell_Deck.pptx",
            )
        with st.container(key="field_output_folder"):
            st.markdown(
                '<div class="tc-panel-static-label">Output folder</div>',
                unsafe_allow_html=True,
            )
            st.markdown(
                f"""
                <div class="tc-output-location" title="{html.escape(str(downloads_dir))}">
                  <span class="tc-output-location-path">{html.escape(str(downloads_dir))}</span>
                </div>
                """,
                unsafe_allow_html=True,
            )

    with st.expander("Advanced options", expanded=False):
        keep_temp = st.checkbox(
            "Keep temporary files",
            key="keep_temp",
            help="Retain workflow staging files for troubleshooting.",
        )
        debug = st.checkbox(
            "Debug logging",
            key="debug_logging",
            help="Include verbose diagnostic entries in the build log.",
        )

    if is_new_upload:
        st.session_state["canvas_stage"] = "review"
        st.session_state["workbook_review_finalized"] = False
        st.session_state["preview_sheet_name"] = ""
        st.session_state["preview_source_role"] = "aura"
        st.session_state["preview_input_navigation_token"] = ""
        st.session_state["selected_aura_sheet_override"] = ""
        st.session_state["resolved_aura_sheet"] = ""
        st.session_state["selected_taxonomy_sheet_override"] = ""
        st.session_state["resolved_taxonomy_sheet"] = ""
        st.session_state["resolved_geocost_sheet"] = ""
        st.session_state["taxonomy_sheet_profile"] = {}
        st.session_state["selected_geocost_sheet_override"] = ""
        st.session_state["selected_geocost_scope_override"] = ""
        st.session_state["selected_geocost_global_sheet_override"] = ""
        st.session_state["selected_geocost_us_sheet_override"] = ""
        st.session_state["workbook_range_selections"] = {}
        st.session_state["active_range_selection"] = None
        st.session_state["slide_input_mappings"] = {}
        st.session_state["active_input_by_section"] = {}
        st.session_state["l1_mapping_validation"] = {}
        st.session_state["l1_auto_detection"] = {}
        st.session_state["geocost_auto_detection"] = {}
        st.session_state["geocost_input_profile"] = {}
        st.session_state["geocost_mapping_validation"] = {}
        st.session_state["l1_master_label_profile"] = {}
        st.session_state["l1_function_lookup_profile"] = {}
        st.session_state["auto_mapping_rerun_signature"] = ""
        st.session_state["segment_profile_rerun_signature"] = ""

    workbook_finalized = bool(
        st.session_state.get("workbook_review_finalized", False)
    )
    disabled = (
        workbook_path is None
        or workbook_error is not None
        or bool(layout_problems)
        or not selected_sections
        or not workbook_finalized
    )

    with st.container(key="panel_action"):
        generate_clicked = st.button(
            "Generate PowerPoint",
            type="primary",
            use_container_width=True,
            disabled=disabled,
            key="generate_powerpoint",
        )

        footnote = ""
        footnote_class = ""
        if workbook_path is None and workbook_error is None:
            footnote = "Upload a workbook to begin."
        elif layout_problems:
            footnote = "The backend layout needs attention."
            footnote_class = " tc-panel-footnote-error"
        elif not selected_sections:
            footnote = "Select at least one slide type."
            footnote_class = " tc-panel-footnote-error"
        elif not workbook_finalized:
            footnote = "Finalize the workbook on the right to enable generation."

        if footnote:
            st.markdown(
                f'<div class="tc-panel-footnote{footnote_class}">{html.escape(footnote)}</div>',
                unsafe_allow_html=True,
            )

    if generate_clicked:
        st.session_state["canvas_stage"] = "console"

    # Render the decisive control-color override after all sidebar widgets so
    # it wins against Streamlit/BaseWeb styles injected during widget creation.
    _inject_final_light_control_text_css()
    input_sidebar_active = bool(selected_sections)
    sortable_selected = bool(
        set(selected_sections).intersection(SORTABLE_DECK_SECTIONS)
    )
    if not sortable_selected:
        st.session_state["sort_sidebar_open"] = False
    # Restore the proven v59 rail lifecycle: the third rail is emitted only
    # after the Company sort launcher is used. Once present, its own arrow can
    # collapse it to the 2.5% compact rail and expand it again independently.
    segment_sidebar_active = (
        sortable_selected
        and bool(st.session_state.get("sort_sidebar_open"))
    )
    segment_sidebar_requested_open = segment_sidebar_active
    _inject_rendered_control_text_repair(
        input_sidebar_active=input_sidebar_active,
        segment_sidebar_active=segment_sidebar_active,
        segment_sidebar_requested_open=segment_sidebar_requested_open,
    )

    return ControlPanelState(
        workbook_path=workbook_path,
        preview_workbook_path=preview_workbook_path,
        preview_metadata=preview_metadata,
        preview_error=preview_error,
        workbook_error=workbook_error,
        workbook_display_name=str(
            st.session_state.get("uploaded_workbook_name") or "Source workbook"
        ),
        is_new_upload=is_new_upload,
        workflow_mode=workflow_mode,
        selected_sections=selected_sections,
        input_section=input_section,
        l1_sort_order=str(
            st.session_state.get("l1_sort_order") or L1_SORT_ORDER_DEFAULT
        ),
        l1_segment_sort_direction=str(
            st.session_state.get("l1_segment_sort_direction")
            or L1_SEGMENT_SORT_DIRECTION_DEFAULT
        ),
        only_l0=only_l0,
        output_dir=downloads_dir,
        output_name=output_name,
        keep_temp=bool(keep_temp),
        debug=bool(debug),
        workbook_finalized=workbook_finalized,
        generate_clicked=bool(generate_clicked),
    )


__all__ = [
    "AppPaths",
    "BuildOptions",
    "ControlPanelState",
    "DECK_SECTION_LABELS",
    "DECK_SECTION_ORDER",
    "AURA_INPUT_BY_KEY",
    "AURA_INPUT_FIELDS",
    "AURA_INPUT_KEYS",
    "AURA_INPUT_MAP_SCHEMA_VERSION",
    "AURA_SHARED_INPUT_KEYS",
    "AURA_L0_QUARTILE_INPUT_KEYS",
    "AURA_L1_BENCHMARK_INPUT_KEYS",
    "GEOCOST_INPUT_BY_KEY",
    "GEOCOST_INPUT_FIELDS",
    "GEOCOST_INPUT_KEYS",
    "GEOCOST_INPUT_MAP_SCHEMA_VERSION",
    "STANDARD_INPUT_BY_KEY",
    "STANDARD_INPUT_FIELDS",
    "STANDARD_INPUT_KEYS",
    "INPUT_SECTION_LABELS",
    "INPUT_SECTION_BY_DECK_SECTION",
    "L1_INPUT_BY_KEY",
    "L1_INPUT_FIELDS",
    "L1_INPUT_KEYS",
    "L1_INPUT_MAP_SCHEMA_VERSION",
    "L1_SORT_ORDER_DEFAULT",
    "L1_SORT_ORDER_LABELS",
    "L1_SEGMENT_SORT_DIRECTION_DEFAULT",
    "L1_SEGMENT_SORT_DIRECTION_LABELS",
    "PRESENTATION_SUFFIXES",
    "UploadInspection",
    "WorkbookLoadProgress",
    "active_l1_input_key",
    "aura_input_keys_for_sections",
    "geocost_input_keys_for_sections",
    "standardized_input_keys_for_sections",
    "apply_l1_segment_assignment_overrides",
    "apply_l1_segment_name_overrides",
    "get_aura_input_mappings",
    "get_ignored_aura_taxonomy_functions",
    "add_ignored_aura_taxonomy_function",
    "remove_ignored_aura_taxonomy_function",
    "get_geocost_input_mappings",
    "get_standard_input_mappings",
    "get_l1_input_mappings",
    "get_l1_segment_assignment_overrides",
    "get_l1_segment_name_overrides",
    "initialize_control_panel_state",
    "normalize_deck_sections",
    "workflow_mode_from_sections",
    "normalize_output_name",
    "render_control_panel",
    "render_input_sidebar",
    "render_segment_sidebar",
    "set_aura_input_mappings",
    "set_geocost_input_mappings",
    "set_standard_input_mappings",
    "set_l1_input_mappings",
    "set_l1_segment_assignment_overrides",
    "set_l1_segment_name_overrides",
    "validate_l1_input_mappings",
    "validate_geocost_input_mappings",
    "validate_standard_input_mappings",
    "validate_project_layout",
    "announce_ui_diagnostics",
    "ui_diagnostic_log_path",
    "ui_log",
    "ui_stage",
]
