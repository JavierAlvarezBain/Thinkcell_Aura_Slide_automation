from __future__ import annotations

from contextlib import contextmanager
import argparse
import json
import gc
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import logging
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import zipfile
from typing import Any, Iterable

try:
    import pythoncom
except ImportError:  # Allows workbook detection utilities to run outside Windows/Office.
    class _PythonComShim:
        def PumpWaitingMessages(self) -> None:
            return None

        def CoInitialize(self) -> None:
            return None

        def CoUninitialize(self) -> None:
            return None

    pythoncom = _PythonComShim()

MSO_TRUE = -1
MSO_FALSE = 0

SHARED_BUILD_ID = "shared-geocost-input-map-v2026-09-20-r70"

BASE_DIR = Path(__file__).parent.resolve()


def _find_aura_root(base_dir: Path) -> Path:
    """Support both App/Aura/... and legacy App/.../Aura/... installations."""
    if (base_dir / "Aura L1").is_dir():
        return base_dir
    nested = base_dir / "Aura"
    if (nested / "Aura L1").is_dir():
        return nested
    return base_dir


AURA_ROOT = _find_aura_root(BASE_DIR)
APP_ROOT = AURA_ROOT.parent
AURA_L0_DIR = AURA_ROOT / "Aura L0"
AURA_L1_DIR = AURA_ROOT / "Aura L1"
AURA_SUMMARY_DIR = AURA_ROOT / "Aura Summary"
GEOCOST_DIR = AURA_ROOT / "Global Geo Cost"

DEFAULT_EXCEL_PATH = AURA_L1_DIR / "Aura_Summary_Test.xlsx"
DEFAULT_L0_TEMPLATE_PPT = AURA_L0_DIR / "ThinkCell_L0.pptx"
DEFAULT_TEMPLATE_PPT = AURA_L1_DIR / "ThinkCell_Test.pptx"
DEFAULT_SUMMARY_TEMPLATE_PPT = AURA_SUMMARY_DIR / "Table Test.pptx"
DEFAULT_GEOCOST_GLOBAL_TEMPLATE_PPT = GEOCOST_DIR / "Geo_Cost_Global_Template.pptx"
DEFAULT_GEOCOST_US_TEMPLATE_PPT = GEOCOST_DIR / "Geo_Cost_US_Template.pptx"
DEFAULT_OUTPUT_DIR = AURA_L1_DIR / "output"


@dataclass
class BuildConfig:
    excel_path: Path
    template_ppt: Path
    output_dir: Path
    output_ppt_name: str = "Generated_ThinkCell_Deck.pptx"

    # Aura source workbook layout. Leave as None for automatic detection.
    sheet_name: str | None = None
    header_row_excel: int | None = None
    function_col: int | None = None
    company_start_col: int | None = None
    company_end_col: int | None = None
    quartile_start_col: int | None = None
    quartile_end_col: int | None = None
    num_companies: int | None = None
    autodetect_source_layout: bool = True
    detected_source_layout: dict[str, Any] | None = None

    # Optional explicit range contract produced by the standardized Streamlit
    # Inputs sidebar. Aura L0 and Aura L1 consume the same shared ranges;
    # function-level benchmark columns are required only when L1 is selected.
    aura_input_map_path: Path | None = None

    # Legacy alias retained for older commands/manifests. New callers should
    # populate ``aura_input_map_path``; the source reader accepts either field.
    l1_input_map_path: Path | None = None

    # Universal company/bar order. The first company is always the client and
    # remains fixed; this setting controls peer bars in Aura L0, Aura L1, and
    # Geo Cost. ``l1_*`` aliases are retained for older callers.
    company_sort_order: str | None = None
    company_segment_sort_direction: str | None = None
    company_sort_manifest_path: Path | None = None
    company_sort_assignments: list[dict[str, Any]] | None = None

    # Legacy Aura L1 aliases. New callers should populate the universal fields
    # above; readers fall back to these values for backward compatibility.
    l1_sort_order: str = "revenue_descending"
    l1_segment_sort_direction: str = "revenue_descending"
    l1_segment_assignments: list[dict[str, Any]] | None = None

    # Aura L1 taxonomy source. Leave as None to auto-detect the strongest
    # worksheet whose tab name contains "taxonomy" (case-insensitive).
    taxonomy_sheet_name: str | None = None

    chart_name: str = "fte_chart"
    chart_rows: int = 6
    chart_cols: int = 11
    chart_header_row: int = 1
    revenue_chart_row: int = 2
    fte_chart_row: int = 6
    chart_slide_search_text: str = "FTEs per $B Revenue"

    max_slides: int | None = None
    only_function: str | None = None
    update_quartiles: bool = True
    update_agenda_text: bool = True

    excel_visible: bool = True
    ppt_visible: bool = True
    with_window: bool = True

    quit_office_apps: bool = False
    keep_temp: bool = False
    debug: bool = False
    dump_shapes: bool = False

    kill_office_first: bool = False
    com_check_only: bool = False

    hang_dump_seconds: int = 0
    thinkcell_retries: int = 3
    post_chart_update_wait_seconds: float = 1.2

    # Stable Aura L1 workflow. Automated think-cell carry-over remains disabled.
    carryover_mode: str = "defer"

    # Aura L0 is generated as its own section and patched only after the final
    # consolidated deck has been merged, saved, and closed.
    include_l0: bool = True
    l0_template_ppt: Path | None = None
    l0_only_group: str | None = None
    l0_max_slides: int | None = None
    l0_update_quartiles: bool = True
    l0_post_chart_wait_seconds: float = 1.5

    # Optional Geo Cost sections appended after Aura L1 slides.
    # The standardized UI can pin the Overall table row span plus the repeated
    # >>> function-marker row through a compact JSON manifest.
    geocost_input_map_path: Path | None = None
    include_geocost: bool = True
    geocost_scope: str = "auto"  # auto/all/global/us
    geocost_global_sheet_name: str | None = None
    geocost_us_sheet_name: str | None = None
    geocost_global_template_ppt: Path | None = None
    geocost_us_template_ppt: Path | None = None
    geocost_requested_function: str = "ALL"
    geocost_max_functions: int | None = None
    geocost_max_companies: int = 8
    geocost_label_threshold: float = 3.0
    geocost_global_chart_name: str = ""
    geocost_us_chart_name: str = ""
    geocost_post_chart_wait_seconds: float = 2.0

    @property
    def output_ppt(self) -> Path:
        return self.output_dir / self.output_ppt_name


def setup_logging(config: BuildConfig) -> logging.Logger:
    config.output_dir.mkdir(parents=True, exist_ok=True)
    log_dir = config.output_dir / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"run_{timestamp}.log"

    logger = logging.getLogger("thinkcell_automation")
    logger.setLevel(logging.DEBUG if config.debug else logging.INFO)
    logger.handlers.clear()

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG if config.debug else logging.INFO)
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%H:%M:%S",
    )
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    logger.info("Log file: %s", log_path)
    logger.info("Shared automation build: %s", SHARED_BUILD_ID)
    return logger


@contextmanager
def stage(logger: logging.Logger, name: str):
    start = time.perf_counter()
    logger.info("START: %s", name)
    try:
        yield
    except Exception:
        logger.exception("FAILED: %s", name)
        raise
    else:
        logger.info("DONE: %s (%.2fs)", name, time.perf_counter() - start)


def safe_text(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(value)).strip("_")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_letters(value: Any) -> str:
    return re.sub(r"[^A-Za-z]", "", str(value)).lower()


def normalize_sheet_name_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def looks_like_geocost_sheet_name(value: Any) -> bool:
    key = normalize_sheet_name_key(value)
    return "geo" in key and "cost" in key


def geocost_scope_hint_from_sheet_name(value: Any) -> str | None:
    text = clean_text(value).lower()
    key = normalize_sheet_name_key(value)
    tokens = set(re.findall(r"[a-z0-9]+", text))
    if (
        any(token in tokens for token in {"us", "usa", "domestic"})
        or "unitedstates" in key
        or "usgeocost" in key
        or "usageocost" in key
    ):
        return "us"
    if any(token in tokens for token in {"global", "world", "worldwide", "international", "intl"}):
        return "global"
    if looks_like_geocost_sheet_name(value):
        return "unknown"
    return None


def agenda_display_text(function_name: str) -> str:
    return function_name.upper()


def to_number_or_none(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str):
        cleaned = value.replace("$", "").replace(",", "").replace("\u00a0", "").strip()
        if cleaned in {"", "-", "nan", "None", "<no data>"}:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    try:
        return float(value)
    except Exception:
        return None


def to_int_or_none(value: Any) -> int | None:
    number = to_number_or_none(value)
    if number is None:
        return None
    return int(round(number))


def parse_numeric_text(text: Any) -> float | None:
    cleaned = str(text).replace("|", "").replace("$", "").replace(",", "").replace("\u00a0", "").strip()
    if cleaned in {"", "-", "nan", "None", "<no data>"}:
        return None
    try:
        return float(cleaned)
    except Exception:
        return None



COMPANY_SORT_ORDER_LABELS = {
    "revenue_ascending": "increasing revenue",
    "revenue_descending": "decreasing revenue",
    "sheet_order": "Excel sheet order",
    "segment": "segment order",
}


def normalize_company_sort_order(value: Any) -> str:
    """Return one supported universal company-axis order key."""
    raw = str(value or "revenue_descending").strip().lower()
    aliases = {
        "ascending": "revenue_ascending",
        "increasing": "revenue_ascending",
        "increasing_revenue": "revenue_ascending",
        "revenue_asc": "revenue_ascending",
        "revenue_ascending": "revenue_ascending",
        "descending": "revenue_descending",
        "decreasing": "revenue_descending",
        "decreasing_revenue": "revenue_descending",
        "revenue_desc": "revenue_descending",
        "revenue_descending": "revenue_descending",
        "sheet": "sheet_order",
        "source": "sheet_order",
        "source_order": "sheet_order",
        "excel": "sheet_order",
        "excel_order": "sheet_order",
        "sheet_order": "sheet_order",
        "segment": "segment",
        "segment_order": "segment",
        "by_segment": "segment",
    }
    normalized = aliases.get(raw, raw)
    if normalized not in COMPANY_SORT_ORDER_LABELS:
        raise ValueError(
            f"Unsupported company sort order {value!r}. "
            f"Expected one of {list(COMPANY_SORT_ORDER_LABELS)}."
        )
    return normalized


def normalize_company_segment_sort_direction(value: Any) -> str:
    raw = str(value or "revenue_descending").strip().lower()
    aliases = {
        "ascending": "revenue_ascending",
        "increasing": "revenue_ascending",
        "revenue_ascending": "revenue_ascending",
        "descending": "revenue_descending",
        "decreasing": "revenue_descending",
        "revenue_descending": "revenue_descending",
    }
    normalized = aliases.get(raw, raw)
    if normalized not in {"revenue_ascending", "revenue_descending"}:
        raise ValueError(
            f"Unsupported within-segment order {value!r}. "
            "Expected revenue_ascending or revenue_descending."
        )
    return normalized


def normalized_company_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def load_company_sort_manifest(path: Path | str | None) -> dict[str, Any]:
    """Load the UI-generated universal sort profile, when supplied."""
    if path is None or not str(path).strip():
        return {}
    resolved = Path(path).expanduser().resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"Company sort manifest not found: {resolved}")
    try:
        payload = json.loads(resolved.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Company sort manifest is not readable JSON: {resolved}") from exc
    if not isinstance(payload, dict):
        raise ValueError("Company sort manifest must contain one JSON object.")
    assignments = payload.get("assignments")
    if assignments is not None and not isinstance(assignments, list):
        raise ValueError("Company sort manifest assignments must be a list.")
    return payload


def company_sort_settings_from_config(
    config: Any,
) -> tuple[str, str, list[dict[str, Any]]]:
    """Resolve universal settings while accepting the former L1-only fields."""
    order_value = getattr(config, "company_sort_order", None)
    if not order_value:
        order_value = getattr(config, "l1_sort_order", "revenue_descending")
    direction_value = getattr(config, "company_segment_sort_direction", None)
    if not direction_value:
        direction_value = getattr(
            config,
            "l1_segment_sort_direction",
            "revenue_descending",
        )
    assignments = getattr(config, "company_sort_assignments", None)
    if assignments is None:
        assignments = getattr(config, "l1_segment_assignments", None)
    return (
        normalize_company_sort_order(order_value),
        normalize_company_segment_sort_direction(direction_value),
        [dict(item) for item in (assignments or []) if isinstance(item, dict)],
    )


def _assignment_lookup(
    assignments: Iterable[dict[str, Any]] | None,
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    by_company: dict[str, dict[str, Any]] = {}
    segment_order: list[str] = []
    ordered_items = sorted(
        [dict(item) for item in (assignments or []) if isinstance(item, dict)],
        key=lambda item: int(item.get("company_index") or 0),
    )
    for item in ordered_items:
        company_key = normalized_company_key(item.get("company"))
        if company_key and company_key not in by_company:
            by_company[company_key] = item
        segment = clean_text(item.get("segment"))
        if (
            segment
            and segment.lower() not in {"client", "unassigned"}
            and segment not in segment_order
        ):
            segment_order.append(segment)
    return by_company, segment_order


def build_company_order(
    company_headers: list[Any],
    revenue_values: list[Any] | None = None,
    *,
    sort_order: Any = "revenue_descending",
    reference_assignments: Iterable[dict[str, Any]] | None = None,
    segment_sort_direction: Any = "revenue_descending",
    logger: logging.Logger | None = None,
    context: str = "company axis",
) -> list[int]:
    """Build one client-first order usable by Aura L0, Aura L1, and Geo Cost.

    Revenue values from the current source take precedence. For a Geo Cost
    block, where no revenue row exists, the UI-generated reference assignment
    supplies revenue and segment metadata by normalized company name. Companies
    missing from the reference are retained at the end in their original order.
    """
    mode = normalize_company_sort_order(sort_order)
    count = len(company_headers)
    if count <= 1 or mode == "sheet_order":
        return list(range(count))

    source_revenue = list(revenue_values or [])
    if source_revenue and len(source_revenue) != count:
        raise ValueError(
            f"{context}: company and revenue counts differ: "
            f"companies={count}, revenue={len(source_revenue)}."
        )

    assignment_by_company, reference_segment_order = _assignment_lookup(
        reference_assignments
    )

    def assignment_for(index: int) -> dict[str, Any]:
        return assignment_by_company.get(
            normalized_company_key(company_headers[index]),
            {},
        )

    def revenue_for(index: int) -> float | None:
        if source_revenue:
            value = to_number_or_none(source_revenue[index])
            if value is not None:
                return float(value)
        value = to_number_or_none(assignment_for(index).get("revenue"))
        return None if value is None else float(value)

    peer_indices = list(range(1, count))
    direction = normalize_company_segment_sort_direction(segment_sort_direction)

    if logger is not None and not source_revenue and assignment_by_company:
        unmatched_names = [
            clean_text(company_headers[index])
            for index in peer_indices
            if not assignment_for(index)
        ]
        if unmatched_names:
            logger.warning(
                "%s: %s peer company/companies were not found in the universal "
                "sort reference and will remain at the end in source order: %s",
                context,
                len(unmatched_names),
                unmatched_names,
            )

    def revenue_key(index: int, *, descending: bool) -> tuple[int, float, int]:
        revenue = revenue_for(index)
        if revenue is None:
            return (1, 0.0, index)
        return (0, -revenue if descending else revenue, index)

    if mode in {"revenue_ascending", "revenue_descending"}:
        descending = mode == "revenue_descending"
        order = [0] + sorted(
            peer_indices,
            key=lambda index: revenue_key(index, descending=descending),
        )
    else:
        if not assignment_by_company:
            raise ValueError(
                f"{context}: segment sorting was requested, but no company-to-segment "
                "assignments were supplied."
            )
        buckets: dict[str, list[int]] = {}
        observed_order: list[str] = []
        unmatched: list[int] = []
        for index in peer_indices:
            segment = clean_text(assignment_for(index).get("segment"))
            if not segment or segment.lower() in {"client", "unassigned"}:
                unmatched.append(index)
                continue
            if segment not in buckets:
                buckets[segment] = []
                observed_order.append(segment)
            buckets[segment].append(index)

        ordered_segments = [
            segment for segment in reference_segment_order if segment in buckets
        ]
        ordered_segments.extend(
            segment for segment in observed_order if segment not in ordered_segments
        )
        descending = direction == "revenue_descending"
        order = [0]
        for segment in ordered_segments:
            order.extend(
                sorted(
                    buckets[segment],
                    key=lambda index: revenue_key(index, descending=descending),
                )
            )
        order.extend(unmatched)

    if logger is not None:
        ordered_names = [company_headers[index] for index in order]
        ordered_revenue = [revenue_for(index) for index in order]
        logger.info(
            "%s company order: client fixed first; mode=%s; source=%s; chart=%s; revenue=%s",
            context,
            mode,
            list(company_headers),
            ordered_names,
            ordered_revenue,
        )
    return order


def apply_company_order(
    company_headers: list[Any],
    order: list[int],
    *,
    revenue_values: list[Any] | None = None,
    records: list[dict[str, Any]] | None = None,
    record_value_key: str = "fte_values",
) -> tuple[list[Any], list[Any] | None, list[dict[str, Any]] | None]:
    """Apply one index map to labels, revenue, and parallel record vectors."""
    count = len(company_headers)
    if sorted(order) != list(range(count)):
        raise ValueError(
            f"Company order must be a permutation of 0..{count - 1}: {order}"
        )
    ordered_headers = [company_headers[index] for index in order]
    ordered_revenue = None
    if revenue_values is not None:
        values = list(revenue_values)
        if len(values) != count:
            raise ValueError(
                "Company and revenue counts differ while applying sort order: "
                f"companies={count}, revenue={len(values)}."
            )
        ordered_revenue = [values[index] for index in order]

    ordered_records = None
    if records is not None:
        ordered_records = []
        for record in records:
            values = list(record.get(record_value_key) or [])
            if len(values) != count:
                raise ValueError(
                    f"Record {record.get('function') or record.get('l1')!r} has "
                    f"{len(values)} values; expected {count}."
                )
            copied = dict(record)
            copied[record_value_key] = [values[index] for index in order]
            ordered_records.append(copied)
    return ordered_headers, ordered_revenue, ordered_records

def pump_office_ui(seconds: float = 0.5) -> None:
    end_time = time.time() + max(0.0, float(seconds))
    while time.time() < end_time:
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        time.sleep(0.05)


def get_output_path(config: BuildConfig) -> Path:
    output_ppt = config.output_ppt
    if not output_ppt.exists():
        return output_ppt
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return config.output_dir / f"{output_ppt.stem}_{timestamp}{output_ppt.suffix}"


def require_file(path: Path, label: str) -> None:
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")


def kill_office_processes(logger: logging.Logger) -> None:
    logger.warning("Killing POWERPNT.EXE and EXCEL.EXE because --kill-office was used.")
    for process_name in ["POWERPNT.EXE", "EXCEL.EXE"]:
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", process_name, "/T"],
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception as exc:
            logger.warning("Could not kill %s: %s", process_name, exc)


_MERGE_WORKER_FLAG = "--powerpoint-merge-worker"
_SHUTDOWN_WORKER_FLAG = "--powerpoint-idle-shutdown-worker"
_MERGE_WORKER_TIMEOUT_SECONDS = 15 * 60
_SHUTDOWN_WORKER_TIMEOUT_SECONDS = 45
_MERGE_WORKER_ATTEMPTS = 2
_SLIDE_XML_PATTERN = re.compile(r"^ppt/slides/slide[0-9]+[.]xml$")


def _pptx_slide_count(path: Path) -> int:
    """Count physical slides without opening PowerPoint."""
    try:
        with zipfile.ZipFile(path, "r") as archive:
            return sum(
                1
                for name in archive.namelist()
                if _SLIDE_XML_PATTERN.fullmatch(name)
            )
    except (OSError, zipfile.BadZipFile) as exc:
        raise RuntimeError(f"Presentation is not a readable PPTX/PPTM package: {path}") from exc


def presentation_slide_count(
    path: Path,
    *,
    timeout_seconds: float = 90.0,
) -> int:
    """Count section slides after PowerPoint/OneDrive releases the package."""
    return _wait_for_pptx_readable(
        Path(path).resolve(),
        timeout_seconds=timeout_seconds,
    )


def _wait_for_file_stable(
    path: Path,
    *,
    timeout_seconds: float = 45.0,
    stable_checks: int = 3,
) -> None:
    deadline = time.time() + max(1.0, float(timeout_seconds))
    last_size = -1
    unchanged = 0
    while time.time() < deadline:
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            size = -1
        if size > 0 and size == last_size:
            unchanged += 1
            if unchanged >= max(1, int(stable_checks)):
                return
        else:
            last_size = size
            unchanged = 0
        time.sleep(0.35)
    raise RuntimeError(f"PowerPoint file did not stabilize: {path}")


def _wait_for_pptx_readable(
    path: Path,
    *,
    timeout_seconds: float = 90.0,
) -> int:
    """Wait until a saved presentation can be opened as a ZIP package.

    A successful stat/size check is not enough on Windows: PowerPoint, think-cell,
    antivirus, or OneDrive may still hold a deny-read handle for a few seconds.
    Return the physical slide count once the package is genuinely readable.
    """
    deadline = time.time() + max(1.0, float(timeout_seconds))
    last_error: BaseException | None = None
    while time.time() < deadline:
        try:
            return _pptx_slide_count(path)
        except RuntimeError as exc:
            last_error = exc
            time.sleep(0.5)
    raise RuntimeError(
        "Presentation remained locked or unreadable after "
        f"{timeout_seconds:.0f} seconds: {path}"
    ) from last_error


def _copy_pptx_with_retry(
    source: Path,
    destination: Path,
    *,
    timeout_seconds: float = 90.0,
) -> None:
    """Copy a presentation after Office/OneDrive releases its source handle."""
    deadline = time.time() + max(1.0, float(timeout_seconds))
    last_error: BaseException | None = None
    while time.time() < deadline:
        try:
            destination.unlink(missing_ok=True)
            shutil.copy2(source, destination)
            _wait_for_file_stable(destination, timeout_seconds=15.0)
            _wait_for_pptx_readable(destination, timeout_seconds=15.0)
            return
        except (OSError, RuntimeError) as exc:
            last_error = exc
            try:
                destination.unlink(missing_ok=True)
            except OSError:
                pass
            time.sleep(0.5)
    raise RuntimeError(
        "Could not stage presentation locally after "
        f"{timeout_seconds:.0f} seconds: {source}"
    ) from last_error


def _powerpoint_process_id(ppt_app: Any) -> int | None:
    """Return the process ID for the merge-only PowerPoint instance."""
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        hwnd = int(ppt_app.HWND)
        process_id = wintypes.DWORD(0)
        ctypes.windll.user32.GetWindowThreadProcessId(
            wintypes.HWND(hwnd),
            ctypes.byref(process_id),
        )
        return int(process_id.value) or None
    except Exception:
        return None


def _running_powerpoint_process_ids() -> set[int]:
    """Snapshot PowerPoint PIDs so a recovery kill never targets a pre-existing user session."""
    if os.name != "nt":
        return set()
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
        )
    except Exception:
        return set()

    pids: set[int] = set()
    for line in result.stdout.splitlines():
        match = re.match(r'^"POWERPNT[.]EXE","([0-9]+)"', line.strip(), flags=re.IGNORECASE)
        if match:
            pids.add(int(match.group(1)))
    return pids


def running_powerpoint_process_ids() -> set[int]:
    """Public snapshot used before a workflow starts PowerPoint automation."""
    return _running_powerpoint_process_ids()


def _run_idle_powerpoint_shutdown_worker() -> int:
    """Quit PowerPoint only when it has no open presentations."""
    if os.name != "nt":
        return 0
    try:
        import pythoncom as worker_pythoncom
        import win32com.client as win32
    except ImportError:
        traceback.print_exc()
        return 2

    ppt = None
    worker_pythoncom.CoInitialize()
    try:
        try:
            ppt = win32.GetActiveObject("PowerPoint.Application")
        except Exception:
            print("No active PowerPoint instance remains.", flush=True)
            return 0

        count = int(ppt.Presentations.Count)
        if count > 0:
            print(
                f"PowerPoint has {count} open presentation(s); refusing to close a pre-existing user session.",
                flush=True,
            )
            return 3
        ppt.Quit()
        ppt = None
        print("Closed idle pre-existing PowerPoint instance.", flush=True)
        return 0
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        ppt = None
        gc.collect()
        try:
            worker_pythoncom.CoUninitialize()
        except Exception:
            pass


def _prepare_powerpoint_for_fresh_merge(
    preexisting_pids: set[int],
    logger: logging.Logger,
) -> None:
    """Ensure no PowerPoint process remains before the final merge worker starts.

    PowerPoint is a single-process Automation server. Any process created after the
    workflow began is safe to terminate after every section deck has been saved.
    A process that existed before the workflow is only closed when it is idle.
    """
    current_pids = _running_powerpoint_process_ids()
    workflow_owned_pids = current_pids - preexisting_pids
    for process_id in sorted(workflow_owned_pids):
        _terminate_merge_powerpoint(process_id, set(), logger)

    deadline = time.time() + 20.0
    while time.time() < deadline and (_running_powerpoint_process_ids() - preexisting_pids):
        time.sleep(0.5)

    remaining = _running_powerpoint_process_ids()
    preexisting_remaining = remaining & preexisting_pids
    if preexisting_remaining:
        logger.info(
            "PowerPoint was already running before the workflow; checking whether it is idle before merge."
        )
        command = [
            sys.executable,
            str(Path(__file__).resolve()),
            _SHUTDOWN_WORKER_FLAG,
        ]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                timeout=_SHUTDOWN_WORKER_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                "PowerPoint was already running before this workflow and did not respond while "
                "preparing the final merge. Close PowerPoint and rerun the merge."
            ) from exc

        _log_merge_worker_output(logger, result.stdout)
        _log_merge_worker_output(logger, result.stderr, error=True)
        if result.returncode == 3:
            raise RuntimeError(
                "PowerPoint was already open before this workflow and still contains an open "
                "presentation. Close/save those PowerPoint files before running the integrated deck."
            )
        if result.returncode != 0:
            raise RuntimeError(
                f"Could not close the idle pre-existing PowerPoint process (worker code {result.returncode})."
            )

    deadline = time.time() + 20.0
    while time.time() < deadline:
        if not _running_powerpoint_process_ids():
            logger.info("PowerPoint is fully closed; the final merge can start cleanly.")
            return
        time.sleep(0.5)
    raise RuntimeError(
        "PowerPoint did not fully exit before the final merge. Close PowerPoint and rerun."
    )


def _terminate_merge_powerpoint(
    process_id: int | None,
    preexisting_pids: set[int],
    logger: logging.Logger,
) -> None:
    if not process_id or os.name != "nt":
        return
    if process_id in preexisting_pids:
        logger.warning(
            "Merge worker reported pre-existing PowerPoint PID %s; it will not be terminated.",
            process_id,
        )
        return
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/PID", str(process_id), "/T"],
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode == 0:
            logger.info("Terminated PowerPoint automation process %s before continuing.", process_id)
        else:
            logger.debug(
                "Could not terminate PowerPoint automation process %s: %s",
                process_id,
                (result.stderr or result.stdout).strip(),
            )
    except Exception as exc:
        logger.debug("Could not terminate PowerPoint automation process %s: %r", process_id, exc)


def _close_merge_worker_objects(pres: Any, ppt: Any) -> None:
    try:
        if pres is not None:
            pres.Close()
    except Exception:
        pass
    try:
        if ppt is not None:
            ppt.Quit()
    except Exception:
        pass
    pres = None
    ppt = None
    gc.collect()


def _run_powerpoint_merge_worker(
    deck_paths: list[Path],
    output_ppt: Path,
    *,
    visible: bool,
    with_window: bool,
    pid_file: Path,
    forbidden_pids: set[int],
) -> int:
    """Run the final merge inside a short-lived child process and fresh COM apartment."""
    if os.name != "nt":
        print("PowerPoint merge worker requires Windows.", file=sys.stderr)
        return 2

    try:
        import pythoncom as worker_pythoncom
        import win32com.client as win32
    except ImportError:
        traceback.print_exc()
        return 2

    ppt = None
    pres = None
    worker_pythoncom.CoInitialize()
    try:
        for path in deck_paths:
            if not path.exists():
                raise FileNotFoundError(f"Merge input not found: {path}")

        output_ppt.parent.mkdir(parents=True, exist_ok=True)
        if output_ppt.exists():
            output_ppt.unlink()
        shutil.copy2(deck_paths[0], output_ppt)

        # Never reuse GetActiveObject here. The previous failure happened because
        # the final merge attached to the same PowerPoint server that had just
        # completed the heavy think-cell/label workload.
        ppt = win32.DispatchEx("PowerPoint.Application")
        try:
            ppt.Visible = True
        except Exception:
            pass
        try:
            ppt.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass

        process_id = _powerpoint_process_id(ppt)
        if process_id:
            pid_file.write_text(str(process_id), encoding="ascii")
            if process_id in forbidden_pids:
                raise RuntimeError(
                    "DispatchEx returned a pre-existing PowerPoint process instead of an isolated one."
                )

        pres = ppt.Presentations.Open(
            str(output_ppt),
            MSO_FALSE,
            MSO_FALSE,
            MSO_TRUE if with_window else MSO_FALSE,
        )
        expected_slides = _pptx_slide_count(deck_paths[0])
        actual_slides = int(pres.Slides.Count)
        if actual_slides != expected_slides:
            raise RuntimeError(
                f"Base deck opened with {actual_slides} slide(s); expected {expected_slides}."
            )

        for path in deck_paths[1:]:
            slide_count = _pptx_slide_count(path)
            if slide_count <= 0:
                continue
            before = int(pres.Slides.Count)
            pres.Slides.InsertFromFile(
                str(path),
                before,
                1,
                slide_count,
            )
            pump_office_ui(0.5)
            after = int(pres.Slides.Count)
            inserted = after - before
            if inserted != slide_count:
                raise RuntimeError(
                    f"Inserted {inserted} slide(s) from {path.name}; expected {slide_count}."
                )
            expected_slides += slide_count
            print(f"Appended {slide_count} slide(s) from {path.name}", flush=True)

        final_count = int(pres.Slides.Count)
        if final_count != expected_slides:
            raise RuntimeError(
                f"Combined deck has {final_count} slide(s); expected {expected_slides}."
            )

        pres.Save()
        pump_office_ui(1.0)
        pres.Close()
        pres = None
        ppt.Quit()
        ppt = None

        _wait_for_file_stable(output_ppt)
        package_count = _pptx_slide_count(output_ppt)
        if package_count != expected_slides:
            raise RuntimeError(
                f"Saved combined package has {package_count} slide(s); expected {expected_slides}."
            )
        print(
            f"Created merge worker output with {expected_slides} slide(s): {output_ppt}",
            flush=True,
        )
        return 0
    except Exception:
        traceback.print_exc()
        return 1
    finally:
        _close_merge_worker_objects(pres, ppt)
        try:
            worker_pythoncom.CoUninitialize()
        except Exception:
            pass


def _merge_worker_cli(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--pid-file", type=Path, required=True)
    parser.add_argument("--visible", action="store_true")
    parser.add_argument("--with-window", action="store_true")
    parser.add_argument("--deck", action="append", type=Path, required=True)
    parser.add_argument("--forbidden-pid", action="append", type=int, default=[])
    args = parser.parse_args(argv)
    return _run_powerpoint_merge_worker(
        [path.resolve() for path in args.deck],
        args.output.resolve(),
        visible=args.visible,
        with_window=args.with_window,
        pid_file=args.pid_file.resolve(),
        forbidden_pids=set(args.forbidden_pid),
    )


def _log_merge_worker_output(logger: logging.Logger, text: str, *, error: bool = False) -> None:
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if error:
            logger.debug("Merge worker: %s", line)
        else:
            logger.info("Merge worker: %s", line)


def _preserve_merge_recovery_files(
    paths: list[Path],
    output_ppt: Path,
    logger: logging.Logger,
    reason: str,
) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    recovery_dir = output_ppt.parent / f"{output_ppt.stem}_MERGE_RECOVERY_{timestamp}"
    recovery_dir.mkdir(parents=True, exist_ok=True)
    for index, path in enumerate(paths, start=1):
        destination = recovery_dir / f"{index:02d}_{path.name}"
        shutil.copy2(path, destination)

    # Aura L0 is finalized only after the consolidated deck has been saved and
    # closed. Preserve its compact patch manifest beside the completed section
    # decks so the recovery utility can reconstruct the exact final slide range
    # without rerunning or reinterpreting the L0 source logic.
    l0_metadata_name = "01_Aura_L0_Patch_Metadata.json"
    copied_l0_metadata = False
    for parent in dict.fromkeys(path.parent.resolve() for path in paths):
        candidate = parent / l0_metadata_name
        if not candidate.is_file():
            continue
        shutil.copy2(candidate, recovery_dir / l0_metadata_name)
        copied_l0_metadata = True
        break

    recovery_notes = [
        "The individual sections were generated successfully, but the final PowerPoint merge failed.",
        "They are preserved here in the required order.",
    ]
    if copied_l0_metadata:
        recovery_notes.append(
            "Aura L0 native-patch metadata is included; recover_completed_run.py will apply it after the merge."
        )
    recovery_notes.extend(["", f"Failure: {reason}"])
    (recovery_dir / "README.txt").write_text(
        "\n".join(recovery_notes) + "\n",
        encoding="utf-8",
    )
    logger.error("Preserved completed section decks for recovery: %s", recovery_dir)
    return recovery_dir


def _merge_path_key(value: Any) -> str:
    """Return a Windows-safe key for comparing presentation paths."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        raw = str(Path(raw).resolve())
    except Exception:
        raw = os.path.abspath(raw)
    return os.path.normcase(raw)


def _merge_com_call(
    label: str,
    operation,
    logger: logging.Logger,
    *,
    attempts: int = 10,
):
    """Retry short PowerPoint busy/RPC windows without starting another process."""
    last_error: BaseException | None = None
    for attempt in range(1, max(1, int(attempts)) + 1):
        try:
            return operation()
        except Exception as exc:
            last_error = exc
            if attempt >= attempts:
                break
            logger.debug(
                "PowerPoint was busy during %s (attempt %s/%s): %r",
                label,
                attempt,
                attempts,
                exc,
            )
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(min(0.20 * attempt, 1.25))
    raise RuntimeError(f"PowerPoint failed during {label} after {attempts} attempts.") from last_error


def _find_open_presentation_by_path(ppt: Any, source: Path, logger: logging.Logger):
    wanted = _merge_path_key(source)
    try:
        count = int(
            _merge_com_call(
                "read open presentation count",
                lambda: ppt.Presentations.Count,
                logger,
            )
        )
    except Exception:
        return None

    for index in range(count, 0, -1):
        try:
            presentation = ppt.Presentations.Item(index)
            full_name = str(presentation.FullName or "").strip()
        except Exception:
            continue
        if full_name and _merge_path_key(full_name) == wanted:
            return presentation
    return None


def _stage_section_for_same_session_merge(
    ppt: Any,
    source: Path,
    destination: Path,
    logger: logging.Logger,
) -> int:
    """Stage one section locally, even when that section is still open in PowerPoint."""
    destination.unlink(missing_ok=True)
    open_pres = _find_open_presentation_by_path(ppt, source, logger)

    if open_pres is not None:
        slide_count = int(
            _merge_com_call(
                f"read slide count for open section {source.name}",
                lambda: open_pres.Slides.Count,
                logger,
            )
        )
        try:
            is_saved = bool(open_pres.Saved)
        except Exception:
            is_saved = True
        if not is_saved:
            _merge_com_call(
                f"save workflow section {source.name}",
                lambda: open_pres.Save(),
                logger,
            )
        _merge_com_call(
            f"save a local copy of open section {source.name}",
            lambda: open_pres.SaveCopyAs(str(destination)),
            logger,
        )
        local_count = _wait_for_pptx_readable(destination, timeout_seconds=90.0)
        if local_count != slide_count:
            raise RuntimeError(
                f"The local copy of {source.name} has {local_count} slide(s); "
                f"the open section has {slide_count}."
            )

        # The path is an exact workflow output path, so closing this presentation
        # is safe. Failure is non-fatal because the local copy is already complete.
        try:
            open_pres.Saved = True
        except Exception:
            pass
        try:
            _merge_com_call(
                f"close workflow section {source.name}",
                lambda: open_pres.Close(),
                logger,
                attempts=6,
            )
            logger.info("Closed saved workflow section before collation: %s", source.name)
        except Exception as exc:
            logger.warning(
                "PowerPoint kept %s open, but its complete local copy is available; "
                "collation will continue in the same session: %r",
                source.name,
                exc,
            )
        return slide_count

    source_count = _wait_for_pptx_readable(source, timeout_seconds=90.0)
    if source_count <= 0:
        raise ValueError(f"Presentation section contains no slides: {source}")
    _copy_pptx_with_retry(source, destination, timeout_seconds=90.0)
    local_count = _wait_for_pptx_readable(destination, timeout_seconds=30.0)
    if local_count != source_count:
        raise RuntimeError(
            f"Local staging changed the slide count for {source.name}: "
            f"source={source_count}, staged={local_count}."
        )
    return source_count


def _publish_same_session_merge(
    merged_local: Path,
    requested_output: Path,
    expected_slides: int,
    logger: logging.Logger,
) -> Path:
    """Publish atomically; use a timestamped name instead of crashing on a locked old output."""
    requested_output.parent.mkdir(parents=True, exist_ok=True)
    publish_temp = requested_output.parent / (
        f".{requested_output.stem}.publishing_{int(time.time())}{requested_output.suffix}"
    )
    publish_temp.unlink(missing_ok=True)
    shutil.copy2(merged_local, publish_temp)
    _wait_for_pptx_readable(publish_temp, timeout_seconds=60.0)

    published = requested_output
    try:
        os.replace(publish_temp, requested_output)
    except OSError as exc:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        published = requested_output.with_name(
            f"{requested_output.stem}_{timestamp}{requested_output.suffix}"
        )
        logger.warning(
            "The requested output path is locked. Publishing the completed deck "
            "under a timestamped name instead of failing: %s (%r)",
            published,
            exc,
        )
        os.replace(publish_temp, published)

    actual = _wait_for_pptx_readable(published, timeout_seconds=90.0)
    if actual != expected_slides:
        raise RuntimeError(
            f"Published combined deck contains {actual} slide(s); expected {expected_slides}."
        )
    return published


def combine_presentation_files(
    deck_paths: Iterable[Path],
    output_ppt: Path,
    logger: logging.Logger,
    *,
    visible: bool = True,
    with_window: bool = True,
    quit_office: bool = False,
    stage_name: str = "Combine presentation sections",
    preexisting_powerpoint_pids: Iterable[int] | None = None,
) -> Path:
    """Combine saved sections in the PowerPoint session already used by the workflow.

    The former implementation first demanded that PowerPoint exit and then
    launched a second merge worker. That conflicts with section builders that
    intentionally reuse the same PowerPoint application. This implementation
    never shuts down PowerPoint before collation, never starts a merge subprocess,
    and never closes an unrelated user presentation. It creates one temporary
    master presentation, appends Summary/L0/L1/Geo Cost in order, saves a copy,
    closes only that master, and returns the closed PPTX for the final L0 package
    patch.
    """
    paths = [Path(path).resolve() for path in deck_paths]
    if not paths:
        raise ValueError("At least one presentation is required for combination.")
    for index, path in enumerate(paths, start=1):
        require_file(path, f"Presentation section {index}")

    requested_output = Path(output_ppt).resolve()
    requested_output.parent.mkdir(parents=True, exist_ok=True)
    preexisting_pids = {int(pid) for pid in (preexisting_powerpoint_pids or [])}
    merge_dir = Path(tempfile.mkdtemp(prefix="aura_same_session_merge_"))

    ppt = None
    master = None
    powerpoint_pid: int | None = None
    pythoncom.CoInitialize()
    try:
        with stage(logger, stage_name):
            if os.name != "nt":
                raise RuntimeError("PowerPoint collation requires Windows and pywin32.")
            try:
                import win32com.client as win32
            except ImportError as exc:
                raise RuntimeError(
                    "pywin32 is unavailable. Run app.py with the .thinkcell Python environment."
                ) from exc

            try:
                ppt = win32.GetActiveObject("PowerPoint.Application")
                logger.info(
                    "Using the existing workflow PowerPoint session for final collation."
                )
            except Exception:
                ppt = win32.Dispatch("PowerPoint.Application")
                logger.info("Started PowerPoint for same-session final collation.")

            powerpoint_pid = _powerpoint_process_id(ppt)
            try:
                ppt.Visible = bool(visible)
            except Exception:
                pass

            local_paths: list[Path] = []
            slide_counts: list[int] = []
            for index, source in enumerate(paths, start=1):
                local = merge_dir / f"section_{index:02d}{source.suffix.lower()}"
                count = _stage_section_for_same_session_merge(
                    ppt,
                    source,
                    local,
                    logger,
                )
                local_paths.append(local)
                slide_counts.append(count)
                logger.info(
                    "Prepared section %s for collation: %s slide(s) from %s",
                    index,
                    count,
                    source.name,
                )

            expected_total = sum(slide_counts)
            if expected_total <= 0:
                raise RuntimeError("The completed sections contain no slides.")

            master_working = merge_dir / f"combined_working{local_paths[0].suffix}"
            shutil.copy2(local_paths[0], master_working)
            _wait_for_pptx_readable(master_working, timeout_seconds=30.0)

            master = _merge_com_call(
                "open the temporary combined presentation",
                lambda: ppt.Presentations.Open(
                    str(master_working),
                    MSO_FALSE,
                    MSO_FALSE,
                    MSO_TRUE if with_window else MSO_FALSE,
                ),
                logger,
            )
            actual = int(
                _merge_com_call(
                    "read the base combined slide count",
                    lambda: master.Slides.Count,
                    logger,
                )
            )
            if actual != slide_counts[0]:
                raise RuntimeError(
                    f"Base combined presentation opened with {actual} slide(s); "
                    f"expected {slide_counts[0]}."
                )

            cumulative = slide_counts[0]
            for local, section_count in zip(local_paths[1:], slide_counts[1:]):
                insertion_index = int(
                    _merge_com_call(
                        "read the current combined slide count",
                        lambda: master.Slides.Count,
                        logger,
                    )
                )
                _merge_com_call(
                    f"append {local.name}",
                    lambda local=local, insertion_index=insertion_index, section_count=section_count: (
                        master.Slides.InsertFromFile(
                            str(local),
                            insertion_index,
                            1,
                            section_count,
                        )
                    ),
                    logger,
                )
                cumulative += section_count
                actual = int(
                    _merge_com_call(
                        f"verify slide count after {local.name}",
                        lambda: master.Slides.Count,
                        logger,
                    )
                )
                if actual != cumulative:
                    raise RuntimeError(
                        f"Combined deck has {actual} slide(s) after {local.name}; "
                        f"expected {cumulative}."
                    )
                logger.info(
                    "Appended %s slide(s) from %s; combined count=%s.",
                    section_count,
                    local.name,
                    actual,
                )

            if cumulative != expected_total:
                raise RuntimeError(
                    f"Combined slide count mismatch: cumulative={cumulative}, "
                    f"expected={expected_total}."
                )

            # SaveCopyAs writes a closed, independent package. The temporary
            # master can then be discarded without ever locking the requested
            # output or the file that receives the final L0 native patch.
            merged_local = merge_dir / f"combined_complete{requested_output.suffix}"
            _merge_com_call(
                "save a closed copy of the combined presentation",
                lambda: master.SaveCopyAs(str(merged_local)),
                logger,
            )
            local_total = _wait_for_pptx_readable(merged_local, timeout_seconds=120.0)
            if local_total != expected_total:
                raise RuntimeError(
                    f"PowerPoint saved {local_total} slide(s); expected {expected_total}."
                )

            try:
                master.Saved = True
            except Exception:
                pass
            try:
                _merge_com_call(
                    "close the temporary combined presentation",
                    lambda: master.Close(),
                    logger,
                    attempts=8,
                )
                master = None
            except Exception as exc:
                logger.warning(
                    "PowerPoint kept only the disposable temporary master open. "
                    "The independent combined copy is already complete, so the "
                    "workflow will publish it instead of failing: %r",
                    exc,
                )

            published = _publish_same_session_merge(
                merged_local,
                requested_output,
                expected_total,
                logger,
            )
            logger.info(
                "Created combined deck in the existing PowerPoint session: %s (%s slides)",
                published,
                expected_total,
            )
            return published
    except Exception as exc:
        recovery_dir = _preserve_merge_recovery_files(
            paths,
            requested_output,
            logger,
            repr(exc),
        )
        raise RuntimeError(
            "Final collation failed, but every completed section deck was preserved at "
            f"{recovery_dir}"
        ) from exc
    finally:
        if master is not None:
            try:
                master.Saved = True
            except Exception:
                pass
            try:
                master.Close()
            except Exception:
                pass
            master = None

        if ppt is not None:
            try:
                open_count = int(ppt.Presentations.Count)
            except Exception:
                open_count = -1
            workflow_owned_session = bool(
                powerpoint_pid is not None and powerpoint_pid not in preexisting_pids
            ) or not preexisting_pids
            if open_count == 0 and (quit_office or workflow_owned_session):
                try:
                    ppt.Quit()
                    logger.info("Closed the idle workflow PowerPoint session after collation.")
                except Exception:
                    pass
            elif open_count > 0:
                logger.info(
                    "Left PowerPoint running with %s unrelated/open presentation(s); "
                    "none were closed by the collator.",
                    open_count,
                )
            ppt = None

        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        shutil.rmtree(merge_dir, ignore_errors=True)


if __name__ == "__main__":
    if _MERGE_WORKER_FLAG in sys.argv:
        flag_index = sys.argv.index(_MERGE_WORKER_FLAG)
        raise SystemExit(_merge_worker_cli(sys.argv[flag_index + 1 :]))
    if _SHUTDOWN_WORKER_FLAG in sys.argv:
        raise SystemExit(_run_idle_powerpoint_shutdown_worker())
