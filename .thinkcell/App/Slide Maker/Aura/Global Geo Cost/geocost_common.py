from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import logging
import re
import sys
import time
import unicodedata
from typing import Any, Callable, Literal

try:
    import pythoncom
    import win32com.client as win32
except ImportError:  # Source detection and ChartData audits work without Windows/Office.
    pythoncom = None
    win32 = None


GeoCostScope = Literal["global", "us"]

# Office constants. Numeric values avoid a dependency on generated COM enums.
MSO_TRUE = -1
MSO_FALSE = 0
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_ANCHOR_MIDDLE = 3
PP_ALIGN_CENTER = 2
XL_DATA_LABELS_SHOW_VALUE = 2
XL_LABEL_POSITION_CENTER = -4108
MSO_LINE = 9

AUTOLABEL_PREFIX = "GeoCost_AutoLabel_"
BUILD_ID = "geocost-integrated-auto-sheet-global-us-modular-v2026-07-27"
DEFAULT_LABEL_THRESHOLD = 3.0
DEFAULT_THINKCELL_ATTEMPTS = 5
DEFAULT_OFFICE_RETRY_ATTEMPTS = 5
PLAIN_VALUE_NUMBER_FORMAT = "0"
PLAIN_PERCENT_NUMBER_FORMAT = "0%"
LOW_COST_ROW_LABEL = "Low cost %"

RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
OBJECT_DISCONNECTED = -2147417848

DISPLAY_LABEL_ALIASES: dict[str, str] = {
    "unitedstates": "USA",
    "unitedkingdom": "UK",
    "unitedarabemirates": "UAE",
    "saudiarabia": "Saudi",
    "russianfederation": "Russia",
    "republicofkorea": "South Korea",
    "districtofcolumbia": "DC",
}

US_STATE_KEYS = {
    re.sub(r"[^a-z0-9]+", "", name.lower())
    for name in (
        "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
        "Connecticut", "Delaware", "District of Columbia", "Florida", "Georgia",
        "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky",
        "Louisiana", "Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota",
        "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire",
        "New Jersey", "New Mexico", "New York", "North Carolina", "North Dakota",
        "Ohio", "Oklahoma", "Oregon", "Pennsylvania", "Rhode Island",
        "South Carolina", "South Dakota", "Tennessee", "Texas", "Utah", "Vermont",
        "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming",
        "Puerto Rico", "Guam", "US Virgin Islands",
    )
}


@dataclass(frozen=True)
class GeoCostSourceLayout:
    marker_row: int
    header_row: int
    data_start_row: int
    segment_bucket_col: int
    high_low_col: int
    raw_country_col: int
    display_country_col: int
    entity_label: str = "Country"
    # Optional bounds supplied by the standardized Geo Cost input map. Legacy
    # auto-detection leaves them unset and retains the previous scan behavior.
    data_end_row: int | None = None
    # Optional explicit row used for the chart's bottom "Low cost %" series.
    # When absent, the legacy backend calculates the ratio from country rows.
    low_cost_row: int | None = None
    low_cost_label_col: int | None = None
    company_start_col: int | None = None
    company_end_col: int | None = None
    marker_start_col: int | None = None
    marker_end_col: int | None = None
    table_start_col: int | None = None
    table_end_col: int | None = None


@dataclass(frozen=True)
class GeoCostTemplateLayout:
    chart_name: str
    series_row_count: int
    country_first_row: int
    country_last_row: int
    data_row: int
    country_names: list[str]
    header_row: int = 1
    low_cost_row_label: str = LOW_COST_ROW_LABEL

    @property
    def country_count(self) -> int:
        return len(self.country_names)


@dataclass
class GeoCostSheetMatch:
    scope: GeoCostScope
    sheet_name: str
    source_layout: GeoCostSourceLayout
    name_score: int
    structure_score: int
    inferred_from_content: bool = False

    @property
    def total_score(self) -> int:
        return self.name_score + self.structure_score


@dataclass
class GeoCostWorkbookDetection:
    global_sheet: GeoCostSheetMatch | None = None
    us_sheet: GeoCostSheetMatch | None = None
    candidates: list[GeoCostSheetMatch] | None = None

    def selected(self, scopes: set[str] | None = None) -> list[GeoCostSheetMatch]:
        scopes = scopes or {"global", "us"}
        result: list[GeoCostSheetMatch] = []
        if "global" in scopes and self.global_sheet is not None:
            result.append(self.global_sheet)
        if "us" in scopes and self.us_sheet is not None:
            result.append(self.us_sheet)
        return result


@dataclass
class FunctionBlock:
    name: str
    start_col: int
    next_start_col: int | None = None


@dataclass
class CountryRow:
    row_num: int
    country: str
    segment_bucket: str
    high_low: str
    values: list[float]

    @property
    def is_low_cost(self) -> bool:
        return self.high_low.lower() == "low" or self.segment_bucket.lower().endswith("-low")


@dataclass
class LabelTarget:
    company_index: int
    row_index: int
    country: str
    template_country: str
    display_label: str
    segment_bucket: str
    value: float
    total: float
    share_pct: float


@dataclass
class GeoCostData:
    scope: GeoCostScope
    sheet_name: str
    function_name: str
    companies: list[str]
    country_rows: list[CountryRow]
    template_layout: GeoCostTemplateLayout
    totals: list[float]
    low_cost_ratio: list[float | None]
    label_targets: list[LabelTarget]
    chart_data_path: Path

    @property
    def template_country_names(self) -> list[str]:
        return self.template_layout.country_names


@dataclass(frozen=True)
class GeoCostBuildJob:
    scope: GeoCostScope
    sheet_name: str
    template_path: Path
    requested_function: str = "ALL"
    max_functions: int | None = None
    max_companies: int = 8
    chart_name: str = ""
    label_threshold: float = DEFAULT_LABEL_THRESHOLD
    post_chart_wait: float = 2.0


def clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def compact_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", clean_text(value).lower())


def normalize_letters(value: Any) -> str:
    return re.sub(r"[^A-Za-z]+", "", clean_text(value)).lower()


def to_number(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("$", "").replace("%", "").strip()
        if cleaned.lower() in {"", "-", "na", "n/a", "nan", "none"}:
            return 0.0
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def safe_filename(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_") or "geocost"


def normalize_segment_label(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(r"[_\s]+", "-", text)
    text = re.sub(r"-+", "-", text).strip("-")
    parts = text.split("-")
    if len(parts) >= 2:
        return f"{parts[0].upper()}-{parts[-1].capitalize()}"
    return text


def normalize_high_low(value: Any, segment_bucket: str) -> str:
    text = clean_text(value).capitalize()
    if text in {"High", "Low"}:
        return text
    bucket = normalize_segment_label(segment_bucket).lower()
    if bucket.endswith("-low"):
        return "Low"
    if bucket.endswith("-high"):
        return "High"
    return text


def repair_mojibake(value: Any) -> str:
    text = clean_text(value)
    for _ in range(3):
        try:
            repaired = text.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            break
        if repaired == text:
            break
        text = repaired
    return text


def canonical_country_key(value: Any) -> str:
    text = repair_mojibake(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    key = re.sub(r"[^a-z0-9]+", "", text.lower())
    aliases = {
        "usa": "unitedstates",
        "us": "unitedstates",
        "unitedstatesofamerica": "unitedstates",
        "uk": "unitedkingdom",
        "greatbritain": "unitedkingdom",
        "uae": "unitedarabemirates",
        "czechia": "czechrepublic",
        "russia": "russianfederation",
        "southkorea": "republicofkorea",
        "korearepublicof": "republicofkorea",
        "washingtondc": "districtofcolumbia",
        "dc": "districtofcolumbia",
    }
    return aliases.get(key, key)


def display_label_for_country(template_country: str) -> str:
    return DISPLAY_LABEL_ALIASES.get(canonical_country_key(template_country), template_country)


def font_hex_for_bucket(bucket: str) -> str:
    normalized = normalize_segment_label(bucket).lower()
    return "FFFFFF" if normalized.startswith("nam-") or normalized.endswith("-low") else "000000"


def country_key_set(template_country_names: list[str]) -> set[str]:
    keys = {canonical_country_key(name) for name in template_country_names}
    keys.update(canonical_country_key(display_label_for_country(name)) for name in template_country_names)
    return {key for key in keys if key}


def setup_logger(debug: bool, output_dir: Path | None = None) -> logging.Logger:
    logger = logging.getLogger("geocost_automation")
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG if debug else logging.INFO)

    formatter = logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s", "%H:%M:%S")
    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG if debug else logging.INFO)
    console.setFormatter(formatter)
    logger.addHandler(console)

    if output_dir is not None:
        output_dir.mkdir(parents=True, exist_ok=True)
        log_dir = output_dir / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f"geocost_{time.strftime('%Y%m%d_%H%M%S')}.log"
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
        logger.info("Log file: %s", log_file)
    return logger


def unique_output_path(path: Path) -> Path:
    if not path.exists():
        return path
    return path.with_name(f"{path.stem}_{time.strftime('%Y%m%d_%H%M%S')}{path.suffix}")


def pump_office_ui(seconds: float = 0.5) -> None:
    end = time.time() + max(0.0, float(seconds))
    while time.time() < end:
        try:
            if pythoncom is not None:
                pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        time.sleep(0.05)


def require_windows_com() -> None:
    if pythoncom is None or win32 is None:
        raise RuntimeError(
            "PowerPoint generation requires Windows, pywin32, PowerPoint, Excel, and think-cell. "
            "Use --chart-data-only to validate the workbook and ChartData without Office."
        )


def com_hresult(exc: BaseException) -> int | None:
    value = getattr(exc, "hresult", None)
    if isinstance(value, int):
        return value
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return int(args[0])
    return None


def is_transient_office_error(exc: BaseException) -> bool:
    hresult = com_hresult(exc)
    if hresult in {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER, OBJECT_DISCONNECTED}:
        return True
    text = str(exc).lower()
    phrases = (
        "call was rejected by callee",
        "server busy",
        "retry later",
        "object invoked has disconnected",
        "disconnected from its clients",
        "rpc server is unavailable",
        "application is busy",
    )
    return any(phrase in text for phrase in phrases)


def is_disconnected_office_error(exc: BaseException) -> bool:
    hresult = com_hresult(exc)
    if hresult == OBJECT_DISCONNECTED:
        return True
    text = str(exc).lower()
    return (
        "object invoked has disconnected" in text
        or "disconnected from its clients" in text
        or "rpc server is unavailable" in text
    )


def retry_office_call(
    description: str,
    action: Callable[[], Any],
    logger: logging.Logger,
    max_attempts: int = DEFAULT_OFFICE_RETRY_ATTEMPTS,
    base_delay: float = 0.8,
):
    last_error: BaseException | None = None
    attempts = max(1, int(max_attempts))
    for attempt in range(1, attempts + 1):
        try:
            pump_office_ui(0.2)
            return action()
        except Exception as exc:
            last_error = exc
            if attempt >= attempts or not is_transient_office_error(exc):
                raise
            logger.warning(
                "%s was temporarily rejected by Office (attempt %s/%s); retrying.",
                description,
                attempt,
                attempts,
            )
            logger.debug("Transient Office COM error during %s: %r", description, exc)
            pump_office_ui(base_delay * attempt)
    raise RuntimeError(f"{description} failed after {attempts} attempts.") from last_error


def save_presentation_as_with_retry(
    pres,
    path: Path,
    logger: logging.Logger,
    max_attempts: int = DEFAULT_OFFICE_RETRY_ATTEMPTS,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    retry_office_call(
        f"SaveAs {path.name}",
        lambda: pres.SaveAs(str(path)),
        logger,
        max_attempts=max_attempts,
        base_delay=1.0,
    )


def close_presentation_with_retry(
    pres,
    logger: logging.Logger,
    label: str,
    max_attempts: int = DEFAULT_OFFICE_RETRY_ATTEMPTS,
    raise_on_failure: bool = False,
) -> bool:
    if pres is None:
        return True
    last_error: BaseException | None = None
    attempts = max(1, int(max_attempts))
    for attempt in range(1, attempts + 1):
        try:
            pump_office_ui(0.2)
            pres.Close()
            return True
        except Exception as exc:
            last_error = exc
            if attempt < attempts and is_transient_office_error(exc):
                logger.warning(
                    "PowerPoint was busy while closing %s (attempt %s/%s); retrying.",
                    label,
                    attempt,
                    attempts,
                )
                pump_office_ui(0.8 * attempt)
                continue
            break
    message = f"Could not close {label} after {attempts} attempt(s)."
    if raise_on_failure:
        raise RuntimeError(message) from last_error
    logger.warning("%s The already-saved output remains valid. Error: %r", message, last_error)
    return False


def wait_for_saved_file_stable(
    path: Path,
    logger: logging.Logger,
    timeout_seconds: float = 30.0,
    stable_checks: int = 3,
) -> None:
    deadline = time.time() + max(1.0, float(timeout_seconds))
    last_size = -1
    unchanged = 0
    while time.time() < deadline:
        pump_office_ui(0.25)
        try:
            size = path.stat().st_size
        except FileNotFoundError:
            size = -1
        if size > 0 and size == last_size:
            unchanged += 1
            if unchanged >= max(1, int(stable_checks)):
                logger.info("Saved output is stable: %s (%s bytes).", path, size)
                return
        else:
            unchanged = 0
            last_size = size
    raise RuntimeError(f"PowerPoint output did not stabilize after SaveAs: {path}")
