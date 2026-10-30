from __future__ import annotations

"""Serializable Aura L0 native-patch metadata for final merge and recovery."""

import json
from pathlib import Path
from typing import Any

from l0_config import L0_BUILD_ID


L0_PATCH_MANIFEST_VERSION = 1
DEFAULT_PATCH_MANIFEST_NAME = "01_Aura_L0_Patch_Metadata.json"


def _serializable_anchor(anchor: Any) -> dict[str, Any] | None:
    if not isinstance(anchor, dict):
        return None
    result: dict[str, Any] = {}
    for rank in ("highest", "middle", "lowest"):
        value = anchor.get(rank)
        if isinstance(value, dict):
            result[rank] = value
    return result or None


def patch_record_from_slide_record(record: dict[str, Any]) -> dict[str, Any]:
    """Keep only the fields required by the closed-deck native patch."""
    segments: list[dict[str, Any]] = []
    for segment in record.get("segments") or []:
        segments.append(
            {
                "function": segment.get("function"),
                "l1": segment.get("l1"),
            }
        )
    return {
        "l0": str(record.get("l0") or record.get("function") or ""),
        "segments": segments,
        "quartile_values": [float(value) for value in (record.get("quartile_values") or [])],
        "quartile_anchor": _serializable_anchor(record.get("quartile_anchor")),
    }


def write_l0_patch_manifest(
    path: Path,
    slide_records: list[dict[str, Any]],
    *,
    section_deck_name: str,
    native_patch_pending: bool = True,
) -> Path:
    path = Path(path).resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "manifest_version": L0_PATCH_MANIFEST_VERSION,
        "l0_build_id": L0_BUILD_ID,
        "section_deck_name": section_deck_name,
        "native_patch_pending": bool(native_patch_pending),
        "slide_count": len(slide_records),
        "slide_records": [patch_record_from_slide_record(record) for record in slide_records],
    }
    temp_path = path.with_suffix(path.suffix + ".writing")
    temp_path.write_text(
        json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False),
        encoding="utf-8",
    )
    temp_path.replace(path)
    return path


def read_l0_patch_manifest(path: Path) -> dict[str, Any]:
    path = Path(path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Aura L0 patch metadata not found: {path}")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if int(payload.get("manifest_version", -1)) != L0_PATCH_MANIFEST_VERSION:
        raise ValueError(
            f"Unsupported Aura L0 patch metadata version: {payload.get('manifest_version')!r}"
        )
    if payload.get("l0_build_id") != L0_BUILD_ID:
        raise ValueError(
            "Aura L0 patch metadata was created by a different L0 build: "
            f"metadata={payload.get('l0_build_id')!r}, required={L0_BUILD_ID!r}."
        )
    native_patch_pending = payload.get("native_patch_pending", True)
    if not isinstance(native_patch_pending, bool):
        raise ValueError(
            "Aura L0 patch metadata native_patch_pending must be true or false."
        )
    payload["native_patch_pending"] = native_patch_pending

    records = payload.get("slide_records")
    if not isinstance(records, list) or not records:
        raise ValueError("Aura L0 patch metadata contains no slide records.")
    if int(payload.get("slide_count", -1)) != len(records):
        raise ValueError(
            "Aura L0 patch metadata slide count does not match its record list: "
            f"declared={payload.get('slide_count')}, records={len(records)}."
        )
    for index, record in enumerate(records, start=1):
        if not isinstance(record, dict) or not record.get("l0"):
            raise ValueError(f"Aura L0 patch metadata record {index} has no L0 name.")
        if len(record.get("quartile_values") or []) < 3:
            raise ValueError(
                f"Aura L0 patch metadata record {index} has fewer than three quartile values."
            )
        if not record.get("segments"):
            raise ValueError(f"Aura L0 patch metadata record {index} has no segments.")
    return payload
