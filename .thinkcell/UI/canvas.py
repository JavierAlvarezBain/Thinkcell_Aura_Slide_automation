from __future__ import annotations

"""Right-side canvas for the ThinkCell Slide Maker Streamlit application.

This module owns the formatted workbook preview, live backend console, output
verification, and the visual system shared by both sides of the application.
The control rail lives in ``UI/control_panel.py`` and can collapse from its
expanded 25% layout to a compact 2.5% launcher, with a second 2.5%-collapsible input rail.
"""

import hashlib
import html
import json
import os
import queue
import re
import subprocess
import tempfile
import threading
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime, time as datetime_time
from pathlib import Path
from typing import Any, Iterable
from functools import lru_cache

import streamlit as st
import streamlit.components.v1 as components

from UI.control_panel import (
    AppPaths,
    BuildOptions,
    ControlPanelState,
    WorkbookLoadProgress,
    DECK_SECTION_LABELS,
    SORTABLE_DECK_SECTIONS,
    AURA_INPUT_BY_KEY,
    AURA_INPUT_KEYS,
    AURA_INPUT_MAP_SCHEMA_VERSION,
    aura_input_keys_for_sections,
    GEOCOST_INPUT_BY_KEY,
    GEOCOST_INPUT_KEYS,
    GEOCOST_INPUT_MAP_SCHEMA_VERSION,
    STANDARD_INPUT_BY_KEY,
    STANDARD_INPUT_KEYS,
    standardized_input_keys_for_sections,
    L1_INPUT_BY_KEY,
    L1_INPUT_KEYS,
    L1_INPUT_MAP_SCHEMA_VERSION,
    L1_SORT_ORDER_DEFAULT,
    L1_SORT_ORDER_LABELS,
    L1_SEGMENT_SORT_DIRECTION_DEFAULT,
    L1_SEGMENT_SORT_DIRECTION_LABELS,
    PRESENTATION_SUFFIXES,
    UploadInspection,
    active_l1_input_key as _active_l1_input_key,
    apply_l1_segment_assignment_overrides as _apply_l1_segment_assignment_overrides,
    get_l1_input_mappings as _get_l1_input_mappings,
    get_l1_segment_assignment_overrides as _get_l1_segment_assignment_overrides,
    get_ignored_aura_taxonomy_functions as _get_ignored_aura_taxonomy_functions,
    add_ignored_aura_taxonomy_function as _add_ignored_aura_taxonomy_function,
    remove_ignored_aura_taxonomy_function as _remove_ignored_aura_taxonomy_function,
    set_l1_input_mappings as _set_l1_input_mappings,
    validate_l1_input_mappings,
    validate_geocost_input_mappings,
    validate_standard_input_mappings,
    standard_input_status,
    ui_log,
    ui_stage,
)


MAX_VISIBLE_LOG_LINES = 650
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
PREVIEW_ROWS_AFTER_LAST_TABLE = 5
# The formatted preview is HTML, so rendering an entire style-heavy worksheet
# can create hundreds of thousands of DOM cells and keep Streamlit in a
# perpetual "running" state. Keep the initial viewport bounded and move the
# window to the active mapped range when necessary.
MAX_FORMATTED_PREVIEW_CELLS = 16_000
# Geo Cost tabs are usually much wider than Aura/Taxonomy tabs. The generic
# DOM-cell budget therefore used to reduce a wide Geo Cost preview to only 120
# rows even when meaningful data continued much farther down the sheet. Give
# Geo Cost a larger, still-bounded budget so its full entity table is normally
# visible, while retaining paged navigation for genuinely large worksheets.
GEOCOST_FORMATTED_PREVIEW_CELLS = 64_000
MAX_FORMATTED_PREVIEW_ROWS = 1_200
MAX_FORMATTED_PREVIEW_COLUMNS = 140
PREVIEW_MAPPING_CONTEXT_ROWS = 18
PREVIEW_CONTENT_SCAN_MAX_ROW = 5_000
PREVIEW_CONTENT_SCAN_MAX_COLUMN = 512
UI_BUILD_ID = "root-segment-workspace-restored-v2026-09-21-r78"
WINDOWS_PPTX_PATTERN = re.compile(
    r"[A-Za-z]:\\[^\r\n\"']+?\.(?:pptx|pptm)",
    flags=re.IGNORECASE,
)




@dataclass
class ConsoleTargets:
    progress: object
    log: object
    result: object


_SOURCE_SELECTOR_COMPONENT = None
_SOURCE_SELECTOR_COMPONENT_BUILD = "thinkcell-source-split-selector-v2026-09-15-r1"
_SOURCE_SELECTOR_COMPONENT_HTML = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>ThinkCell source selector</title>
  <style>
    html, body {
      margin: 0;
      padding: 0;
      width: 100%;
      overflow: hidden;
      background: transparent;
      font-family: "Segoe UI", Inter, Arial, sans-serif;
      color: #20242a;
    }
    * { box-sizing: border-box; }
    #root { width: 100%; }
    .tc-source-row {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 8px;
      width: 100%;
      padding: 0;
      margin: 0;
    }
    .tc-source-control {
      min-width: 0;
      height: 46px;
      display: flex;
      align-items: stretch;
      overflow: hidden;
      border: 1px solid #d7dce2;
      border-radius: 8px;
      background: #ffffff;
      box-shadow: none;
      transition: border-color 90ms ease, background 90ms ease, color 90ms ease;
    }
    .tc-source-control:hover {
      border-color: #b8c0ca;
    }
    .tc-source-control.is-active {
      border-color: #f4474d;
      background: #f4474d;
    }
    .tc-preview-button {
      min-width: 0;
      flex: 1 1 auto;
      display: flex;
      flex-direction: column;
      justify-content: center;
      gap: 2px;
      padding: 5px 9px 5px 10px;
      border: 0;
      outline: 0;
      background: transparent;
      text-align: left;
      cursor: pointer;
    }
    .tc-preview-button:disabled {
      cursor: default;
    }
    .tc-source-title {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      color: #20242a;
      font-size: 11px;
      font-weight: 750;
      line-height: 1.05;
    }
    .tc-source-status {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      color: #66707c;
      font-size: 9.5px;
      font-weight: 520;
      line-height: 1.05;
    }
    .tc-source-control.is-active .tc-source-title,
    .tc-source-control.is-active .tc-source-status {
      color: #ffffff;
    }
    .tc-source-control.is-unavailable .tc-preview-button {
      opacity: 0.55;
    }
    .tc-menu-zone {
      position: relative;
      flex: 0 0 38px;
      width: 38px;
      display: flex;
      align-items: center;
      justify-content: center;
      border-left: 1px solid #e0e4e8;
      background: rgba(0, 0, 0, 0.012);
    }
    .tc-source-control.is-active .tc-menu-zone {
      border-left-color: rgba(255, 255, 255, 0.38);
      background: rgba(112, 0, 0, 0.08);
    }
    .tc-menu-zone:hover {
      background: rgba(0, 0, 0, 0.045);
    }
    .tc-source-control.is-active .tc-menu-zone:hover {
      background: rgba(112, 0, 0, 0.15);
    }
    .tc-menu-select {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      margin: 0;
      opacity: 0;
      cursor: pointer;
      z-index: 3;
    }
    .tc-chevron {
      width: 8px;
      height: 8px;
      border-right: 1.6px solid #3f4751;
      border-bottom: 1.6px solid #3f4751;
      transform: rotate(45deg) translate(-1px, -1px);
      pointer-events: none;
    }
    .tc-source-control.is-active .tc-chevron {
      border-color: #ffffff;
    }
    .tc-source-control:focus-within {
      box-shadow: 0 0 0 2px rgba(201, 0, 0, 0.15);
    }
    @media (max-width: 840px) {
      .tc-source-row { gap: 5px; }
      .tc-source-title { font-size: 10px; }
      .tc-source-status { font-size: 8.5px; }
      .tc-menu-zone { flex-basis: 34px; width: 34px; }
    }
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

      function emit(action, role, value = "") {
        setComponentValue({
          event_id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
          action,
          role,
          value,
        });
      }

      function render() {
        const controls = Array.isArray(args.controls) ? args.controls : [];
        const previewEnabled = Boolean(args.preview_enabled);
        const activeRole = String(args.active_role || "");

        root.innerHTML = "";
        const row = document.createElement("div");
        row.className = "tc-source-row";

        controls.forEach((control) => {
          const role = String(control.role || "");
          const title = String(control.title || "");
          const status = String(control.status || "");
          const actualSheet = String(control.actual_sheet || "");
          const available = Boolean(actualSheet);
          const active = previewEnabled && role === activeRole;

          const shell = document.createElement("div");
          shell.className = "tc-source-control";
          if (active) shell.classList.add("is-active");
          if (!available) shell.classList.add("is-unavailable");
          shell.title = actualSheet
            ? `${title}: ${status}. Previewing uses ${actualSheet}.`
            : `${title}: ${status}. Choose a worksheet from the arrow menu.`;

          const previewButton = document.createElement("button");
          previewButton.type = "button";
          previewButton.className = "tc-preview-button";
          previewButton.disabled = !previewEnabled || !available;
          previewButton.setAttribute("aria-label", `Preview ${title}`);
          previewButton.addEventListener("click", () => emit("preview", role));

          const titleNode = document.createElement("span");
          titleNode.className = "tc-source-title";
          titleNode.textContent = title;

          const statusNode = document.createElement("span");
          statusNode.className = "tc-source-status";
          statusNode.textContent = status;

          previewButton.appendChild(titleNode);
          previewButton.appendChild(statusNode);

          const menuZone = document.createElement("div");
          menuZone.className = "tc-menu-zone";
          menuZone.title = `Choose ${title} worksheet`;

          const select = document.createElement("select");
          select.className = "tc-menu-select";
          select.setAttribute("aria-label", `Choose ${title} worksheet`);
          const options = Array.isArray(control.options) ? control.options : [];
          options.forEach((option) => {
            const optionNode = document.createElement("option");
            optionNode.value = String(option.value || "");
            optionNode.textContent = String(option.label || option.value || "");
            select.appendChild(optionNode);
          });
          select.value = String(control.choice || "");
          select.addEventListener("change", () => {
            emit("select", role, select.value);
          });

          const chevron = document.createElement("span");
          chevron.className = "tc-chevron";
          chevron.setAttribute("aria-hidden", "true");

          menuZone.appendChild(select);
          menuZone.appendChild(chevron);
          shell.appendChild(previewButton);
          shell.appendChild(menuZone);
          row.appendChild(shell);
        });

        root.appendChild(row);
        setFrameHeight(48);
      }

      window.addEventListener("message", (event) => {
        if (!event.data || event.data.type !== "streamlit:render") return;
        args = event.data.args || {};
        render();
      });

      sendMessage("streamlit:componentReady", { apiVersion: 1 });
      setFrameHeight(48);
    })();
  </script>
</body>
</html>
"""


def _get_source_selector_component():
    """Return the compact split-button source selector component."""
    global _SOURCE_SELECTOR_COMPONENT
    if _SOURCE_SELECTOR_COMPONENT is not None:
        return _SOURCE_SELECTOR_COMPONENT

    component_dir = (
        Path(tempfile.gettempdir())
        / "ThinkCellSlideMaker"
        / _SOURCE_SELECTOR_COMPONENT_BUILD
    )
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        current = index_path.read_text(encoding="utf-8")
    except OSError:
        current = ""
    if current != _SOURCE_SELECTOR_COMPONENT_HTML:
        temporary_path = index_path.with_suffix(".html.writing")
        temporary_path.write_text(_SOURCE_SELECTOR_COMPONENT_HTML, encoding="utf-8")
        temporary_path.replace(index_path)

    _SOURCE_SELECTOR_COMPONENT = components.declare_component(
        "thinkcell_source_split_selector_v1",
        path=str(component_dir),
    )
    return _SOURCE_SELECTOR_COMPONENT

_RANGE_SELECTOR_COMPONENT = None
_RANGE_SELECTOR_COMPONENT_BUILD = "thinkcell-excel-range-selector-v2026-09-24-r13"
_RANGE_SELECTOR_COMPONENT_HTML = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>ThinkCell workbook range selector</title>
  <style>
    html, body {
      margin: 0;
      padding: 0;
      width: 100%;
      overflow: hidden;
      background: transparent;
      font-family: "Segoe UI", Inter, Arial, sans-serif;
      color: #20242a;
    }
    #root { width: 100%; }
    .tc-component-shell,
    .tc-component-shell * { box-sizing: border-box; }
    .tc-range-toolbar {
      min-height: 46px;
      display: grid;
      grid-template-columns: minmax(118px, 170px) minmax(0, 1fr) auto auto;
      gap: 10px;
      align-items: center;
      padding: 6px 8px;
      margin: 0 0 6px 0;
      border: 1px solid #dfe3e8;
      border-radius: 12px;
      background: #ffffff;
      box-shadow: 0 3px 12px rgba(18, 24, 33, 0.04);
    }
    .tc-range-name-box {
      min-width: 0;
      height: 32px;
      display: flex;
      align-items: center;
      padding: 0 11px;
      border: 1px solid #cfd5dc;
      border-radius: 8px;
      background: #f8f9fb;
      color: #1d232b;
      font-family: Consolas, "SFMono-Regular", monospace;
      font-size: 12px;
      font-weight: 700;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .tc-range-copy { min-width: 0; }
    .tc-range-title {
      color: #1d232b;
      font-size: 12px;
      line-height: 1.25;
      font-weight: 750;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .tc-range-help {
      margin-top: 2px;
      color: #7b838e;
      font-size: 10.5px;
      line-height: 1.3;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .tc-zoom-controls,
    .tc-range-actions {
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .tc-zoom-controls {
      padding: 2px;
      border: 1px solid #d7dce2;
      border-radius: 9px;
      background: #f8f9fb;
    }
    .tc-zoom-button,
    .tc-zoom-value,
    .tc-range-button {
      height: 32px;
      border: 1px solid #d7dce2;
      border-radius: 8px;
      background: #ffffff;
      color: #3b424c;
      font-size: 10.5px;
      font-weight: 700;
      cursor: pointer;
      transition: background 100ms ease, border-color 100ms ease, color 100ms ease;
    }
    .tc-zoom-button {
      width: 32px;
      padding: 0;
      font-size: 17px;
      line-height: 1;
    }
    .tc-zoom-value {
      min-width: 56px;
      padding: 0 8px;
      border-color: transparent;
      background: transparent;
      font-variant-numeric: tabular-nums;
    }
    .tc-zoom-button:hover,
    .tc-zoom-value:hover,
    .tc-range-button:hover:not(:disabled) {
      background: #f1f3f6;
      border-color: #bfc6ce;
    }
    .tc-range-button {
      padding: 0 11px;
    }
    .tc-range-button.tc-visibility-button {
      min-width: 84px;
      color: #303741;
      border-color: #cfd5dc;
      background: #f8f9fb;
    }
    .tc-range-button.tc-visibility-button:hover:not(:disabled) {
      color: #20262e;
      border-color: #b9c0c8;
      background: #eef1f4;
    }
    .tc-range-button.tc-copy-button {
      color: #a50000;
      border-color: #e3baba;
      background: #fff7f7;
    }
    .tc-range-button.tc-copy-button:hover:not(:disabled) {
      color: #870000;
      border-color: #d69b9b;
      background: #fff0f0;
    }
    .tc-range-button.tc-assign-button {
      min-width: 94px;
      color: #ffffff;
      border-color: #b50000;
      background: #c90000;
      box-shadow: 0 2px 7px rgba(201, 0, 0, 0.16);
    }
    .tc-range-button.tc-assign-button:hover:not(:disabled) {
      color: #ffffff;
      border-color: #930000;
      background: #a90000;
    }
    .tc-range-button.tc-assign-button:disabled {
      color: #ffffff;
      border-color: #c8cdd3;
      background: #aeb4bb;
      box-shadow: none;
    }
    .tc-range-button:disabled {
      opacity: 0.45;
      cursor: default;
    }
    .tc-selection-status {
      display: flex;
      visibility: hidden;
      align-items: center;
      gap: 7px;
      min-height: 24px;
      margin: 0 0 6px 0;
      padding: 4px 8px;
      border-radius: 8px;
      color: #5f6874;
      background: #f7f8fa;
      border: 1px solid #e5e8ec;
      font-size: 10.5px;
      line-height: 1.35;
    }
    .tc-selection-status.visible { visibility: visible; }
    .tc-selection-status.status-ok {
      color: #315b42;
      background: #f3fbf6;
      border-color: #cfe8d7;
    }
    .tc-selection-status.status-warning {
      color: #72591d;
      background: #fff9e8;
      border-color: #ead79b;
    }
    .tc-selection-status.status-error {
      color: #8d2020;
      background: #fff4f4;
      border-color: #efcaca;
    }
    .tc-selection-status.status-info {
      color: #5f6874;
      background: #f7f8fa;
      border-color: #e5e8ec;
    }
    .tc-selection-status-dot {
      flex: 0 0 auto;
      width: 7px;
      height: 7px;
      border-radius: 999px;
      background: #c90000;
    }
    .tc-selection-status.status-ok .tc-selection-status-dot { background: #24a35a; }
    .tc-selection-status.status-warning .tc-selection-status-dot { background: #e4a400; }
    .tc-selection-status.status-error .tc-selection-status-dot { background: #c90000; }
    .tc-selection-status.status-info .tc-selection-status-dot { background: #c90000; }
    .tc-selection-status-message {
      flex: 1 1 auto;
      min-width: 0;
    }
    .tc-selection-status-action {
      flex: 0 0 auto;
      margin-left: auto;
      padding: 4px 9px;
      border-radius: 999px;
      border: 1px solid currentColor;
      background: rgba(255,255,255,0.72);
      color: inherit;
      font: inherit;
      font-weight: 800;
      cursor: pointer;
      white-space: nowrap;
    }
    .tc-selection-status-action:hover {
      background: #ffffff;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .tc-component-shell .excel-viewport {
      border-radius: 12px !important;
      box-shadow: none !important;
      cursor: cell;
      user-select: none;
      -webkit-user-select: none;
      outline: none;
      overscroll-behavior: contain;
    }
    .tc-component-shell .excel-viewport:focus {
      border-color: #bdc5cf !important;
      box-shadow: inset 0 0 0 1px rgba(201, 0, 0, 0.14) !important;
    }
    .tc-component-shell .excel-table {
      user-select: none;
      -webkit-user-select: none;
      zoom: var(--tc-excel-zoom, 1);
      transform-origin: top left;
    }
    .tc-component-shell .excel-cell,
    .tc-component-shell .row-header,
    .tc-component-shell .column-header,
    .tc-component-shell .corner-cell {
      cursor: cell;
    }
    .tc-component-shell .row-header,
    .tc-component-shell .column-header {
      transition: background 80ms ease, color 80ms ease;
    }
    .tc-component-shell .excel-cell.tc-range-selected::before {
      content: "";
      position: absolute;
      inset: 0;
      z-index: 2;
      pointer-events: none;
      background: rgba(201, 0, 0, 0.105);
    }
    .tc-component-shell .excel-cell.tc-range-anchor::before {
      background: rgba(201, 0, 0, 0.16);
    }
    .tc-component-shell .excel-cell.tc-active-cell::after {
      content: "";
      position: absolute;
      inset: 2px;
      z-index: 3;
      pointer-events: none;
      border: 1px solid rgba(143, 0, 0, 0.9);
    }
    .tc-component-shell .excel-cell.tc-range-top {
      border-top: 2px solid #c90000 !important;
    }
    .tc-component-shell .excel-cell.tc-range-bottom {
      border-bottom: 2px solid #c90000 !important;
    }
    .tc-component-shell .excel-cell.tc-range-left {
      border-left: 2px solid #c90000 !important;
    }
    .tc-component-shell .excel-cell.tc-range-right {
      border-right: 2px solid #c90000 !important;
    }
    .tc-component-shell .row-header.tc-header-selected,
    .tc-component-shell .column-header.tc-header-selected {
      color: #8f0000 !important;
      background: #fdeaea !important;
      font-weight: 750 !important;
    }
    .tc-component-shell.tc-dragging,
    .tc-component-shell.tc-dragging * {
      cursor: crosshair !important;
    }
    .tc-component-shell .tc-copy-feedback {
      color: #14704f;
      font-weight: 700;
    }
    @media (max-width: 920px) {
      .tc-range-toolbar {
        grid-template-columns: minmax(100px, 145px) minmax(0, 1fr) auto;
      }
      .tc-range-actions { grid-column: 2 / -1; justify-content: flex-end; }
    }
    @media (max-width: 700px) {
      .tc-range-toolbar {
        grid-template-columns: minmax(100px, 145px) minmax(0, 1fr);
      }
      .tc-zoom-controls,
      .tc-range-actions { justify-content: flex-end; }
    }
  </style>
</head>
<body>
  <div id="root"></div>
  <script>
    (() => {
      const root = document.getElementById("root");
      const MIN_ZOOM = 40;
      const MAX_ZOOM = 180;
      const ZOOM_STEP = 10;
      let args = {};
      let shell = null;
      let viewport = null;
      let selection = null;
      let anchor = null;
      let activeCell = null;
      let dragMode = null;
      let dragging = false;
      let lastPointer = null;
      let autoScrollTimer = null;
      let draftSaveTimer = null;
      let cellLookup = new Map();
      let zoomPercent = 100;
      let committedSelectionSignature = "";
      let selectionDirty = false;
      let rowVisibilityState = { hiddenRows: new Set(), shownRows: new Set() };

      function sendMessage(type, data = {}) {
        window.parent.postMessage(
          Object.assign({ isStreamlitMessage: true, type }, data),
          "*"
        );
      }

      function setFrameHeight(height) {
        sendMessage("streamlit:setFrameHeight", { height });
      }

      function availableFrameHeight() {
        const configured = Math.max(620, Number(args.component_height || 680));
        try {
          const frame = window.frameElement;
          const parentHeight = Number(window.parent.innerHeight || 0);
          const frameTop = frame ? Number(frame.getBoundingClientRect().top || 0) : 0;
          if (parentHeight > 0 && frameTop >= 0) {
            return Math.max(560, Math.floor(parentHeight - frameTop - 8));
          }
        } catch (_error) {
          // Cross-origin/sandboxed runtimes use the configured fallback.
        }
        return configured;
      }

      function fitPreviewToViewport() {
        const requestedHeight = availableFrameHeight();
        if (shell && viewport) {
          const toolbar = shell.querySelector(".tc-range-toolbar");
          const status = shell.querySelector(".tc-selection-status");
          const chromeHeight =
            (toolbar ? toolbar.offsetHeight : 54) +
            (status ? status.offsetHeight : 28) +
            18;
          viewport.style.height = `${Math.max(360, requestedHeight - chromeHeight)}px`;
        }
        if (Math.abs(requestedHeight - appliedFrameHeight) > 1) {
          appliedFrameHeight = requestedHeight;
          setFrameHeight(requestedHeight);
        }
      }

      let resizeTimer = null;
      let appliedFrameHeight = 0;
      function scheduleViewportFit() {
        if (resizeTimer !== null) window.clearTimeout(resizeTimer);
        resizeTimer = window.setTimeout(() => {
          resizeTimer = null;
          fitPreviewToViewport();
        }, 60);
      }

      function setComponentValue(value) {
        sendMessage("streamlit:setComponentValue", {
          value,
          dataType: "json",
        });
      }

      function columnLabel(value) {
        let column = Math.max(1, Number(value) || 1);
        let result = "";
        while (column > 0) {
          const remainder = (column - 1) % 26;
          result = String.fromCharCode(65 + remainder) + result;
          column = Math.floor((column - 1) / 26);
        }
        return result;
      }

      function cellReference(row, column) {
        return `${columnLabel(column)}${row}`;
      }

      function rangeReference(value) {
        if (!value) return "";
        const first = cellReference(value.r1, value.c1);
        const last = cellReference(value.r2, value.c2);
        return first === last ? first : `${first}:${last}`;
      }

      function quoteSheetName(name) {
        return `'${String(name || "").replaceAll("'", "''")}'`;
      }

      function escapeHtml(value) {
        return String(value || "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");
      }

      function visibleLimits() {
        const startRow = Number(args.visible_start_row || 1);
        const endRow = Number(args.visible_end_row || startRow);
        const startColumn = Number(args.visible_start_column || 1);
        const endColumn = Number(args.visible_end_column || startColumn);
        return { startRow, endRow, startColumn, endColumn };
      }

      function boundsForCell(cell) {
        return {
          r1: Number(cell.dataset.row),
          c1: Number(cell.dataset.col),
          r2: Number(cell.dataset.rowEnd || cell.dataset.row),
          c2: Number(cell.dataset.colEnd || cell.dataset.col),
        };
      }

      function normalizedBounds(first, second) {
        return {
          r1: Math.min(first.r1, second.r1),
          c1: Math.min(first.c1, second.c1),
          r2: Math.max(first.r2, second.r2),
          c2: Math.max(first.c2, second.c2),
        };
      }

      function clamp(value, minimum, maximum) {
        return Math.max(minimum, Math.min(maximum, Number(value) || minimum));
      }

      function coordinateKey(row, column) {
        return `${Number(row)}:${Number(column)}`;
      }

      function rebuildCellLookup() {
        cellLookup = new Map();
        if (!shell) return;
        shell.querySelectorAll("td.excel-cell[data-row][data-col]").forEach((cell) => {
          const bounds = boundsForCell(cell);
          for (let row = bounds.r1; row <= bounds.r2; row += 1) {
            for (let column = bounds.c1; column <= bounds.c2; column += 1) {
              cellLookup.set(coordinateKey(row, column), cell);
            }
          }
        });
      }

      function cellForCoordinate(row, column) {
        return cellLookup.get(coordinateKey(row, column)) || null;
      }

      function boundsForCoordinate(row, column) {
        const limits = visibleLimits();
        const safeRow = clamp(row, limits.startRow, limits.endRow);
        const safeColumn = clamp(column, limits.startColumn, limits.endColumn);
        const cell = cellForCoordinate(safeRow, safeColumn);
        return cell
          ? boundsForCell(cell)
          : { r1: safeRow, c1: safeColumn, r2: safeRow, c2: safeColumn };
      }

      function selectionFromTarget(target, mode) {
        const limits = visibleLimits();
        if (mode === "cell") {
          const cell = target && target.closest ? target.closest("td.excel-cell") : null;
          return cell ? boundsForCell(cell) : null;
        }
        if (mode === "row") {
          const header = target && target.closest ? target.closest("th.row-header") : null;
          if (!header) return null;
          const row = Number(header.dataset.row);
          return { r1: row, c1: limits.startColumn, r2: row, c2: limits.endColumn };
        }
        if (mode === "column") {
          const header = target && target.closest ? target.closest("th.column-header") : null;
          if (!header) return null;
          const column = Number(header.dataset.col);
          return { r1: limits.startRow, c1: column, r2: limits.endRow, c2: column };
        }
        if (mode === "all") {
          return {
            r1: limits.startRow,
            c1: limits.startColumn,
            r2: limits.endRow,
            c2: limits.endColumn,
          };
        }
        return null;
      }

      function targetMode(target) {
        if (!target || !target.closest) return null;
        if (target.closest("td.excel-cell")) return "cell";
        if (target.closest("th.row-header")) return "row";
        if (target.closest("th.column-header")) return "column";
        if (target.closest("th.corner-cell")) return "all";
        return null;
      }

      function intersects(cellBounds, selectedBounds) {
        return !(
          cellBounds.r2 < selectedBounds.r1 ||
          cellBounds.r1 > selectedBounds.r2 ||
          cellBounds.c2 < selectedBounds.c1 ||
          cellBounds.c1 > selectedBounds.c2
        );
      }

      function zoomStorageKey() {
        return `thinkcell-range-zoom:${String(args.source_signature || "")}:${String(args.sheet_name || "")}`;
      }

      function focusStorageKey() {
        return `thinkcell-range-focus:${String(args.source_signature || "")}:${String(args.sheet_name || "")}`;
      }

      function draftStorageKey() {
        const inputKey = String(args.active_input_key || "free-selection");
        return `thinkcell-range-draft:${String(args.source_signature || "")}:${String(args.sheet_name || "")}:${inputKey}`;
      }

      function rowVisibilityStorageKey() {
        return `thinkcell-row-visibility:${String(args.source_signature || "")}:${String(args.sheet_name || "")}`;
      }

      function normalizeStoredRows(values) {
        const rows = Array.isArray(values) ? values : [];
        return new Set(
          rows
            .map((value) => Number(value))
            .filter(
              (value) => Number.isFinite(value) && value >= 1 && value <= 1048576
            )
        );
      }

      function loadRowVisibilityState() {
        const stored = storedJson(rowVisibilityStorageKey()) || {};
        rowVisibilityState = {
          hiddenRows: normalizeStoredRows(stored.hidden_rows),
          shownRows: normalizeStoredRows(stored.shown_rows),
        };
      }

      function persistRowVisibilityState() {
        try {
          window.localStorage.setItem(
            rowVisibilityStorageKey(),
            JSON.stringify({
              hidden_rows: Array.from(rowVisibilityState.hiddenRows).sort((a, b) => a - b),
              shown_rows: Array.from(rowVisibilityState.shownRows).sort((a, b) => a - b),
            })
          );
        } catch (_error) {
          // Storage can be disabled in hardened browser profiles.
        }
      }

      function rowElement(rowNumber) {
        if (!shell) return null;
        return shell.querySelector(`tr[data-row="${Number(rowNumber)}"]`);
      }

      function rowIsHidden(rowNumber) {
        const row = rowElement(rowNumber);
        return Boolean(row && row.style.display === "none");
      }

      function applyRowVisibilityState() {
        if (!shell) return;
        shell.querySelectorAll("tr[data-row]").forEach((row) => {
          const rowNumber = Number(row.dataset.row);
          const sourceHidden = String(row.dataset.sourceHidden || "0") === "1";
          const forceHidden = rowVisibilityState.hiddenRows.has(rowNumber);
          const forceShown = rowVisibilityState.shownRows.has(rowNumber);
          row.style.display = forceHidden || (sourceHidden && !forceShown) ? "none" : "";
        });
      }

      function selectedRowNumbers() {
        if (!selection) return [];
        const rows = [];
        for (let row = selection.r1; row <= selection.r2; row += 1) rows.push(row);
        return rows;
      }

      function selectionContainsHiddenRows() {
        return selectedRowNumbers().some((row) => rowIsHidden(row));
      }

      function toggleSelectedRowsVisibility() {
        const rows = selectedRowNumbers();
        if (!rows.length) return;
        const shouldShow = rows.some((row) => rowIsHidden(row));
        rows.forEach((row) => {
          if (shouldShow) {
            rowVisibilityState.hiddenRows.delete(row);
            rowVisibilityState.shownRows.add(row);
          } else {
            rowVisibilityState.shownRows.delete(row);
            rowVisibilityState.hiddenRows.add(row);
          }
        });
        persistRowVisibilityState();
        applyRowVisibilityState();
        rebuildCellLookup();
        applyHighlight();
        fitPreviewToViewport();
      }

      function selectionSignature(value) {
        if (!value) return "";
        const r1 = Number(value.r1 ?? value.start_row);
        const c1 = Number(value.c1 ?? value.start_column);
        const r2 = Number(value.r2 ?? value.end_row);
        const c2 = Number(value.c2 ?? value.end_column);
        if (![r1, c1, r2, c2].every((item) => Number.isFinite(item) && item >= 1)) {
          return "";
        }
        return `${Math.min(r1, r2)}:${Math.min(c1, c2)}:${Math.max(r1, r2)}:${Math.max(c1, c2)}`;
      }

      function storedJson(key) {
        try {
          const raw = window.localStorage.getItem(key);
          return raw ? JSON.parse(raw) : null;
        } catch (_error) {
          return null;
        }
      }

      function clearDraftSelection() {
        if (draftSaveTimer !== null) {
          window.clearTimeout(draftSaveTimer);
          draftSaveTimer = null;
        }
        try {
          window.localStorage.removeItem(draftStorageKey());
        } catch (_error) {
          // Storage can be disabled in hardened browser profiles.
        }
      }

      function rememberDraftSelection() {
        if (!selection || !selectionDirty) {
          if (!selectionDirty) clearDraftSelection();
          return;
        }
        try {
          window.localStorage.setItem(
            draftStorageKey(),
            JSON.stringify(selectionPayload(false))
          );
        } catch (_error) {
          // Storage can be disabled in hardened browser profiles.
        }
      }

      function scheduleRememberDraft(delay = 120) {
        if (draftSaveTimer !== null) window.clearTimeout(draftSaveTimer);
        draftSaveTimer = window.setTimeout(() => {
          draftSaveTimer = null;
          rememberDraftSelection();
        }, Math.max(0, Number(delay) || 0));
      }

      function storedNumber(key, fallback) {
        try {
          const value = Number(window.localStorage.getItem(key));
          return Number.isFinite(value) && value > 0 ? value : fallback;
        } catch (_error) {
          return fallback;
        }
      }

      function rememberFocusForRerun() {
        try {
          window.localStorage.setItem(focusStorageKey(), "1");
        } catch (_error) {
          // Storage can be disabled in hardened browser profiles.
        }
      }

      function consumeFocusRestore() {
        try {
          const shouldRestore = window.localStorage.getItem(focusStorageKey()) === "1";
          if (shouldRestore) window.localStorage.removeItem(focusStorageKey());
          return shouldRestore;
        } catch (_error) {
          return false;
        }
      }

      function setZoom(nextZoom, persist = true) {
        if (!shell) return;
        const previousZoom = Math.max(MIN_ZOOM, Number(zoomPercent) || 100);
        const requested = Math.round(Number(nextZoom) / ZOOM_STEP) * ZOOM_STEP;
        const next = clamp(requested, MIN_ZOOM, MAX_ZOOM);
        if (next === zoomPercent) return;

        let centerX = 0;
        let centerY = 0;
        if (viewport) {
          const previousScale = previousZoom / 100;
          centerX = (viewport.scrollLeft + viewport.clientWidth / 2) / previousScale;
          centerY = (viewport.scrollTop + viewport.clientHeight / 2) / previousScale;
        }

        zoomPercent = next;
        shell.style.setProperty("--tc-excel-zoom", String(zoomPercent / 100));
        const label = shell.querySelector("#tc-zoom-value");
        if (label) {
          label.textContent = `${zoomPercent}%`;
          label.title = "Reset worksheet zoom to 100%";
        }
        const zoomOut = shell.querySelector("#tc-zoom-out");
        const zoomIn = shell.querySelector("#tc-zoom-in");
        if (zoomOut) zoomOut.disabled = zoomPercent <= MIN_ZOOM;
        if (zoomIn) zoomIn.disabled = zoomPercent >= MAX_ZOOM;

        if (persist) {
          try {
            window.localStorage.setItem(zoomStorageKey(), String(zoomPercent));
          } catch (_error) {
            // Storage can be disabled in hardened browser profiles.
          }
        }

        window.requestAnimationFrame(() => {
          if (viewport) {
            const nextScale = zoomPercent / 100;
            viewport.scrollLeft = Math.max(0, centerX * nextScale - viewport.clientWidth / 2);
            viewport.scrollTop = Math.max(0, centerY * nextScale - viewport.clientHeight / 2);
            ensureActiveCellVisible();
            viewport.focus({ preventScroll: true });
          }
        });
      }

      function updateToolbar() {
        if (!shell) return;
        const nameBox = shell.querySelector("#tc-range-name");
        const title = shell.querySelector("#tc-range-title");
        const help = shell.querySelector("#tc-range-help");
        const status = shell.querySelector("#tc-selection-status");
        const assignButton = shell.querySelector("#tc-assign-range");
        const visibilityButton = shell.querySelector("#tc-toggle-rows");
        const copyButton = shell.querySelector("#tc-copy-range");
        const clearButton = shell.querySelector("#tc-clear-range");
        const activeLabel = String(args.active_input_label || "").trim();
        const activeStatus = String(args.active_input_status || "").trim().toLowerCase();
        const activeStatusMessage = String(args.active_input_status_message || "").trim();
        const ignoreMismatchName = String(args.active_input_ignore_mismatch || "").trim();
        const restoreMismatchName = String(args.active_input_restore_mismatch || "").trim();

        function emitMismatchAction(action, name) {
          if (!name) return;
          rememberFocusForRerun();
          setComponentValue({
            version: 1,
            action,
            mismatch_name: name,
            sheet_name: String(args.sheet_name || ""),
            range_a1: selection ? rangeReference(selection) : "",
            event_id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
          });
        }

        function showStatus(level, message, actionName = "", actionLabel = "") {
          if (!status) return;
          const normalized = ["ok", "warning", "error", "info"].includes(level) ? level : "info";
          status.className = `tc-selection-status visible status-${normalized}`;
          const actionHtml = actionName
            ? `<button type="button" class="tc-selection-status-action">${escapeHtml(actionLabel || "Ignore mismatch")}</button>`
            : "";
          status.innerHTML = `<span class="tc-selection-status-dot"></span><span class="tc-selection-status-message">${escapeHtml(message || "")}</span>${actionHtml}`;
          if (actionName) {
            const actionButton = status.querySelector(".tc-selection-status-action");
            if (actionButton) {
              actionButton.addEventListener("click", (event) => {
                event.preventDefault();
                event.stopPropagation();
                if (actionName === "ignore_taxonomy_mismatch") {
                  emitMismatchAction(actionName, ignoreMismatchName);
                } else if (actionName === "restore_taxonomy_mismatch") {
                  emitMismatchAction(actionName, restoreMismatchName);
                }
              });
            }
          }
        }

        if (!selection) {
          nameBox.textContent = "No selection";
          title.textContent = activeLabel ? `Map ${activeLabel}` : "Select a worksheet range";
          help.textContent = activeLabel
            ? "Select and refine the cells, then click Assign range."
            : "Select and refine the cells, then click Use range.";
          if (activeLabel && activeStatusMessage) {
            if (ignoreMismatchName && activeStatus === "error") {
              showStatus(activeStatus, activeStatusMessage, "ignore_taxonomy_mismatch", `Ignore ${ignoreMismatchName}`);
            } else if (restoreMismatchName && (activeStatus === "warning" || activeStatus === "ok")) {
              showStatus(activeStatus, activeStatusMessage, "restore_taxonomy_mismatch", `Re-enable ${restoreMismatchName}`);
            } else {
              showStatus(activeStatus || "info", activeStatusMessage);
            }
          } else {
            status.className = "tc-selection-status";
            status.innerHTML = "";
          }
          assignButton.textContent = activeLabel ? "Assign range" : "Use range";
          assignButton.disabled = true;
          if (visibilityButton) {
            visibilityButton.textContent = "Hide rows";
            visibilityButton.disabled = true;
            visibilityButton.title = "Select one or more worksheet rows first. This changes only the canvas preview.";
          }
          copyButton.disabled = true;
          clearButton.disabled = true;
          return;
        }

        const rangeA1 = rangeReference(selection);
        const rowCount = selection.r2 - selection.r1 + 1;
        const columnCount = selection.c2 - selection.c1 + 1;
        const cellCount = rowCount * columnCount;
        const qualified = `${quoteSheetName(args.sheet_name)}!${rangeA1}`;
        nameBox.textContent = rangeA1;
        nameBox.title = qualified;
        title.textContent = qualified;
        help.textContent = `${rowCount.toLocaleString()} row${rowCount === 1 ? "" : "s"} × ${columnCount.toLocaleString()} column${columnCount === 1 ? "" : "s"} · ${cellCount.toLocaleString()} cell${cellCount === 1 ? "" : "s"} · ${zoomPercent}% zoom`;
        if (dragging) {
          showStatus("warning", "Release the mouse to finish selecting. Nothing has been assigned yet.");
        } else if (selectionDirty) {
          showStatus("warning", `Selection ready. Click ${activeLabel ? "Assign range" : "Use range"} or press Ctrl+Enter.`);
        } else if (activeLabel && activeStatusMessage) {
          if (ignoreMismatchName && activeStatus === "error") {
            showStatus(activeStatus, activeStatusMessage, "ignore_taxonomy_mismatch", `Ignore ${ignoreMismatchName}`);
          } else if (restoreMismatchName && (activeStatus === "warning" || activeStatus === "ok")) {
            showStatus(activeStatus, activeStatusMessage, "restore_taxonomy_mismatch", `Re-enable ${restoreMismatchName}`);
          } else {
            showStatus(activeStatus || "info", activeStatusMessage);
          }
        } else {
          showStatus("info", activeLabel ? `Assigned to ${activeLabel}. Select a different range to change it.` : "Range saved. Select a different range to change it.");
        }

        assignButton.textContent = activeLabel ? "Assign range" : "Use range";
        assignButton.disabled = !selectionDirty;
        if (visibilityButton) {
          const hasHiddenRows = selectionContainsHiddenRows();
          visibilityButton.textContent = hasHiddenRows ? "Unhide rows" : "Hide rows";
          visibilityButton.disabled = false;
          visibilityButton.title = hasHiddenRows
            ? "Show hidden rows inside the selected range in this canvas preview. The source workbook is unchanged."
            : "Hide the rows covered by the selected range in this canvas preview. The source workbook is unchanged.";
        }
        copyButton.disabled = false;
        clearButton.disabled = false;
      }

      function applyHighlight() {
        if (!shell) return;
        const activeElement = activeCell
          ? cellForCoordinate(activeCell.row, activeCell.column)
          : null;
        const cells = shell.querySelectorAll("td.excel-cell[data-row][data-col]");
        cells.forEach((cell) => {
          const cellBounds = boundsForCell(cell);
          const selected = Boolean(selection && intersects(cellBounds, selection));
          cell.classList.toggle("tc-range-selected", selected);
          cell.classList.toggle(
            "tc-range-anchor",
            selected && anchor && intersects(cellBounds, anchor)
          );
          cell.classList.toggle("tc-active-cell", cell === activeElement);
          cell.classList.toggle(
            "tc-range-top",
            selected && cellBounds.r1 <= selection.r1 && cellBounds.r2 >= selection.r1
          );
          cell.classList.toggle(
            "tc-range-bottom",
            selected && cellBounds.r1 <= selection.r2 && cellBounds.r2 >= selection.r2
          );
          cell.classList.toggle(
            "tc-range-left",
            selected && cellBounds.c1 <= selection.c1 && cellBounds.c2 >= selection.c1
          );
          cell.classList.toggle(
            "tc-range-right",
            selected && cellBounds.c1 <= selection.c2 && cellBounds.c2 >= selection.c2
          );
          cell.setAttribute("aria-selected", selected ? "true" : "false");
        });

        shell.querySelectorAll("th.row-header[data-row]").forEach((header) => {
          const row = Number(header.dataset.row);
          header.classList.toggle(
            "tc-header-selected",
            Boolean(selection && row >= selection.r1 && row <= selection.r2)
          );
        });
        shell.querySelectorAll("th.column-header[data-col]").forEach((header) => {
          const column = Number(header.dataset.col);
          header.classList.toggle(
            "tc-header-selected",
            Boolean(selection && column >= selection.c1 && column <= selection.c2)
          );
        });
        selectionDirty = selectionSignature(selection) !== committedSelectionSignature;
        updateToolbar();
      }

      function selectionPayload(cleared = false) {
        if (!selection || cleared) {
          return {
            version: 1,
            sheet_name: String(args.sheet_name || ""),
            source_signature: String(args.source_signature || ""),
            input_key: String(args.active_input_key || ""),
            range_a1: "",
            qualified_range: "",
            zoom_percent: zoomPercent,
            cleared: true,
          };
        }
        const rangeA1 = rangeReference(selection);
        const rowCount = selection.r2 - selection.r1 + 1;
        const columnCount = selection.c2 - selection.c1 + 1;
        return {
          version: 1,
          sheet_name: String(args.sheet_name || ""),
          source_signature: String(args.source_signature || ""),
          input_key: String(args.active_input_key || ""),
          range_a1: rangeA1,
          qualified_range: `${quoteSheetName(args.sheet_name)}!${rangeA1}`,
          start_row: selection.r1,
          start_column: selection.c1,
          end_row: selection.r2,
          end_column: selection.c2,
          row_count: rowCount,
          column_count: columnCount,
          cell_count: rowCount * columnCount,
          anchor_row: anchor ? anchor.r1 : selection.r1,
          anchor_column: anchor ? anchor.c1 : selection.c1,
          active_row: activeCell ? activeCell.row : selection.r2,
          active_column: activeCell ? activeCell.column : selection.c2,
          zoom_percent: zoomPercent,
          visible_window: {
            start_row: Number(args.visible_start_row || 1),
            end_row: Number(args.visible_end_row || 1),
            start_column: Number(args.visible_start_column || 1),
            end_column: Number(args.visible_end_column || 1),
          },
          scroll_left: viewport ? Math.max(0, Math.round(viewport.scrollLeft)) : 0,
          scroll_top: viewport ? Math.max(0, Math.round(viewport.scrollTop)) : 0,
          cleared: false,
          selected_at: new Date().toISOString(),
        };
      }

      function commitSelection() {
        if (!selection || !selectionDirty) return;
        if (draftSaveTimer !== null) {
          window.clearTimeout(draftSaveTimer);
          draftSaveTimer = null;
        }
        const payload = selectionPayload(false);
        try {
          // Keep the just-assigned draft through Streamlit's first rerun. The
          // next render clears it once Python returns the same committed range.
          window.localStorage.setItem(draftStorageKey(), JSON.stringify(payload));
        } catch (_error) {
          // Storage can be disabled in hardened browser profiles.
        }
        committedSelectionSignature = selectionSignature(selection);
        selectionDirty = false;
        updateToolbar();
        rememberFocusForRerun();
        setComponentValue(payload);
      }

      function stopAutoScroll() {
        if (autoScrollTimer !== null) {
          window.clearInterval(autoScrollTimer);
          autoScrollTimer = null;
        }
      }

      function autoScrollStep() {
        if (!dragging || !viewport || !lastPointer) return;
        const rect = viewport.getBoundingClientRect();
        const edge = 34;
        let dx = 0;
        let dy = 0;
        if (lastPointer.x < rect.left + edge) dx = -16;
        else if (lastPointer.x > rect.right - edge) dx = 16;
        if (lastPointer.y < rect.top + edge) dy = -14;
        else if (lastPointer.y > rect.bottom - edge) dy = 14;
        if (dx || dy) {
          viewport.scrollBy(dx, dy);
          const hovered = document.elementFromPoint(lastPointer.x, lastPointer.y);
          const endpoint = selectionFromTarget(hovered, dragMode);
          if (endpoint && anchor) {
            selection = normalizedBounds(anchor, endpoint);
            activeCell = { row: endpoint.r2, column: endpoint.c2 };
            applyHighlight();
          }
        }
      }

      function focusViewport() {
        if (viewport) viewport.focus({ preventScroll: true });
      }

      function beginDrag(event) {
        if (event.button !== 0) return;
        const mode = targetMode(event.target);
        if (!mode) return;
        event.preventDefault();
        focusViewport();

        const endpoint = selectionFromTarget(event.target, mode);
        if (!endpoint) return;

        dragMode = mode;
        dragging = true;
        shell.classList.add("tc-dragging");
        lastPointer = { x: event.clientX, y: event.clientY };

        if (!(event.shiftKey && selection && anchor)) {
          anchor = endpoint;
        }
        selection = normalizedBounds(anchor, endpoint);
        activeCell = { row: endpoint.r2, column: endpoint.c2 };
        applyHighlight();
        scheduleRememberDraft();
        stopAutoScroll();
        autoScrollTimer = window.setInterval(autoScrollStep, 45);
      }

      function continueDrag(event) {
        if (!dragging) return;
        lastPointer = { x: event.clientX, y: event.clientY };
        const target = document.elementFromPoint(event.clientX, event.clientY);
        const endpoint = selectionFromTarget(target, dragMode);
        if (!endpoint || !anchor) return;
        selection = normalizedBounds(anchor, endpoint);
        activeCell = { row: endpoint.r2, column: endpoint.c2 };
        applyHighlight();
        scheduleRememberDraft();
      }

      function endDrag(event) {
        if (!dragging) return;
        continueDrag(event);
        dragging = false;
        dragMode = null;
        shell.classList.remove("tc-dragging");
        stopAutoScroll();
        applyHighlight();
        scheduleRememberDraft(0);
      }

      function clearSelection(emit = true) {
        selection = null;
        anchor = null;
        activeCell = null;
        if (emit) committedSelectionSignature = "";
        clearDraftSelection();
        applyHighlight();
        if (emit) {
          rememberFocusForRerun();
          setComponentValue(selectionPayload(true));
        }
      }

      async function copySelection() {
        if (!selection || !shell) return;
        const value = `${quoteSheetName(args.sheet_name)}!${rangeReference(selection)}`;
        let copied = false;
        try {
          await navigator.clipboard.writeText(value);
          copied = true;
        } catch (_error) {
          const textarea = document.createElement("textarea");
          textarea.value = value;
          textarea.style.position = "fixed";
          textarea.style.opacity = "0";
          document.body.appendChild(textarea);
          textarea.select();
          copied = document.execCommand("copy");
          textarea.remove();
        }
        const status = shell.querySelector("#tc-selection-status");
        if (copied && status) {
          status.innerHTML = `<span class="tc-selection-status-dot"></span><span class="tc-copy-feedback">Copied ${escapeHtml(value)}</span>`;
          window.setTimeout(updateToolbar, 1200);
        }
      }

      function ensureActiveCellVisible() {
        if (!viewport || !activeCell) return;
        const cell = cellForCoordinate(activeCell.row, activeCell.column);
        if (!cell) return;
        const viewRect = viewport.getBoundingClientRect();
        const cellRect = cell.getBoundingClientRect();
        const scale = zoomPercent / 100;
        const stickyTop = 34 * scale;
        const stickyLeft = 54 * scale;
        let dx = 0;
        let dy = 0;
        if (cellRect.left < viewRect.left + stickyLeft) {
          dx = cellRect.left - (viewRect.left + stickyLeft) - 4;
        } else if (cellRect.right > viewRect.right - 4) {
          dx = cellRect.right - viewRect.right + 8;
        }
        if (cellRect.top < viewRect.top + stickyTop) {
          dy = cellRect.top - (viewRect.top + stickyTop) - 4;
        } else if (cellRect.bottom > viewRect.bottom - 4) {
          dy = cellRect.bottom - viewRect.bottom + 8;
        }
        if (dx || dy) viewport.scrollBy(dx, dy);
      }


      function revealSelectionStart() {
        if (!viewport || !selection) return;
        const cell = cellForCoordinate(selection.r1, selection.c1);
        if (!cell) return;
        const viewRect = viewport.getBoundingClientRect();
        const cellRect = cell.getBoundingClientRect();
        const scale = zoomPercent / 100;
        const stickyTop = 34 * scale;
        const stickyLeft = 54 * scale;
        const targetLeft = viewRect.left + stickyLeft + 8;
        const targetTop = viewRect.top + stickyTop + 8;
        const dx = cellRect.left - targetLeft;
        const dy = cellRect.top - targetTop;
        if (Math.abs(dx) > 2 || Math.abs(dy) > 2) {
          viewport.scrollBy(dx, dy);
        }
      }

      function currentActiveBounds() {
        const limits = visibleLimits();
        if (activeCell) return boundsForCoordinate(activeCell.row, activeCell.column);
        if (selection) return boundsForCoordinate(selection.r2, selection.c2);
        return boundsForCoordinate(limits.startRow, limits.startColumn);
      }

      function moveSelectionTo(row, column, extend) {
        const limits = visibleLimits();
        const safeRow = clamp(row, limits.startRow, limits.endRow);
        const safeColumn = clamp(column, limits.startColumn, limits.endColumn);
        const target = boundsForCoordinate(safeRow, safeColumn);
        activeCell = { row: safeRow, column: safeColumn };
        if (!extend || !anchor) anchor = target;
        selection = normalizedBounds(anchor, target);
        applyHighlight();
        ensureActiveCellVisible();
        scheduleRememberDraft();
      }

      function pageRowDelta() {
        if (!viewport) return 20;
        const scale = zoomPercent / 100;
        return Math.max(1, Math.floor(viewport.clientHeight / Math.max(12, 22 * scale)) - 2);
      }

      function handleGridKeydown(event) {
        const key = String(event.key || "");
        const lowerKey = key.toLowerCase();
        const ctrlOrMeta = event.ctrlKey || event.metaKey;

        if (ctrlOrMeta && ["+", "=", "add"].includes(lowerKey)) {
          event.preventDefault();
          setZoom(zoomPercent + ZOOM_STEP);
          return;
        }
        if (ctrlOrMeta && ["-", "_", "subtract"].includes(lowerKey)) {
          event.preventDefault();
          setZoom(zoomPercent - ZOOM_STEP);
          return;
        }
        if (ctrlOrMeta && lowerKey === "0") {
          event.preventDefault();
          setZoom(100);
          return;
        }
        if (ctrlOrMeta && key === "Enter" && selection) {
          event.preventDefault();
          commitSelection();
          return;
        }
        if (key === "Escape") {
          event.preventDefault();
          clearSelection(true);
          return;
        }
        if (ctrlOrMeta && lowerKey === "c" && selection) {
          event.preventDefault();
          copySelection();
          return;
        }

        const limits = visibleLimits();
        if (ctrlOrMeta && lowerKey === "a") {
          event.preventDefault();
          anchor = boundsForCoordinate(limits.startRow, limits.startColumn);
          activeCell = { row: limits.endRow, column: limits.endColumn };
          selection = {
            r1: limits.startRow,
            c1: limits.startColumn,
            r2: limits.endRow,
            c2: limits.endColumn,
          };
          applyHighlight();
          scheduleRememberDraft();
          return;
        }

        const activeBounds = currentActiveBounds();
        let targetRow = activeCell ? activeCell.row : activeBounds.r2;
        let targetColumn = activeCell ? activeCell.column : activeBounds.c2;
        let handled = true;

        if (key === "ArrowLeft" || (key === "Tab" && event.shiftKey)) {
          targetColumn = ctrlOrMeta ? limits.startColumn : activeBounds.c1 - 1;
        } else if (key === "ArrowRight" || key === "Tab") {
          targetColumn = ctrlOrMeta ? limits.endColumn : activeBounds.c2 + 1;
        } else if (key === "ArrowUp" || (key === "Enter" && event.shiftKey)) {
          targetRow = ctrlOrMeta ? limits.startRow : activeBounds.r1 - 1;
        } else if (key === "ArrowDown" || key === "Enter") {
          targetRow = ctrlOrMeta ? limits.endRow : activeBounds.r2 + 1;
        } else if (key === "Home") {
          if (ctrlOrMeta) targetRow = limits.startRow;
          targetColumn = limits.startColumn;
        } else if (key === "End") {
          if (ctrlOrMeta) targetRow = limits.endRow;
          targetColumn = limits.endColumn;
        } else if (key === "PageUp") {
          targetRow = activeBounds.r1 - pageRowDelta();
        } else if (key === "PageDown") {
          targetRow = activeBounds.r2 + pageRowDelta();
        } else if (key === " " && event.shiftKey && !ctrlOrMeta) {
          event.preventDefault();
          const row = clamp(targetRow, limits.startRow, limits.endRow);
          anchor = { r1: row, c1: limits.startColumn, r2: row, c2: limits.startColumn };
          activeCell = { row, column: limits.endColumn };
          selection = { r1: row, c1: limits.startColumn, r2: row, c2: limits.endColumn };
          applyHighlight();
          scheduleRememberDraft();
          return;
        } else if (key === " " && ctrlOrMeta) {
          event.preventDefault();
          const column = clamp(targetColumn, limits.startColumn, limits.endColumn);
          anchor = { r1: limits.startRow, c1: column, r2: limits.startRow, c2: column };
          activeCell = { row: limits.endRow, column };
          selection = { r1: limits.startRow, c1: column, r2: limits.endRow, c2: column };
          applyHighlight();
          scheduleRememberDraft();
          return;
        } else {
          handled = false;
        }

        if (!handled) return;
        event.preventDefault();
        moveSelectionTo(targetRow, targetColumn, Boolean(event.shiftKey));
      }

      function handleZoomWheel(event) {
        if (!event.ctrlKey && !event.metaKey) return;
        event.preventDefault();
        setZoom(zoomPercent + (event.deltaY < 0 ? ZOOM_STEP : -ZOOM_STEP));
      }

      function validInitialSelection(value) {
        if (!value || value.cleared || !value.range_a1) return null;
        if (String(value.sheet_name || "") !== String(args.sheet_name || "")) return null;
        const numbers = [
          Number(value.start_row),
          Number(value.start_column),
          Number(value.end_row),
          Number(value.end_column),
        ];
        if (numbers.some((item) => !Number.isFinite(item) || item < 1)) return null;
        return {
          r1: Math.min(numbers[0], numbers[2]),
          c1: Math.min(numbers[1], numbers[3]),
          r2: Math.max(numbers[0], numbers[2]),
          c2: Math.max(numbers[1], numbers[3]),
        };
      }

      function restoreNavigationState(initialValue) {
        const limits = visibleLimits();
        if (!selection) {
          anchor = null;
          activeCell = null;
          return;
        }
        const anchorRow = clamp(
          Number(initialValue.anchor_row || selection.r1),
          limits.startRow,
          limits.endRow
        );
        const anchorColumn = clamp(
          Number(initialValue.anchor_column || selection.c1),
          limits.startColumn,
          limits.endColumn
        );
        const activeRow = clamp(
          Number(initialValue.active_row || selection.r2),
          limits.startRow,
          limits.endRow
        );
        const activeColumn = clamp(
          Number(initialValue.active_column || selection.c2),
          limits.startColumn,
          limits.endColumn
        );
        anchor = boundsForCoordinate(anchorRow, anchorColumn);
        activeCell = { row: activeRow, column: activeColumn };
      }

      function render() {
        stopAutoScroll();
        if (draftSaveTimer !== null) {
          window.clearTimeout(draftSaveTimer);
          draftSaveTimer = null;
        }
        const tableHtml = String(args.table_html || "");
        root.innerHTML = `
          <div class="tc-component-shell">
            <div class="tc-range-toolbar">
              <div id="tc-range-name" class="tc-range-name-box">No selection</div>
              <div class="tc-range-copy">
                <div id="tc-range-title" class="tc-range-title">Select a worksheet range</div>
                <div id="tc-range-help" class="tc-range-help">Click or drag · Arrow keys move · Shift+Arrow extends · Ctrl+Enter assigns</div>
              </div>
              <div class="tc-zoom-controls" aria-label="Worksheet zoom">
                <button id="tc-zoom-out" class="tc-zoom-button" type="button" title="Zoom out">−</button>
                <button id="tc-zoom-value" class="tc-zoom-value" type="button" title="Reset worksheet zoom to 100%">100%</button>
                <button id="tc-zoom-in" class="tc-zoom-button" type="button" title="Zoom in">+</button>
              </div>
              <div class="tc-range-actions">
                <button id="tc-assign-range" class="tc-range-button tc-assign-button" type="button" disabled>Assign range</button>
                <button id="tc-toggle-rows" class="tc-range-button tc-visibility-button" type="button" disabled>Hide rows</button>
                <button id="tc-copy-range" class="tc-range-button tc-copy-button" type="button" disabled>Copy reference</button>
                <button id="tc-clear-range" class="tc-range-button" type="button" disabled>Clear</button>
              </div>
            </div>
            <div id="tc-selection-status" class="tc-selection-status"></div>
            ${tableHtml}
          </div>
        `;

        shell = root.querySelector(".tc-component-shell");
        viewport = root.querySelector(".excel-viewport");
        loadRowVisibilityState();
        applyRowVisibilityState();
        if (viewport) {
          viewport.tabIndex = 0;
          viewport.setAttribute("role", "application");
          viewport.setAttribute(
            "aria-label",
            "Excel preview. Use arrow keys to move, Shift plus arrow keys to extend, and Control plus Enter to assign the range."
          );
        }
        rebuildCellLookup();

        const initialValue = args.initial_selection || {};
        const storedDraftValue = storedJson(draftStorageKey()) || {};
        const committedSelection = validInitialSelection(initialValue);
        const draftSelection = validInitialSelection(storedDraftValue);
        committedSelectionSignature = selectionSignature(committedSelection);

        let navigationValue = initialValue;
        const hasNamedCommittedSelection = Boolean(
          String(args.active_input_key || "").trim() && committedSelection
        );

        // A named input's assigned range is authoritative whenever that input
        // is reopened.  Draft selections are useful only before an input has a
        // committed mapping.  Allowing an old localStorage draft to override a
        // committed auto-detected/manual mapping caused the visible highlight
        // to jump to a stale cell or row even though the sidebar showed the
        // correct A1 range.
        if (hasNamedCommittedSelection) {
          selection = committedSelection;
          navigationValue = initialValue;
          clearDraftSelection();
        } else if (
          draftSelection &&
          selectionSignature(draftSelection) !== committedSelectionSignature
        ) {
          selection = draftSelection;
          navigationValue = storedDraftValue;
        } else {
          selection = committedSelection;
          if (
            draftSelection &&
            selectionSignature(draftSelection) === committedSelectionSignature
          ) {
            clearDraftSelection();
          }
        }

        restoreNavigationState(navigationValue);
        selectionDirty = selectionSignature(selection) !== committedSelectionSignature;
        const initialZoom = Number(navigationValue.zoom_percent || 0)
          || storedNumber(zoomStorageKey(), 100);
        zoomPercent = 0;
        setZoom(initialZoom, false);
        applyHighlight();

        shell.addEventListener("mousedown", beginDrag);
        shell.addEventListener("mousemove", continueDrag);
        window.onmouseup = endDrag;
        shell.querySelector("#tc-assign-range").addEventListener("click", commitSelection);
        shell.querySelector("#tc-toggle-rows").addEventListener("click", toggleSelectedRowsVisibility);
        shell.querySelector("#tc-clear-range").addEventListener("click", () => clearSelection(true));
        shell.querySelector("#tc-copy-range").addEventListener("click", copySelection);
        shell.querySelector("#tc-zoom-out").addEventListener("click", () => setZoom(zoomPercent - ZOOM_STEP));
        shell.querySelector("#tc-zoom-in").addEventListener("click", () => setZoom(zoomPercent + ZOOM_STEP));
        shell.querySelector("#tc-zoom-value").addEventListener("click", () => setZoom(100));
        if (viewport) {
          viewport.addEventListener("keydown", handleGridKeydown);
          viewport.addEventListener("wheel", handleZoomWheel, { passive: false });
        }

        const restoredScrollLeft = Math.max(0, Number(navigationValue.scroll_left || 0));
        const restoredScrollTop = Math.max(0, Number(navigationValue.scroll_top || 0));
        const restoreFocus = consumeFocusRestore();
        window.requestAnimationFrame(() => {
          if (viewport) {
            viewport.scrollLeft = restoredScrollLeft;
            viewport.scrollTop = restoredScrollTop;
            if (restoreFocus) viewport.focus({ preventScroll: true });
          }
          fitPreviewToViewport();
          applyHighlight();
          if (committedSelection) revealSelectionStart();
          else ensureActiveCellVisible();
          window.requestAnimationFrame(() => {
            fitPreviewToViewport();
            applyHighlight();
            if (committedSelection) revealSelectionStart();
          });
          window.setTimeout(() => {
            applyHighlight();
            if (committedSelection) revealSelectionStart();
          }, 90);
        });
      }

      window.addEventListener("resize", scheduleViewportFit);
      try {
        window.parent.addEventListener("resize", scheduleViewportFit);
      } catch (_error) {
        // The iframe resize event still covers sandboxed component runtimes.
      }

      window.addEventListener("message", (event) => {
        if (!event.data || event.data.type !== "streamlit:render") return;
        args = event.data.args || {};
        render();
      });

      sendMessage("streamlit:componentReady", { apiVersion: 1 });
      appliedFrameHeight = 680;
      setFrameHeight(appliedFrameHeight);
    })();
  </script>
</body>
</html>

"""


def _get_range_selector_component():
    """Create the dependency-free Streamlit component used by the preview grid.

    The frontend is embedded in this Python module so the UI folder still
    contains only the requested two Python scripts. At runtime it is written to
    the user's temporary directory and served by Streamlit as a normal custom
    component.
    """
    global _RANGE_SELECTOR_COMPONENT
    if _RANGE_SELECTOR_COMPONENT is not None:
        return _RANGE_SELECTOR_COMPONENT

    component_dir = (
        Path(tempfile.gettempdir())
        / "ThinkCellSlideMaker"
        / _RANGE_SELECTOR_COMPONENT_BUILD
    )
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        current = index_path.read_text(encoding="utf-8")
    except OSError:
        current = ""
    if current != _RANGE_SELECTOR_COMPONENT_HTML:
        temporary_path = index_path.with_suffix(".html.writing")
        temporary_path.write_text(_RANGE_SELECTOR_COMPONENT_HTML, encoding="utf-8")
        temporary_path.replace(index_path)

    _RANGE_SELECTOR_COMPONENT = components.declare_component(
        "thinkcell_excel_range_selector_v9",
        path=str(component_dir),
    )
    return _RANGE_SELECTOR_COMPONENT


# -----------------------------------------------------------------------------
# Workbook preview
# -----------------------------------------------------------------------------


def _workbook_cache_token(path: Path) -> int:
    try:
        return int(path.stat().st_mtime_ns)
    except OSError:
        return 0


@lru_cache(maxsize=3)
def _analysis_workbook(path_text: str, cache_token: int):
    """Open one streaming, style-aware workbook object per file version.

    The previous implementation used ``read_only=False`` so openpyxl fully
    materialized every worksheet before the preview could appear. A large
    client workbook could therefore look frozen for many minutes. Read-only
    mode still exposes cached values, fonts, fills, alignments, number formats,
    and theme information needed by this UI, while worksheet cells are streamed
    only for the ranges actually inspected.
    """
    del cache_token
    from openpyxl import load_workbook

    path = Path(path_text)
    with ui_stage(
        "workbook.open_streaming",
        warn_after=12,
        stack_after=35,
        path=str(path),
        size_bytes=(path.stat().st_size if path.is_file() else None),
    ):
        try:
            workbook = load_workbook(
                path_text,
                read_only=True,
                data_only=True,
                keep_links=False,
            )
        except TypeError:
            workbook = load_workbook(
                path_text,
                read_only=True,
                data_only=True,
            )
    ui_log(
        "WORKBOOK OPEN",
        path=str(path),
        sheets=len(workbook.sheetnames),
        sheet_names=list(workbook.sheetnames),
        read_only=True,
    )
    return workbook


def _stream_range_values(
    worksheet,
    *,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[list[Any]]:
    """Read one bounded worksheet rectangle in a single streaming pass."""
    min_row = max(1, int(min_row))
    max_row = max(min_row, int(max_row))
    min_col = max(1, int(min_col))
    max_col = max(min_col, int(max_col))
    return [
        list(row)
        for row in worksheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        )
    ]


def _stream_range_cells(
    worksheet,
    *,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
) -> list[list[Any]]:
    """Read one bounded worksheet rectangle with style-aware cell objects."""
    min_row = max(1, int(min_row))
    max_row = max(min_row, int(max_row))
    min_col = max(1, int(min_col))
    max_col = max(min_col, int(max_col))
    return [
        list(row)
        for row in worksheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
            values_only=False,
        )
    ]


def _normalize_sheet_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _preview_sheet_role(sheet_name: str) -> str | None:
    """Classify only the workbook tabs used by the Slide Maker."""
    raw = str(sheet_name or "").strip()
    key = _normalize_sheet_key(raw)
    words = set(re.findall(r"[a-z0-9]+", raw.lower()))

    if "taxonomy" in raw.casefold():
        return "taxonomy"

    if (
        key in {"summary", "aurasummary", "summaryaura"}
        or "aurasummary" in key
        or ("aura" in words and "summary" in words)
    ):
        return "aura"

    is_geocost = "geocost" in key or ("geo" in words and "cost" in words)
    if not is_geocost:
        return None

    if (
        "global" in words
        or "world" in words
        or "worldwide" in words
        or "international" in words
        or "global" in key
    ):
        return "geocost_global"
    if (
        "us" in words
        or "usa" in words
        or "domestic" in words
        or "unitedstates" in key
        or key.startswith("usgeocost")
        or key.startswith("usageocost")
    ):
        return "geocost_us"
    return "geocost"


def _preview_sheet_label(role: str, source_name: str) -> str:
    role_label = {
        "aura": "Aura Summary",
        "taxonomy": "Taxonomy",
        "geocost_global": "Global Geo Cost",
        "geocost_us": "US Geo Cost",
        "geocost": "Geo Cost",
    }.get(role, source_name)
    return role_label if role_label.lower() == source_name.strip().lower() else f"{role_label} — {source_name}"


def _related_ooxml_part(source_part: str, target: str) -> str:
    import posixpath

    if str(target).startswith("/"):
        return str(target).lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(source_part), str(target))
    )


def _sheet_dimension_from_archive(
    archive: zipfile.ZipFile,
    sheet_part: str,
) -> tuple[int, int]:
    from openpyxl.utils.cell import range_boundaries

    try:
        with archive.open(sheet_part) as stream:
            for _event, element in ET.iterparse(stream, events=("start",)):
                local_name = element.tag.rsplit("}", 1)[-1]
                if local_name == "dimension":
                    reference = str(element.attrib.get("ref") or "A1")
                    try:
                        _min_col, _min_row, max_col, max_row = range_boundaries(reference)
                        return max(1, int(max_row)), max(1, int(max_col))
                    except Exception:
                        return 1, 1
                if local_name == "sheetData":
                    break
    except Exception:
        pass
    return 1, 1


@st.cache_data(show_spinner=False, max_entries=24)
def worksheet_content_extent(
    path_text: str,
    cache_token: int,
    sheet_part: str,
) -> dict[str, int]:
    """Find the final meaningful table row without trusting Excel's used range.

    Worksheet dimensions often include hundreds or thousands of rows that are
    merely formatted. Conversely, a formula can be meaningful even when its
    cached value is blank. This scans worksheet XML for cells containing a
    value, inline/shared string, or formula, groups nearby populated rows into
    table-like blocks, and chooses the final qualifying block. A few isolated
    cells far below a real table therefore do not make the default preview
    enormous.
    """
    del cache_token
    from openpyxl.utils.cell import range_boundaries

    row_cell_counts: dict[int, int] = {}
    merge_ranges: list[tuple[int, int, int, int]] = []
    first_content_row = EXCEL_MAX_ROWS
    first_content_column = EXCEL_MAX_COLUMNS
    last_content_row = 1
    last_content_column = 1
    populated_cell_count = 0
    coordinate_pattern = re.compile(r"^([A-Z]+)([0-9]+)$", flags=re.I)

    def column_number(letters: str) -> int:
        result = 0
        for character in letters.upper():
            result = result * 26 + (ord(character) - 64)
        return result

    with zipfile.ZipFile(path_text, "r") as archive:
        with archive.open(sheet_part) as stream:
            for event, element in ET.iterparse(stream, events=("end",)):
                local_name = element.tag.rsplit("}", 1)[-1]
                if local_name == "c":
                    reference = str(element.attrib.get("r") or "").strip()
                    match = coordinate_pattern.fullmatch(reference)
                    if match is not None:
                        has_formula = False
                        has_payload = False
                        for child in element.iter():
                            child_name = child.tag.rsplit("}", 1)[-1]
                            if child_name == "f":
                                has_formula = True
                            elif child_name in {"v", "t"}:
                                if child.text is not None and str(child.text) != "":
                                    has_payload = True
                        if has_formula or has_payload:
                            column = column_number(match.group(1))
                            row = int(match.group(2))
                            if row > PREVIEW_CONTENT_SCAN_MAX_ROW:
                                # Aura/Taxonomy/Geo Cost source structures are
                                # bounded well above their expected table rows.
                                # Stop before formula-filled million-row tails
                                # can block the Streamlit upload indefinitely.
                                element.clear()
                                break
                            if (
                                1 <= row <= PREVIEW_CONTENT_SCAN_MAX_ROW
                                and 1 <= column <= PREVIEW_CONTENT_SCAN_MAX_COLUMN
                            ):
                                populated_cell_count += 1
                                row_cell_counts[row] = row_cell_counts.get(row, 0) + 1
                                first_content_row = min(first_content_row, row)
                                first_content_column = min(first_content_column, column)
                                last_content_row = max(last_content_row, row)
                                last_content_column = max(last_content_column, column)
                    element.clear()
                elif local_name == "mergeCell":
                    reference = str(element.attrib.get("ref") or "").strip()
                    if reference:
                        try:
                            merge_ranges.append(
                                tuple(
                                    int(value)
                                    for value in range_boundaries(reference)
                                )
                            )
                        except Exception:
                            pass
                    element.clear()
                elif local_name in {"row", "mergeCells", "sheetData"}:
                    element.clear()

    content_rows = sorted(row_cell_counts)
    if not content_rows:
        return {
            "first_content_row": 1,
            "first_content_column": 1,
            "last_content_row": 1,
            "last_content_column": 1,
            "last_table_row": 1,
            "preview_end_row": 1 + PREVIEW_ROWS_AFTER_LAST_TABLE,
            "populated_cell_count": 0,
            "table_block_count": 0,
        }

    # A table may contain a few intentionally blank spacer rows. Rows separated
    # by up to three blanks remain part of one visual block.
    blocks: list[list[int]] = []
    current_block: list[int] = [content_rows[0]]
    for row in content_rows[1:]:
        if row - current_block[-1] <= 4:
            current_block.append(row)
        else:
            blocks.append(current_block)
            current_block = [row]
    blocks.append(current_block)

    def table_like(block: list[int]) -> bool:
        total_cells = sum(row_cell_counts[row] for row in block)
        dense_rows = sum(1 for row in block if row_cell_counts[row] >= 2)
        return total_cells >= 4 and (dense_rows >= 1 or len(block) >= 4)

    table_blocks = [block for block in blocks if table_like(block)]
    selected_block = table_blocks[-1] if table_blocks else blocks[-1]
    last_table_row = selected_block[-1]
    selected_start_row = selected_block[0]

    # Preserve the full height of merged cells whose anchors belong to the
    # selected final table block.
    for min_col, min_row, max_col, max_row in merge_ranges:
        if min_row <= last_table_row and max_row >= selected_start_row:
            if min_row <= last_table_row:
                last_table_row = max(last_table_row, min(max_row, EXCEL_MAX_ROWS))
                last_content_column = max(
                    last_content_column,
                    min(max_col, EXCEL_MAX_COLUMNS),
                )

    preview_end_row = min(
        EXCEL_MAX_ROWS,
        max(1, last_table_row) + PREVIEW_ROWS_AFTER_LAST_TABLE,
    )
    return {
        "first_content_row": min(first_content_row, EXCEL_MAX_ROWS),
        "first_content_column": min(first_content_column, EXCEL_MAX_COLUMNS),
        "last_content_row": min(last_content_row, EXCEL_MAX_ROWS),
        "last_content_column": min(last_content_column, EXCEL_MAX_COLUMNS),
        "last_table_row": min(last_table_row, EXCEL_MAX_ROWS),
        "preview_end_row": preview_end_row,
        "populated_cell_count": populated_cell_count,
        "table_block_count": len(table_blocks),
    }


@st.cache_data(show_spinner=False, max_entries=12)
def workbook_sheet_inventory(
    path_text: str,
    cache_token: int,
) -> list[dict[str, Any]]:
    """Return every workbook tab plus lightweight name-based role hints."""
    del cache_token
    workbook_part = "xl/workbook.xml"
    rels_part = "xl/_rels/workbook.xml.rels"
    workbook_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"

    with zipfile.ZipFile(path_text, "r") as archive:
        rel_root = ET.fromstring(archive.read(rels_part))
        target_by_id: dict[str, str] = {}
        for relationship in rel_root.findall(f"{{{package_rel_ns}}}Relationship"):
            rel_id = relationship.attrib.get("Id")
            target = relationship.attrib.get("Target")
            if rel_id and target:
                target_by_id[rel_id] = _related_ooxml_part(workbook_part, target)

        workbook_root = ET.fromstring(archive.read(workbook_part))
        result: list[dict[str, Any]] = []
        for order, sheet in enumerate(
            workbook_root.findall(f".//{{{workbook_ns}}}sheet"),
            start=1,
        ):
            name = str(sheet.attrib.get("name") or "").strip()
            rel_id = sheet.attrib.get(f"{{{office_rel_ns}}}id")
            sheet_part = target_by_id.get(str(rel_id or ""), "")
            if not name or not sheet_part or sheet_part not in archive.namelist():
                continue
            role = _preview_sheet_role(name)
            rows, columns = _sheet_dimension_from_archive(archive, sheet_part)
            result.append(
                {
                    "name": name,
                    "role": role,
                    "display_label": (
                        _preview_sheet_label(role, name) if role else name
                    ),
                    "rows": rows,
                    "columns": columns,
                    "state": str(sheet.attrib.get("state") or "visible"),
                    "sheet_part": sheet_part,
                    "workbook_order": order,
                }
            )

    result.sort(key=lambda item: int(item.get("workbook_order", 9999)))
    return result


@st.cache_data(show_spinner=False, max_entries=12)
def workbook_manifest(path_text: str, cache_token: int) -> list[dict[str, Any]]:
    """Return Aura Summary, Taxonomy, and Geo Cost tabs for the workbook preview."""
    result = [
        dict(item)
        for item in workbook_sheet_inventory(path_text, cache_token)
        if item.get("role") is not None
    ]
    role_order = {
        "aura": 0,
        "taxonomy": 1,
        "geocost_global": 2,
        "geocost_us": 3,
        "geocost": 4,
    }
    result.sort(
        key=lambda item: (
            role_order.get(str(item.get("role")), 99),
            int(item.get("workbook_order", 9999)),
        )
    )
    return result


_TAXONOMY_L0_HEADER_KEYS = {
    "l0",
    "level0",
    "levelzero",
    "l0name",
    "l0category",
    "level0name",
}
_TAXONOMY_L1_HEADER_KEYS = {
    "l1",
    "level1",
    "levelone",
    "l1name",
    "l1function",
    "level1name",
    "function",
}


def _taxonomy_header_descriptor(value: Any) -> tuple[int, str, str] | None:
    """Return ``(level, normalized_prefix, display_prefix)`` for a taxonomy header.

    Supported pairs include L0/L1, Bain L0/Bain L1, VC L0/VC L1, and the
    equivalent Level 0/Level 1 forms.  Matching prefixes ensure that a workbook
    containing several taxonomy variants is interpreted as paired columns.
    """
    raw = re.sub(r"\s+", " ", str(value or "").strip())
    compact = _normalize_sheet_key(raw)
    if not compact:
        return None

    direct_l0 = {
        "l0", "level0", "levelzero", "l0name", "l0category", "level0name"
    }
    direct_l1 = {
        "l1", "level1", "levelone", "l1name", "l1function", "level1name",
        "function",
    }
    if compact in direct_l0:
        return 0, "", ""
    if compact in direct_l1:
        return 1, "", ""

    suffixes: tuple[tuple[str, int], ...] = (
        ("levelzero", 0),
        ("levelone", 1),
        ("level0", 0),
        ("level1", 1),
        ("l0", 0),
        ("l1", 1),
    )
    for suffix, level in suffixes:
        if compact.endswith(suffix) and len(compact) > len(suffix):
            prefix = compact[: -len(suffix)]
            display = re.sub(
                r"(?i)(?:\s|[-_/])*?(?:level\s*(?:zero|one|0|1)|l\s*[01])\s*$",
                "",
                raw,
            ).strip(" -_/|")
            return level, prefix, display
    return None


def _taxonomy_sheet_sort_key(sheet_name: str, workbook_order: int) -> tuple[int, int, int]:
    raw = str(sheet_name or "").strip()
    lowered = raw.casefold()
    compact = _normalize_sheet_key(raw)
    if compact == "taxonomy":
        tier = 0
    elif lowered.startswith("taxonomy") or lowered.endswith("taxonomy"):
        tier = 1
    else:
        tier = 2
    return tier, len(raw), workbook_order


def _match_sheet_name(sheet_names: list[str], requested: str) -> str | None:
    requested_text = str(requested or "").strip()
    if not requested_text:
        return None
    for sheet_name in sheet_names:
        if sheet_name == requested_text or sheet_name.casefold() == requested_text.casefold():
            return sheet_name
    requested_key = _normalize_sheet_key(requested_text)
    matches = [name for name in sheet_names if _normalize_sheet_key(name) == requested_key]
    return matches[0] if len(matches) == 1 else None


def _find_taxonomy_header_pair(
    worksheet,
    *,
    min_row: int,
    max_row: int,
    min_column: int,
    max_column: int,
) -> dict[str, Any] | None:
    """Find one adjacent matching L0/L1 header pair with one streamed scan."""
    scan_start = max(1, int(min_row))
    header_scan_end = max(scan_start, int(max_row))
    min_column = max(1, int(min_column))
    max_column = max(min_column, int(max_column))
    worksheet_max_row = max(scan_start, int(worksheet.max_row or scan_start))
    data_scan_end = min(worksheet_max_row, header_scan_end + 120)

    with ui_stage(
        "workbook.taxonomy_header_scan",
        warn_after=8,
        stack_after=25,
        sheet=str(getattr(worksheet, "title", "")),
        rows=f"{scan_start}:{data_scan_end}",
        columns=f"{min_column}:{max_column}",
    ):
        streamed = [
            list(row)
            for row in worksheet.iter_rows(
                min_row=scan_start,
                max_row=data_scan_end,
                min_col=min_column,
                max_col=max_column,
                values_only=True,
            )
        ]

    def value_at(row_number: int, column_number: int) -> Any:
        row_index = row_number - scan_start
        column_index = column_number - min_column
        if row_index < 0 or row_index >= len(streamed):
            return None
        row_values = streamed[row_index]
        if column_index < 0 or column_index >= len(row_values):
            return None
        return row_values[column_index]

    candidates: list[dict[str, Any]] = []
    for row in range(scan_start, header_scan_end + 1):
        descriptors: list[tuple[int, int, str, str, str]] = []
        for column in range(min_column, max_column + 1):
            raw = str(value_at(row, column) or "").strip()
            descriptor = _taxonomy_header_descriptor(raw)
            if descriptor is None:
                continue
            level, prefix, display_prefix = descriptor
            descriptors.append((column, level, prefix, display_prefix, raw))

        for l0_column, l0_level, l0_prefix, l0_display, l0_raw in descriptors:
            if l0_level != 0:
                continue
            for l1_column, l1_level, l1_prefix, l1_display, l1_raw in descriptors:
                if l1_level != 1 or l0_column == l1_column:
                    continue
                if abs(l1_column - l0_column) != 1:
                    continue
                if l0_prefix != l1_prefix:
                    continue

                sample_nonblank = 0
                sample_l1 = 0
                sample_end = min(data_scan_end, row + 120)
                for sample_row in range(row + 1, sample_end + 1):
                    left = str(value_at(sample_row, l0_column) or "").strip()
                    right = str(value_at(sample_row, l1_column) or "").strip()
                    if left or right:
                        sample_nonblank += 1
                    if right:
                        sample_l1 += 1

                score = 0
                score += 1000 if not l0_prefix else 820
                score += 150 if l1_column == l0_column + 1 else 120
                score += min(sample_l1, 50) * 8
                score += min(sample_nonblank, 50) * 2
                score += max(0, 80 - row)
                candidates.append(
                    {
                        "score": score,
                        "header_row": row,
                        "l0_column": l0_column,
                        "l1_column": l1_column,
                        "header_prefix": l0_display or l1_display,
                        "l0_header": l0_raw,
                        "l1_header": l1_raw,
                    }
                )

    if not candidates:
        return None
    candidates.sort(
        key=lambda item: (
            int(item["score"]),
            -int(item["header_row"]),
            -min(int(item["l0_column"]), int(item["l1_column"])),
        ),
        reverse=True,
    )
    return candidates[0]


def _taxonomy_data_extent(
    worksheet,
    *,
    header_row: int,
    l0_column: int,
    l1_column: int,
    max_row: int | None = None,
) -> tuple[int, int]:
    """Measure Taxonomy extent by streaming its two columns once."""
    row_count = 0
    last_data_row = header_row
    blank_streak = 0
    scan_end = min(
        int(worksheet.max_row or header_row),
        int(max_row) if max_row is not None else header_row + 5000,
    )
    if scan_end <= header_row:
        return row_count, last_data_row

    min_col = min(int(l0_column), int(l1_column))
    max_col = max(int(l0_column), int(l1_column))
    l0_offset = int(l0_column) - min_col
    l1_offset = int(l1_column) - min_col

    with ui_stage(
        "workbook.taxonomy_extent_scan",
        warn_after=8,
        stack_after=25,
        sheet=str(getattr(worksheet, "title", "")),
        rows=f"{header_row + 1}:{scan_end}",
        columns=f"{min_col}:{max_col}",
    ):
        rows = worksheet.iter_rows(
            min_row=header_row + 1,
            max_row=scan_end,
            min_col=min_col,
            max_col=max_col,
            values_only=True,
        )
        for row_number, values in enumerate(rows, start=header_row + 1):
            values = list(values)
            l0_value = str(values[l0_offset] or "").strip() if l0_offset < len(values) else ""
            l1_value = str(values[l1_offset] or "").strip() if l1_offset < len(values) else ""
            if not l0_value and not l1_value:
                blank_streak += 1
                if row_count and blank_streak >= 25:
                    break
                continue
            blank_streak = 0
            if (
                _taxonomy_header_descriptor(l0_value)
                and _taxonomy_header_descriptor(l1_value)
            ):
                continue
            if l1_value and not l1_value.startswith(">>"):
                row_count += 1
                last_data_row = row_number
            elif l0_value:
                last_data_row = row_number
    return row_count, last_data_row


def _taxonomy_profile_for_range(
    workbook,
    *,
    sheet_name: str,
    range_a1: str,
) -> dict[str, Any]:
    from openpyxl.utils.cell import range_boundaries

    profile: dict[str, Any] = {
        "sheet_name": sheet_name,
        "range_a1": range_a1,
        "header_row": None,
        "l0_column": None,
        "l1_column": None,
        "header_prefix": "",
        "l0_header": "",
        "l1_header": "",
        "row_count": 0,
        "error": "",
    }
    if sheet_name not in workbook.sheetnames:
        profile["error"] = f"Taxonomy worksheet {sheet_name!r} was not found."
        return profile
    try:
        min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    except Exception:
        profile["error"] = f"Taxonomy range {range_a1!r} is not a valid Excel range."
        return profile
    if max_col - min_col + 1 != 2:
        profile["error"] = (
            "Taxonomy must be selected as exactly two columns, including the L0/L1 header row."
        )
        return profile
    if max_row - min_row + 1 < 2:
        profile["error"] = "Taxonomy must include a header row and at least one data row."
        return profile

    worksheet = workbook[sheet_name]
    pair = _find_taxonomy_header_pair(
        worksheet,
        min_row=min_row,
        max_row=min(max_row, min_row + 30),
        min_column=min_col,
        max_column=max_col,
    )
    if pair is None:
        profile["error"] = (
            "The selected Taxonomy range must contain a matching two-column header pair such as "
            "L0/L1, Bain L0/Bain L1, or VC L0/VC L1."
        )
        return profile

    row_count, last_data_row = _taxonomy_data_extent(
        worksheet,
        header_row=int(pair["header_row"]),
        l0_column=int(pair["l0_column"]),
        l1_column=int(pair["l1_column"]),
        max_row=max_row,
    )
    profile.update(pair)
    profile["row_count"] = row_count
    profile["last_data_row"] = last_data_row
    if row_count < 1:
        profile["error"] = (
            f"Taxonomy range {sheet_name}!{range_a1} contains no L1 rows beneath its header pair."
        )
    return profile


@st.cache_data(show_spinner=False, max_entries=32)
def taxonomy_range_profile(
    path_text: str,
    cache_token: int,
    sheet_name: str,
    range_a1: str,
) -> dict[str, Any]:
    """Validate one explicit two-column Taxonomy range."""
    try:
        workbook = _analysis_workbook(path_text, cache_token)
    except Exception as exc:
        return {
            "sheet_name": sheet_name,
            "range_a1": range_a1,
            "row_count": 0,
            "error": f"Taxonomy worksheet could not be opened: {exc}",
        }
    return _taxonomy_profile_for_range(
        workbook,
        sheet_name=sheet_name,
        range_a1=range_a1,
    )


@st.cache_data(show_spinner=False, max_entries=24)
def taxonomy_sheet_profile(
    path_text: str,
    cache_token: int,
    requested_sheet: str = "",
) -> dict[str, Any]:
    """Resolve a Taxonomy tab and auto-detect one paired L0/L1 table."""
    from openpyxl.utils import get_column_letter

    profile: dict[str, Any] = {
        "sheet_name": "",
        "range_a1": "",
        "header_row": None,
        "l0_column": None,
        "l1_column": None,
        "header_prefix": "",
        "l0_header": "",
        "l1_header": "",
        "row_count": 0,
        "candidate_count": 0,
        "error": "",
    }
    try:
        workbook = _analysis_workbook(path_text, cache_token)
    except Exception as exc:
        profile["error"] = f"Taxonomy worksheet could not be opened: {exc}"
        return profile

    sheet_names = list(workbook.sheetnames)
    requested = str(requested_sheet or "").strip()
    if requested:
        selected = _match_sheet_name(sheet_names, requested)
        if selected is None:
            profile["error"] = (
                f"Taxonomy worksheet {requested!r} was not found in this workbook."
            )
            return profile
        candidates = [selected]
    else:
        indexed = [
            (index, name)
            for index, name in enumerate(sheet_names)
            if "taxonomy" in str(name).casefold()
        ]
        indexed.sort(key=lambda item: _taxonomy_sheet_sort_key(item[1], item[0]))
        candidates = [name for _, name in indexed]
        if not candidates:
            profile["error"] = (
                "No Taxonomy tab was detected. The worksheet name must contain "
                "'Taxonomy' (case-insensitive), or select it manually."
            )
            return profile

    profile["candidate_count"] = len(candidates)
    errors: list[str] = []
    for selected in candidates:
        worksheet = workbook[selected]
        max_row = max(1, min(int(worksheet.max_row or 1), 120))
        max_column = max(1, min(int(worksheet.max_column or 1), 120))
        pair = _find_taxonomy_header_pair(
            worksheet,
            min_row=1,
            max_row=max_row,
            min_column=1,
            max_column=max_column,
        )
        if pair is None:
            errors.append(f"{selected!r} has no adjacent matching L0/L1 header pair.")
            continue
        row_count, last_data_row = _taxonomy_data_extent(
            worksheet,
            header_row=int(pair["header_row"]),
            l0_column=int(pair["l0_column"]),
            l1_column=int(pair["l1_column"]),
        )
        min_col = min(int(pair["l0_column"]), int(pair["l1_column"]))
        max_col = max(int(pair["l0_column"]), int(pair["l1_column"]))
        range_a1 = (
            f"{get_column_letter(min_col)}{int(pair['header_row'])}:"
            f"{get_column_letter(max_col)}{max(last_data_row, int(pair['header_row']) + 1)}"
        )
        profile.update(pair)
        profile.update(
            {
                "sheet_name": selected,
                "range_a1": range_a1,
                "row_count": row_count,
                "last_data_row": last_data_row,
            }
        )
        if row_count < 1:
            profile["error"] = (
                f"Taxonomy worksheet {selected!r} contains no L1 rows beneath "
                "its detected header pair."
            )
        return profile

    profile["sheet_name"] = candidates[0] if candidates else ""
    profile["error"] = (
        "Could not auto-detect a two-column Taxonomy table. Supported header "
        "pairs include L0/L1, Bain L0/Bain L1, and VC L0/VC L1. "
        + " ".join(errors)
    )
    return profile


# -----------------------------------------------------------------------------
# Aura L1 Taxonomy-to-source-row lookup
# -----------------------------------------------------------------------------


def _l1_function_lookup_key(value: Any) -> str:
    """Normalize an L1 name for safe Taxonomy/source-row matching."""
    raw = "" if value is None else str(value).strip().casefold()
    tokens = re.findall(r"[a-z0-9]+", raw)
    aliases = {
        "mgmt": "management",
        "mgt": "management",
    }
    normalized = [aliases.get(token, token) for token in tokens]
    # Punctuation and conjunction variants such as ``Digital / Innovation``,
    # ``Digital & Innovation``, and ``Digital and Innovation`` should resolve
    # to the same source row.
    normalized = [token for token in normalized if token not in {"and", "the"}]
    return "".join(normalized)


def _l1_lookup_taxonomy_records(
    workbook,
    *,
    sheet_name: str,
    range_a1: str,
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    from openpyxl.utils.cell import range_boundaries

    profile = _taxonomy_profile_for_range(
        workbook,
        sheet_name=sheet_name,
        range_a1=range_a1,
    )
    if profile.get("error"):
        return [], profile

    _min_col, _min_row, _max_col, max_row = range_boundaries(range_a1)
    worksheet = workbook[sheet_name]
    header_row = int(profile["header_row"])
    l0_column = int(profile["l0_column"])
    l1_column = int(profile["l1_column"])
    min_col = min(l0_column, l1_column)
    max_col = max(l0_column, l1_column)
    values = _stream_range_values(
        worksheet,
        min_row=header_row + 1,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    )
    l0_offset = l0_column - min_col
    l1_offset = l1_column - min_col

    records: list[dict[str, Any]] = []
    last_l0 = ""
    for row_number, row_values in enumerate(values, start=header_row + 1):
        raw_l0 = str(row_values[l0_offset] or "").strip()
        function_name = str(row_values[l1_offset] or "").strip()
        if not raw_l0 and not function_name:
            continue
        l0_descriptor = _taxonomy_header_descriptor(raw_l0)
        l1_descriptor = _taxonomy_header_descriptor(function_name)
        if (
            l0_descriptor is not None
            and l1_descriptor is not None
            and l0_descriptor[0] == 0
            and l1_descriptor[0] == 1
        ):
            continue
        if raw_l0:
            last_l0 = raw_l0
        if not function_name or function_name.startswith(">>"):
            continue
        records.append(
            {
                "taxonomy_row": row_number,
                "l0": last_l0,
                "function": function_name,
                "key": _l1_function_lookup_key(function_name),
            }
        )
    return records, profile


def _l1_explicit_function_labels(
    workbook,
    *,
    fte_sheet: str,
    fte_min_row: int,
    fte_max_row: int,
    label_sheet: str,
    label_range: str,
) -> tuple[list[str], str, int] | None:
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.utils import get_column_letter

    if not label_range:
        return None
    resolved_sheet = label_sheet or fte_sheet
    if resolved_sheet != fte_sheet or resolved_sheet not in workbook.sheetnames:
        return None
    try:
        min_col, min_row, max_col, max_row = range_boundaries(label_range)
    except Exception:
        return None
    if min_col != max_col:
        return None
    if min_row != fte_min_row or max_row != fte_max_row:
        return None
    worksheet = workbook[resolved_sheet]
    matrix = _stream_range_values(
        worksheet,
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=min_col,
    )
    labels = [str((row[0] if row else None) or "").strip() for row in matrix]
    normalized_range = (
        f"{get_column_letter(min_col)}{min_row}:"
        f"{get_column_letter(min_col)}{max_row}"
    )
    return labels, normalized_range, min_col


def _l1_detect_function_label_column_in_bounds(
    workbook,
    *,
    fte_sheet: str,
    fte_min_row: int,
    fte_max_row: int,
    candidate_start_column: int,
    candidate_end_column: int,
    taxonomy_records: list[dict[str, Any]],
) -> tuple[list[str], str, int, dict[str, Any]]:
    """Find the source L1/function column inside a bounded set of columns."""
    from openpyxl.utils import get_column_letter

    worksheet = workbook[fte_sheet]
    taxonomy_keys = {
        str(record.get("key") or "")
        for record in taxonomy_records
        if record.get("key")
    }
    header_start_row = max(1, fte_min_row - 12)
    matrix = _stream_range_values(
        worksheet,
        min_row=header_start_row,
        max_row=fte_max_row,
        min_col=candidate_start_column,
        max_col=candidate_end_column,
    )
    data_start_offset = fte_min_row - header_start_row

    candidates: list[
        tuple[tuple[int, int, int, int, int], int, list[str], dict[str, Any]]
    ] = []
    for column in range(candidate_start_column, candidate_end_column + 1):
        column_offset = column - candidate_start_column
        labels = [
            str((row[column_offset] if column_offset < len(row) else None) or "").strip()
            for row in matrix[data_start_offset:]
        ]
        keys = [_l1_function_lookup_key(value) for value in labels]
        nonblank_keys = [key for key in keys if key]
        if not nonblank_keys:
            continue
        match_count = sum(1 for key in keys if key and key in taxonomy_keys)
        unique_match_count = len({key for key in keys if key in taxonomy_keys})
        text_count = sum(
            1 for value in labels if value and re.search(r"[A-Za-z]", value)
        )
        numeric_count = sum(
            1
            for value in labels
            if value
            and not re.search(r"[A-Za-z]", value)
            and re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", value.replace(",", ""))
        )
        header_hint = 0
        for row_values in matrix[:data_start_offset]:
            raw_header = str(
                (row_values[column_offset] if column_offset < len(row_values) else None)
                or ""
            ).strip()
            descriptor = _taxonomy_header_descriptor(raw_header)
            compact = re.sub(r"[^a-z0-9]+", "", raw_header.casefold())
            if descriptor is not None and descriptor[0] == 1:
                header_hint = max(header_hint, 3)
            elif "function" in compact or compact in {"l1", "level1", "levelone"}:
                header_hint = max(header_hint, 2)

        distance_from_values = candidate_end_column - column
        score = (
            match_count,
            unique_match_count,
            header_hint,
            text_count - numeric_count * 2,
            -distance_from_values,
        )
        candidates.append(
            (
                score,
                column,
                labels,
                {
                    "match_count": match_count,
                    "unique_match_count": unique_match_count,
                    "header_hint": header_hint,
                    "text_count": text_count,
                    "numeric_count": numeric_count,
                    "distance_from_values": distance_from_values,
                },
            )
        )

    if not candidates:
        raise ValueError(
            "The selected Function + FTE / $B table contains no usable text "
            "column before the company-value columns. Include the L1/function "
            "labels on the left of the numeric matrix."
        )
    candidates.sort(key=lambda item: item[0], reverse=True)
    score, column, labels, details = candidates[0]
    if int(details.get("match_count") or 0) < 1:
        raise ValueError(
            "No function/L1 column inside the selected Function + FTE / $B "
            "table matches the selected Taxonomy L1 values."
        )
    label_range = (
        f"{get_column_letter(column)}{fte_min_row}:"
        f"{get_column_letter(column)}{fte_max_row}"
    )
    details["score"] = list(score)
    details["candidate_start_column"] = candidate_start_column
    details["candidate_end_column"] = candidate_end_column
    return labels, label_range, column, details


def _l1_detect_function_label_column(
    workbook,
    *,
    fte_sheet: str,
    fte_min_row: int,
    fte_max_row: int,
    fte_min_col: int,
    taxonomy_records: list[dict[str, Any]],
) -> tuple[list[str], str, int, dict[str, Any]]:
    """Legacy lookup for schema-4 numeric-only FTE mappings."""
    scan_start = max(1, fte_min_col - 40)
    if scan_start >= fte_min_col:
        raise ValueError(
            "Could not find an L1/function-name column to the left of the selected "
            "FTE / $B matrix."
        )
    return _l1_detect_function_label_column_in_bounds(
        workbook,
        fte_sheet=fte_sheet,
        fte_min_row=fte_min_row,
        fte_max_row=fte_max_row,
        candidate_start_column=scan_start,
        candidate_end_column=fte_min_col - 1,
        taxonomy_records=taxonomy_records,
    )


@st.cache_data(show_spinner=False, max_entries=32)
def l1_function_lookup_profile(
    path_text: str,
    cache_token: int,
    taxonomy_sheet: str,
    taxonomy_range: str,
    fte_sheet: str,
    fte_range: str,
    explicit_label_sheet: str = "",
    explicit_label_range: str = "",
    company_count: int = 0,
) -> dict[str, Any]:
    """Match Taxonomy L1 rows to labels embedded in the selected FTE table."""
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    result: dict[str, Any] = {
        "attempted": bool(taxonomy_sheet and taxonomy_range and fte_sheet and fte_range),
        "ready": False,
        "taxonomy_sheet": taxonomy_sheet,
        "taxonomy_range": taxonomy_range,
        "fte_sheet": fte_sheet,
        "fte_range": fte_range,
        "fte_header_row": None,
        "fte_value_range": "",
        "function_label_sheet": "",
        "function_label_range": "",
        "function_label_column": None,
        "descriptor_column_count": 0,
        "value_column_count": 0,
        "taxonomy_count": 0,
        "source_row_count": 0,
        "matched_count": 0,
        "matches": [],
        "missing": [],
        "ambiguous": [],
        "extra_source": [],
        "error": "",
    }
    if not result["attempted"]:
        return result

    try:
        workbook = _analysis_workbook(path_text, cache_token)
    except Exception as exc:
        result["error"] = f"Workbook could not be opened for L1 function lookup: {exc}"
        return result

    try:
        if taxonomy_sheet not in workbook.sheetnames:
            result["error"] = f"Taxonomy worksheet {taxonomy_sheet!r} was not found."
            return result
        if fte_sheet not in workbook.sheetnames:
            result["error"] = f"FTE worksheet {fte_sheet!r} was not found."
            return result

        taxonomy_records, taxonomy_profile = _l1_lookup_taxonomy_records(
            workbook,
            sheet_name=taxonomy_sheet,
            range_a1=taxonomy_range,
        )
        if taxonomy_profile.get("error"):
            result["error"] = str(taxonomy_profile["error"])
            return result
        result["taxonomy_count"] = len(taxonomy_records)
        if not taxonomy_records:
            result["error"] = "The selected Taxonomy range contains no L1 functions."
            return result

        try:
            fte_min_col, fte_min_row, fte_max_col, fte_max_row = range_boundaries(
                fte_range
            )
        except Exception:
            result["error"] = (
                f"Function + FTE / $B range {fte_range!r} is not valid."
            )
            return result
        fte_header_row = fte_min_row
        fte_data_min_row = fte_header_row + 1
        if fte_data_min_row > fte_max_row:
            result["error"] = (
                "Function + FTE / $B table must include its header row and at "
                "least one data row beneath it."
            )
            return result
        result["fte_header_row"] = fte_header_row
        result["source_row_count"] = fte_max_row - fte_data_min_row + 1

        selected_column_count = fte_max_col - fte_min_col + 1
        resolved_company_count = max(0, int(company_count or 0))
        detection_details: dict[str, Any] = {}

        if resolved_company_count > 0:
            descriptor_count = selected_column_count - resolved_company_count
            if descriptor_count != 1:
                result["error"] = (
                    "Function + FTE / $B must start with exactly one L1/function "
                    "column followed by the company-value matrix. Do not include L0. "
                    f"Selected columns={selected_column_count}; companies={resolved_company_count}."
                )
                return result
            worksheet = workbook[fte_sheet]
            selected_matrix = _stream_range_values(
                worksheet,
                min_row=fte_header_row,
                max_row=fte_max_row,
                min_col=fte_min_col,
                max_col=fte_max_col,
            )
            function_header = str(
                ((selected_matrix[0][0] if selected_matrix and selected_matrix[0] else None) or "")
            ).strip()
            descriptor = _taxonomy_header_descriptor(function_header)
            if not (
                (descriptor is not None and descriptor[0] == 1)
                or _ui_function_header_score(function_header) > 0
            ):
                result["error"] = (
                    f"Function + FTE / $B begins with {function_header!r}; the first "
                    "selected column must be L1/Function."
                )
                return result
            function_label_column = fte_min_col
            source_labels = [
                str((row_values[0] if row_values else None) or "").strip()
                for row_values in selected_matrix[1:]
            ]
            function_label_range = (
                f"{get_column_letter(function_label_column)}{fte_data_min_row}:"
                f"{get_column_letter(function_label_column)}{fte_max_row}"
            )
            value_start_column = fte_min_col + 1
            detection_details = {
                "origin": "first_selected_l1_column",
                "function_header": function_header,
            }
            fte_value_range = (
                f"{get_column_letter(value_start_column)}{fte_data_min_row}:"
                f"{get_column_letter(fte_max_col)}{fte_max_row}"
            )
            result["descriptor_column_count"] = 1
            result["value_column_count"] = resolved_company_count
            result["fte_value_range"] = fte_value_range
        else:
            # Backward-compatible path for an older numeric-only mapping.
            explicit = _l1_explicit_function_labels(
                workbook,
                fte_sheet=fte_sheet,
                fte_min_row=fte_data_min_row,
                fte_max_row=fte_max_row,
                label_sheet=explicit_label_sheet,
                label_range=explicit_label_range,
            )
            if explicit is not None:
                source_labels, function_label_range, function_label_column = explicit
                detection_details = {"origin": "mapped_metadata"}
            else:
                (
                    source_labels,
                    function_label_range,
                    function_label_column,
                    detection_details,
                ) = _l1_detect_function_label_column(
                    workbook,
                    fte_sheet=fte_sheet,
                    fte_min_row=fte_data_min_row,
                    fte_max_row=fte_max_row,
                    fte_min_col=fte_min_col,
                    taxonomy_records=taxonomy_records,
                )
                detection_details["origin"] = "adjacent_column_detection"
            result["value_column_count"] = selected_column_count
            result["fte_value_range"] = (
                f"{get_column_letter(fte_min_col)}{fte_data_min_row}:"
                f"{get_column_letter(fte_max_col)}{fte_max_row}"
            )

        result["function_label_sheet"] = fte_sheet
        result["function_label_range"] = function_label_range
        result["function_label_column"] = function_label_column
        result["detection"] = detection_details

        source_positions: dict[str, list[int]] = {}
        for source_index, label in enumerate(source_labels):
            key = _l1_function_lookup_key(label)
            if key:
                source_positions.setdefault(key, []).append(source_index)

        taxonomy_key_counts: dict[str, int] = {}
        for record in taxonomy_records:
            key = str(record.get("key") or "")
            if key:
                taxonomy_key_counts[key] = taxonomy_key_counts.get(key, 0) + 1

        matches: list[dict[str, Any]] = []
        missing: list[str] = []
        ambiguous: list[str] = []
        used_source_indices: set[int] = set()
        for taxonomy_index, record in enumerate(taxonomy_records):
            function_name = str(record.get("function") or "").strip()
            key = str(record.get("key") or "")
            positions = source_positions.get(key, []) if key else []
            if not positions:
                missing.append(function_name)
                continue
            if len(positions) != 1 or taxonomy_key_counts.get(key, 0) != 1:
                ambiguous.append(function_name)
                continue
            source_index = positions[0]
            used_source_indices.add(source_index)
            matches.append(
                {
                    "taxonomy_index": taxonomy_index,
                    "taxonomy_row": int(record.get("taxonomy_row") or 0),
                    "l0": str(record.get("l0") or ""),
                    "taxonomy_function": function_name,
                    "source_index": source_index,
                    "source_excel_row": fte_data_min_row + source_index,
                    "source_function": source_labels[source_index],
                    "match_key": key,
                }
            )

        extras = [
            source_labels[index]
            for index in range(len(source_labels))
            if index not in used_source_indices and str(source_labels[index]).strip()
        ]
        result["matches"] = matches
        result["matched_count"] = len(matches)
        result["missing"] = missing
        result["ambiguous"] = ambiguous
        result["extra_source"] = extras
        result["ready"] = (
            not missing
            and not ambiguous
            and len(matches) == len(taxonomy_records)
        )
        return result
    except Exception as exc:
        result["error"] = str(exc)
        return result


# -----------------------------------------------------------------------------
# Aura L1 range auto-detection
# -----------------------------------------------------------------------------


def _ui_clean_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


_UI_MISSING_NUMERIC_TOKENS = {
    "", "-", "na", "n/a", "n\\a", "nd", "n/d", "n.d", "n.d.",
    "nan", "none", "<no data>", "not disclosed", "not available",
}
_UI_EXCEL_ERROR_TOKENS = {
    "#div/0!", "#n/a", "#value!", "#ref!", "#name?", "#num!", "#null!",
}


def _ui_is_missing_numeric_token(value: Any, *, include_blank: bool = True) -> bool:
    text = _ui_clean_text(value).casefold()
    if not include_blank and not text:
        return False
    return text in _UI_MISSING_NUMERIC_TOKENS or text in _UI_EXCEL_ERROR_TOKENS


def _ui_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, str):
        cleaned = value.replace("$", "").replace(",", "").replace("\u00a0", "").strip()
        if _ui_is_missing_numeric_token(cleaned):
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    try:
        return float(value)
    except Exception:
        return None


def _ui_grid_value(rows: list[list[Any]], row: int, column: int) -> Any:
    if row < 1 or column < 1 or row > len(rows):
        return None
    values = rows[row - 1]
    return values[column - 1] if column <= len(values) else None


def _ui_normalized_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", _ui_clean_text(value).lower()).strip()


def _ui_compact_text(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", _ui_clean_text(value).lower())


def _ui_quartile_kind(value: Any) -> str | None:
    text = _ui_normalized_text(value)
    compact = _ui_compact_text(value)
    tokens = set(text.split())
    if not text:
        return None
    if compact in {
        "median", "med", "middleq", "midq", "middlequartile", "midquartile",
        "q2", "p50", "50th", "50thpercentile",
    } or "median" in tokens:
        return "median"
    if compact in {
        "top", "topq", "topquartile", "topquart", "upperq", "upperquartile",
        "q1", "p25", "25th", "25thpercentile", "firstquartile", "1stquartile",
    } or (("top" in tokens or "upper" in tokens) and ({"q", "quartile"} & tokens)):
        return "top"
    if compact in {
        "bottom", "bottomq", "bottomquartile", "bottomquart", "lowerq", "lowerquartile",
        "q3", "p75", "75th", "75thpercentile", "thirdquartile", "3rdquartile",
    } or (("bottom" in tokens or "lower" in tokens) and ({"q", "quartile"} & tokens)):
        return "bottom"
    return None


def _ui_is_l0_header(value: Any) -> bool:
    return _ui_compact_text(value) in {
        "l0", "level0", "levelzero", "l0level", "l0function", "l0category", "categoryl0"
    }


def _ui_is_l1_header(value: Any) -> bool:
    compact = _ui_compact_text(value)
    text = _ui_normalized_text(value)
    return compact in {
        "l1", "level1", "levelone", "l1level", "l1function", "function",
        "functions", "functionname", "category", "workstream", "department", "area",
    } or text in {"level 1", "l 1"}


def _ui_function_header_score(value: Any) -> int:
    if _ui_is_l1_header(value):
        return 45
    if _ui_is_l0_header(value):
        return 5
    compact = _ui_compact_text(value)
    return 25 if compact.startswith("l1") or "function" in compact or "category" in compact else 0


def _ui_looks_like_function(value: Any) -> bool:
    text = _ui_clean_text(value)
    return bool(
        text
        and not text.startswith(">>")
        and not _ui_is_missing_numeric_token(text, include_blank=False)
        and _ui_number(text) is None
        and re.search(r"[A-Za-z]", text)
    )



def _ui_is_combined_hierarchy_header(value: Any) -> bool:
    """Return True for headers that explicitly combine L1 and L0 in one column."""
    compact = _ui_compact_text(value)
    return compact in {
        "l1l0", "l0l1", "level1level0", "level0level1",
        "l1andl0", "l0andl1", "level1andlevel0", "level0andlevel1",
    }


def _ui_contiguous_sum_match(
    rows: list[list[Any]],
    target_row: int,
    source_first_row: int,
    source_last_row: int,
    company_start_column: int,
    company_end_column: int,
) -> bool:
    """Check whether a row is the sum of a contiguous earlier row block.

    Only columns with numeric values in both the target and every source row are
    compared. This lets ordinary missing observations remain missing without
    weakening the structural check for the populated peer columns.
    """
    company_count = company_end_column - company_start_column + 1
    required_comparisons = max(3, min(6, max(1, company_count // 3)))
    compared = 0
    matched = 0
    for column in range(company_start_column, company_end_column + 1):
        target = _ui_number(_ui_grid_value(rows, target_row, column))
        if target is None:
            continue
        source_values = [
            _ui_number(_ui_grid_value(rows, source_row, column))
            for source_row in range(source_first_row, source_last_row + 1)
        ]
        if any(value is None for value in source_values):
            continue
        compared += 1
        source_sum = sum(float(value) for value in source_values if value is not None)
        tolerance = max(1e-6, 1e-6 * max(1.0, abs(float(target)), abs(source_sum)))
        if abs(float(target) - source_sum) <= tolerance:
            matched += 1
    return compared >= required_comparisons and matched / compared >= 0.90


def _ui_find_contiguous_aggregate_source(
    rows: list[list[Any]],
    target_row: int,
    detail_first_row: int,
    detail_last_row: int,
    company_start_column: int,
    company_end_column: int,
) -> tuple[int, int] | None:
    for source_first in range(detail_first_row, detail_last_row + 1):
        for source_last in range(source_first, detail_last_row + 1):
            if _ui_contiguous_sum_match(
                rows,
                target_row,
                source_first,
                source_last,
                company_start_column,
                company_end_column,
            ):
                return source_first, source_last
    return None


def _ui_detect_embedded_aggregate_tail(
    rows: list[list[Any]],
    header_row: int,
    function_column: int,
    company_start_column: int,
    company_end_column: int,
    candidate_end_row: int,
) -> tuple[int, int, int] | None:
    """Split an explicitly combined L1/L0 table using aggregation evidence.

    Some source tables place detailed L1 rows first and L0 aggregate rows at the
    bottom of the same table. When the header explicitly says L1/L0 (or L0/L1),
    identify a trailing block whose company values can each be reconstructed as
    sums of contiguous earlier detail rows. No function names or row numbers are
    assumed.
    """
    if not _ui_is_combined_hierarchy_header(
        _ui_grid_value(rows, header_row, function_column)
    ):
        return None

    detail_first = header_row + 1
    if candidate_end_row - detail_first + 1 < 4:
        return None

    # Leave at least two detail rows and require at least two aggregate rows.
    for aggregate_start in range(detail_first + 2, candidate_end_row):
        detail_last = aggregate_start - 1
        aggregate_rows = range(aggregate_start, candidate_end_row + 1)
        matches: list[tuple[int, int]] = []
        valid = True
        for target_row in aggregate_rows:
            source = _ui_find_contiguous_aggregate_source(
                rows,
                target_row,
                detail_first,
                detail_last,
                company_start_column,
                company_end_column,
            )
            if source is None:
                valid = False
                break
            matches.append(source)
        if not valid:
            continue
        # At least one true multi-row aggregation is required; repeated detail
        # rows alone are not sufficient evidence of an L0 aggregate tail.
        if not any(source_last > source_first for source_first, source_last in matches):
            continue
        return detail_last, aggregate_start, candidate_end_row
    return None

def _ui_contiguous_nonblank_runs(
    rows: list[list[Any]],
    row: int,
    start_col: int,
    end_col: int,
) -> list[tuple[int, int]]:
    result: list[tuple[int, int]] = []
    current: int | None = None
    for column in range(start_col, end_col + 1):
        nonblank = _ui_clean_text(_ui_grid_value(rows, row, column)) != ""
        if nonblank and current is None:
            current = column
        elif not nonblank and current is not None:
            result.append((current, column - 1))
            current = None
    if current is not None:
        result.append((current, end_col))
    return result


def _ui_candidate_company_ranges(
    rows: list[list[Any]],
    header_row: int,
    first_quartile_col: int,
) -> list[tuple[int | None, int, int, int, int]]:
    candidates: list[tuple[int | None, int, int, int, int]] = []
    runs = _ui_contiguous_nonblank_runs(rows, header_row, 1, first_quartile_col - 1)
    for run_start, run_end in runs:
        run_len = run_end - run_start + 1
        if run_len < 3:
            continue
        if (
            run_len >= 5
            and _ui_is_l0_header(_ui_grid_value(rows, header_row, run_start))
            and _ui_is_l1_header(_ui_grid_value(rows, header_row, run_start + 1))
        ):
            candidates.append((run_start, run_start + 1, run_start + 2, run_end, 180 + run_len))
            continue
        if run_len >= 4 and _ui_is_l1_header(_ui_grid_value(rows, header_row, run_start)):
            l0_col = (
                run_start - 1
                if run_start > 1 and _ui_is_l0_header(_ui_grid_value(rows, header_row, run_start - 1))
                else None
            )
            candidates.append((l0_col, run_start, run_start + 1, run_end, 125 + run_len))
            continue
        if run_start > 1:
            function_col = run_start - 1
            l0_col = (
                function_col - 1
                if function_col > 1 and _ui_is_l0_header(_ui_grid_value(rows, header_row, function_col - 1))
                else None
            )
            score = 45 + run_len + _ui_function_header_score(
                _ui_grid_value(rows, header_row, function_col)
            )
            candidates.append((l0_col, function_col, run_start, run_end, score))
        if run_len >= 4 and not _ui_is_l0_header(_ui_grid_value(rows, header_row, run_start)):
            candidates.append((None, run_start, run_start + 1, run_end, 40 + run_len))

    for function_col in range(1, first_quartile_col):
        if not _ui_is_l1_header(_ui_grid_value(rows, header_row, function_col)):
            continue
        l0_col = (
            function_col - 1
            if function_col > 1 and _ui_is_l0_header(_ui_grid_value(rows, header_row, function_col - 1))
            else None
        )
        for company_start, company_end in _ui_contiguous_nonblank_runs(
            rows, header_row, function_col + 1, first_quartile_col - 1
        ):
            if company_end - company_start + 1 >= 3:
                candidates.append(
                    (l0_col, function_col, company_start, company_end, 110 + company_end - company_start + 1)
                )

    best: dict[tuple[int | None, int, int, int], int] = {}
    for l0_col, function_col, company_start, company_end, score in candidates:
        if company_start <= function_col or company_end < company_start:
            continue
        key = (l0_col, function_col, company_start, company_end)
        best[key] = max(score, best.get(key, -1))
    return [(a, b, c, d, score) for (a, b, c, d), score in best.items()]


def _ui_revenue_label_score(value: Any) -> int:
    """Score the revenue label in the first Aura Summary table.

    Bare ``Sales`` L1 rows are not revenue labels. Valid labels include
    Revenue, Revenue ($), Revenue ($B), Revenue ($B) Latest Year, LTM Revenue,
    Total Revenue, and Net Sales.
    """
    raw = _ui_clean_text(value)
    text = _ui_normalized_text(value)
    compact = _ui_compact_text(value)
    tokens = set(text.split())
    if not text or tokens & {
        "source", "comment", "coverage", "employee", "employees", "fte",
        "headcount", "margin", "growth", "gap", "percent", "percentage",
        "cagr", "multiple", "ebitda", "opex",
    }:
        return 0

    has_revenue = bool({"revenue", "revenues", "rev"} & tokens)
    has_qualified_sales = "sales" in tokens and bool(
        {"net", "total", "latest", "ltm", "ttm", "b", "bn", "billion", "m", "mm", "million", "usd"}
        & tokens
        or "$" in raw
    )
    if not (has_revenue or has_qualified_sales):
        return 0

    score = 40
    if "ltm" in tokens or "ttm" in tokens or "l12m" in compact:
        score += 90
    if "last" in tokens and ({"twelve", "12"} & tokens) and ({"months", "month"} & tokens):
        score += 90
    if "total" in tokens:
        score += 40
    if "net" in tokens and "sales" in tokens:
        score += 30
    if "latest" in tokens:
        score += 30
    if "$" in raw or "usd" in tokens:
        score += 15
    has_billions = bool({"b", "bn", "billion", "billions"} & tokens)
    has_millions = bool({"m", "mm", "million", "millions"} & tokens)
    if has_billions:
        score += 25
    if has_millions:
        score += 15
    if "year" in tokens and not ("latest" in tokens or "$" in raw or has_billions or has_millions):
        score -= 35
    return max(score, 0)


def _ui_contains_fte_token(value: Any) -> bool:
    raw = _ui_clean_text(value)
    text = _ui_normalized_text(raw)
    compact = _ui_compact_text(raw)
    tokens = set(text.split())
    if {"fte", "ftes"} & tokens:
        return True
    if "fulltimeequivalent" in compact or "fulltimeequivalents" in compact:
        return True
    return compact.startswith("fteper") or compact.startswith("ftesper")


def _ui_fte_section_marker_score(value: Any) -> int:
    """Score labels such as ``FTEs / $B (Avg)`` without matching ``after``."""
    raw = _ui_clean_text(value)
    text = _ui_normalized_text(raw)
    compact = _ui_compact_text(raw)
    tokens = set(text.split())
    if not _ui_contains_fte_token(raw):
        return 0
    if "%" in raw or "percent" in tokens or "percentage" in tokens or "tfte" in compact:
        return 0
    score = 35
    if "revenue" in tokens or "revenues" in tokens or "rev" in tokens or "sales" in tokens:
        score += 45
    if "$" in raw or "usd" in tokens:
        score += 20
    if "per" in tokens or "/" in raw:
        score += 30
    if "1b" in compact or "billion" in tokens or "b" in tokens or "bn" in tokens:
        score += 50
    if "avg" in tokens or "average" in tokens:
        score += 20
    if raw.lstrip().startswith(">>"):
        score += 10
    return score


def _ui_revenue_value_span(
    rows: list[list[Any]],
    row: int,
    label_column: int,
    max_column: int,
) -> tuple[int, int] | None:
    start_column: int | None = None
    for column in range(label_column + 1, min(max_column, label_column + 4) + 1):
        value = _ui_grid_value(rows, row, column)
        if _ui_number(value) is not None or _ui_is_missing_numeric_token(value, include_blank=False):
            start_column = column
            break
    if start_column is None:
        return None

    end_column = start_column - 1
    blank_streak = 0
    for column in range(start_column, max_column + 1):
        value = _ui_grid_value(rows, row, column)
        text = _ui_clean_text(value)
        if _ui_number(value) is not None or _ui_is_missing_numeric_token(value, include_blank=False):
            end_column = column
            blank_streak = 0
            continue
        if not text:
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        break
    if end_column - start_column + 1 < 3:
        return None
    return start_column, end_column


def _ui_company_header_row_score(
    rows: list[list[Any]],
    row: int,
    start_column: int,
    end_column: int,
) -> int | None:
    values = [
        _ui_clean_text(_ui_grid_value(rows, row, column))
        for column in range(start_column, end_column + 1)
    ]
    nonblank = [value for value in values if value]
    expected_count = end_column - start_column + 1
    if len(nonblank) < max(3, int(expected_count * 0.80)):
        return None
    if any(_ui_number(value) is not None for value in nonblank):
        return None
    keys = [_ui_compact_text(value) for value in nonblank if _ui_compact_text(value)]
    unique_count = len(set(keys))
    unique_ratio = unique_count / max(1, len(keys))
    if unique_ratio < 0.60:
        return None
    generic_keys = {
        "public", "private", "annualreport", "companywebsite", "rds",
        "source", "reported", "client", "peer", "ips", "amc", "pes",
    }
    if sum(1 for key in keys if key in generic_keys) > len(keys) * 0.25:
        return None
    preceding_key = (
        _ui_compact_text(_ui_grid_value(rows, row, start_column - 1))
        if start_column > 1
        else ""
    )
    explicit_label_bonus = 80 if preceding_key in {
        "company", "companies", "companyname", "companynames",
        "peercompanies", "peers",
    } else 0
    return (
        100
        + len(nonblank) * 5
        + unique_count * 4
        + int(unique_ratio * 40)
        + explicit_label_bonus
    )


def _ui_detect_primary_company_revenue_table(rows: list[list[Any]]) -> dict[str, Any]:
    max_row = len(rows)
    max_column = max((len(row) for row in rows), default=0)
    marker_rows = [
        row
        for row in range(1, max_row + 1)
        if max(
            (
                _ui_fte_section_marker_score(_ui_grid_value(rows, row, column))
                for column in range(1, max_column + 1)
            ),
            default=0,
        ) >= 80
    ]
    search_end_row = min(marker_rows) - 1 if marker_rows else max_row
    candidates: list[dict[str, Any]] = []

    for revenue_row in range(1, search_end_row + 1):
        for label_column in range(1, max_column + 1):
            label = _ui_grid_value(rows, revenue_row, label_column)
            label_score = _ui_revenue_label_score(label)
            if label_score <= 0:
                continue
            span = _ui_revenue_value_span(rows, revenue_row, label_column, max_column)
            if span is None:
                continue
            company_start_column, company_end_column = span
            numeric_values = [
                _ui_number(_ui_grid_value(rows, revenue_row, column))
                for column in range(company_start_column, company_end_column + 1)
            ]
            numeric_values = [value for value in numeric_values if value is not None]
            if len(numeric_values) < 3:
                continue
            year_like_count = sum(
                1
                for value in numeric_values
                if 1900 <= float(value) <= 2100 and abs(float(value) - round(float(value))) < 1e-9
            )
            label_tokens = set(_ui_normalized_text(label).split())
            label_has_units = bool(
                "$" in _ui_clean_text(label)
                or {"b", "bn", "billion", "billions", "m", "mm", "million", "millions", "usd"}
                & label_tokens
            )
            if (
                year_like_count / len(numeric_values) >= 0.60
                and "year" in label_tokens
                and not label_has_units
            ):
                continue

            for company_row in range(max(1, revenue_row - 30), revenue_row):
                company_score = _ui_company_header_row_score(
                    rows,
                    company_row,
                    company_start_column,
                    company_end_column,
                )
                if company_score is None:
                    continue
                distance = revenue_row - company_row
                score = (
                    label_score
                    + len(numeric_values) * 12
                    + company_score
                    + max(0, 35 - distance)
                    - revenue_row * 0.05
                )
                candidates.append(
                    {
                        "score": score,
                        "company_header_row": company_row,
                        "company_start_column": company_start_column,
                        "company_end_column": company_end_column,
                        "company_headers": [
                            _ui_clean_text(_ui_grid_value(rows, company_row, column))
                            for column in range(company_start_column, company_end_column + 1)
                        ],
                        "revenue_row": revenue_row,
                        "revenue_label_column": label_column,
                        "revenue_label": _ui_clean_text(label),
                    }
                )

    if not candidates:
        raise ValueError(
            "Could not auto-detect the first Aura company/revenue table. Expected "
            "a unique company row above a labeled Revenue row."
        )
    candidates.sort(
        key=lambda item: (float(item["score"]), -int(item["revenue_row"])),
        reverse=True,
    )
    return candidates[0]


def _ui_find_company_sequence(
    rows: list[list[Any]],
    row: int,
    company_headers: list[str],
    first_column: int,
    last_column: int,
) -> tuple[int, int] | None:
    expected = [_ui_compact_text(value) for value in company_headers]
    if not expected or any(not key for key in expected):
        return None
    width = len(expected)
    if last_column - first_column + 1 < width:
        return None
    for start_column in range(first_column, last_column - width + 2):
        observed = [
            _ui_compact_text(_ui_grid_value(rows, row, column))
            for column in range(start_column, start_column + width)
        ]
        if observed == expected:
            return start_column, start_column + width - 1
    return None


def _ui_detect_aura_layout_legacy(rows: list[list[Any]]) -> dict[str, Any]:
    max_row = len(rows)
    max_column = max((len(row) for row in rows), default=0)
    candidates: list[dict[str, Any]] = []
    for row in range(1, max_row + 1):
        quartiles: dict[str, int] = {}
        for column in range(1, max_column + 1):
            kind = _ui_quartile_kind(_ui_grid_value(rows, row, column))
            if kind and kind not in quartiles:
                quartiles[kind] = column
        if not {"top", "median", "bottom"}.issubset(quartiles):
            continue
        qcols = [quartiles["top"], quartiles["median"], quartiles["bottom"]]
        for l0_col, function_col, company_start, company_end, structural in _ui_candidate_company_ranges(
            rows, row, min(qcols)
        ):
            company_count = company_end - company_start + 1
            if company_count < 2:
                continue
            valid_rows = 0
            min_company_numbers = max(2, min(4, company_count // 2 if company_count else 2))
            for data_row in range(row + 1, min(max_row, row + 50) + 1):
                if not _ui_looks_like_function(_ui_grid_value(rows, data_row, function_col)):
                    continue
                company_numbers = sum(
                    1
                    for column in range(company_start, company_end + 1)
                    if _ui_number(_ui_grid_value(rows, data_row, column)) is not None
                )
                quartile_numbers = sum(
                    1 for column in qcols if _ui_number(_ui_grid_value(rows, data_row, column)) is not None
                )
                if company_numbers >= min_company_numbers and quartile_numbers >= 2:
                    valid_rows += 1
            if valid_rows < 2:
                continue
            score = (
                structural
                + _ui_function_header_score(_ui_grid_value(rows, row, function_col))
                + (35 if l0_col is not None else 0)
                + 120
                + min(company_count, 20) * 4
                + min(valid_rows, 30) * 10
            )
            candidates.append(
                {
                    "score": score,
                    "header_row": row,
                    "l0_column": l0_col,
                    "function_column": function_col,
                    "company_start_column": company_start,
                    "company_end_column": company_end,
                    "quartile_columns": dict(quartiles),
                    "company_count": company_count,
                    "valid_rows": valid_rows,
                }
            )
    if not candidates:
        raise ValueError(
            "Could not auto-detect the Aura L1 table. Expected company headers plus Top Q, Median, and Bottom Q columns."
        )
    candidates.sort(key=lambda item: int(item["score"]), reverse=True)
    layout = candidates[0]

    qcols = list(layout["quartile_columns"].values())
    end_row = int(layout["header_row"])
    for row in range(int(layout["header_row"]) + 1, max_row + 1):
        function_name = _ui_clean_text(_ui_grid_value(rows, row, int(layout["function_column"])))
        company_numbers = sum(
            1
            for column in range(int(layout["company_start_column"]), int(layout["company_end_column"]) + 1)
            if _ui_number(_ui_grid_value(rows, row, column)) is not None
        )
        quartile_numbers = sum(
            1 for column in qcols if _ui_number(_ui_grid_value(rows, row, column)) is not None
        )
        if not function_name:
            if end_row > int(layout["header_row"]):
                break
            continue
        if function_name.startswith(">>") and end_row > int(layout["header_row"]):
            break
        if company_numbers == 0 and quartile_numbers == 0 and end_row > int(layout["header_row"]):
            break
        if _ui_looks_like_function(function_name):
            end_row = row
    if end_row <= int(layout["header_row"]):
        raise ValueError("The detected Aura L1 table contains no function rows.")
    layout["fte_start_row"] = int(layout["header_row"]) + 1
    layout["fte_end_row"] = end_row

    revenue_candidates: list[tuple[int, int, int]] = []
    for row in range(1, max_row + 1):
        for column in range(1, min(max_column, int(layout["company_start_column"])) + 1):
            base = _ui_revenue_label_score(_ui_grid_value(rows, row, column))
            if base <= 0:
                continue
            numeric_count = sum(
                1
                for value_col in range(int(layout["company_start_column"]), int(layout["company_end_column"]) + 1)
                if _ui_number(_ui_grid_value(rows, row, value_col)) is not None
            )
            if numeric_count:
                revenue_candidates.append((base + numeric_count * 14, row, column))
    if not revenue_candidates:
        raise ValueError(
            "Could not auto-detect a Revenue row aligned with the Aura company columns."
        )
    revenue_candidates.sort(reverse=True)
    _, revenue_row, revenue_label_column = revenue_candidates[0]
    layout["revenue_row"] = revenue_row
    layout["revenue_label_column"] = revenue_label_column
    return layout


def _ui_detect_aura_layout(rows: list[list[Any]]) -> dict[str, Any]:
    """Detect the standardized Aura inputs from the canonical first table and FTE/$B section."""
    max_row = len(rows)
    max_column = max((len(row) for row in rows), default=0)
    primary = _ui_detect_primary_company_revenue_table(rows)
    canonical_headers = list(primary["company_headers"])
    canonical_count = len(canonical_headers)
    candidates: list[dict[str, Any]] = []

    for header_row in range(1, max_row + 1):
        quartiles: dict[str, int] = {}
        for column in range(1, max_column + 1):
            kind = _ui_quartile_kind(_ui_grid_value(rows, header_row, column))
            if kind and kind not in quartiles:
                quartiles[kind] = column
        if not {"top", "median", "bottom"}.issubset(quartiles):
            continue

        first_quartile_column = min(quartiles.values())
        company_sequence = _ui_find_company_sequence(
            rows,
            header_row,
            canonical_headers,
            1,
            first_quartile_column - 1,
        )
        if company_sequence is None:
            continue
        company_start_column, company_end_column = company_sequence
        if company_end_column - company_start_column + 1 != canonical_count:
            continue

        function_column = company_start_column - 1
        if function_column < 1 or _ui_function_header_score(
            _ui_grid_value(rows, header_row, function_column)
        ) <= 0:
            continue
        l0_column = (
            function_column - 1
            if function_column > 1
            and _ui_is_l0_header(_ui_grid_value(rows, header_row, function_column - 1))
            else None
        )

        marker_row: int | None = None
        marker_score = 0
        for scan_row in range(max(1, header_row - 16), header_row + 1):
            row_score = max(
                (
                    _ui_fte_section_marker_score(
                        _ui_grid_value(rows, scan_row, column)
                    )
                    for column in range(1, max_column + 1)
                ),
                default=0,
            )
            if row_score > marker_score:
                marker_row = scan_row
                marker_score = row_score
        if marker_row is None or marker_score < 80:
            continue

        qcols = [quartiles["top"], quartiles["median"], quartiles["bottom"]]
        valid_rows = 0
        for data_row in range(header_row + 1, min(max_row, header_row + 60) + 1):
            function_name = _ui_clean_text(
                _ui_grid_value(rows, data_row, function_column)
            )
            function_key = _ui_compact_text(function_name)
            if function_key == "total" or function_key.startswith("totalexcluding"):
                break
            company_numbers = sum(
                1
                for column in range(company_start_column, company_end_column + 1)
                if _ui_number(_ui_grid_value(rows, data_row, column)) is not None
            )
            quartile_numbers = sum(
                1
                for column in qcols
                if _ui_number(_ui_grid_value(rows, data_row, column)) is not None
            )
            if (
                _ui_looks_like_function(function_name)
                and company_numbers >= max(2, min(4, canonical_count // 2))
                and quartile_numbers >= 2
            ):
                valid_rows += 1
            elif valid_rows and (
                not function_name or (company_numbers == 0 and quartile_numbers == 0)
            ):
                break
        if valid_rows < 2:
            continue

        score = (
            400
            + marker_score
            + _ui_function_header_score(
                _ui_grid_value(rows, header_row, function_column)
            )
            + (35 if l0_column is not None else 0)
            + min(canonical_count, 30) * 5
            + min(valid_rows, 40) * 10
            + max(0, 20 - (header_row - marker_row))
        )
        candidates.append(
            {
                "score": score,
                "header_row": header_row,
                "l0_column": l0_column,
                "function_column": function_column,
                "company_start_column": company_start_column,
                "company_end_column": company_end_column,
                "quartile_columns": dict(quartiles),
                "company_count": canonical_count,
                "valid_rows": valid_rows,
                "fte_section_row": marker_row,
                "fte_section_score": marker_score,
            }
        )

    if not candidates:
        raise ValueError(
            "Could not auto-detect the FTE/$B table beneath an FTEs/$B (Avg) "
            "section label using the company order from the first Aura table."
        )
    candidates.sort(key=lambda item: int(item["score"]), reverse=True)
    layout = candidates[0]

    qcols = list(layout["quartile_columns"].values())
    end_row = int(layout["header_row"])
    for row in range(int(layout["header_row"]) + 1, max_row + 1):
        function_name = _ui_clean_text(
            _ui_grid_value(rows, row, int(layout["function_column"]))
        )
        function_key = _ui_compact_text(function_name)
        if function_key == "total" or function_key.startswith("totalexcluding"):
            break
        company_numbers = sum(
            1
            for column in range(
                int(layout["company_start_column"]),
                int(layout["company_end_column"]) + 1,
            )
            if _ui_number(_ui_grid_value(rows, row, column)) is not None
        )
        quartile_numbers = sum(
            1
            for column in qcols
            if _ui_number(_ui_grid_value(rows, row, column)) is not None
        )
        if not function_name:
            if end_row > int(layout["header_row"]):
                break
            continue
        if function_name.startswith(">>") and end_row > int(layout["header_row"]):
            break
        if company_numbers == 0 and quartile_numbers == 0 and end_row > int(layout["header_row"]):
            break
        if _ui_looks_like_function(function_name):
            end_row = row
    combined_split = _ui_detect_embedded_aggregate_tail(
        rows,
        int(layout["header_row"]),
        int(layout["function_column"]),
        int(layout["company_start_column"]),
        int(layout["company_end_column"]),
        end_row,
    )
    if combined_split is not None:
        l1_end_row, aggregate_start_row, aggregate_end_row = combined_split
        end_row = l1_end_row
        layout["embedded_l0_start_row"] = aggregate_start_row
        layout["embedded_l0_end_row"] = aggregate_end_row

    if end_row <= int(layout["header_row"]):
        raise ValueError("The detected FTE/$B table contains no function rows.")

    layout["fte_start_row"] = int(layout["header_row"]) + 1
    layout["fte_end_row"] = end_row
    layout["company_header_row"] = int(primary["company_header_row"])
    layout["canonical_company_start_column"] = int(primary["company_start_column"])
    layout["canonical_company_end_column"] = int(primary["company_end_column"])
    layout["canonical_company_headers"] = canonical_headers
    layout["revenue_row"] = int(primary["revenue_row"])
    layout["revenue_label_column"] = int(primary["revenue_label_column"])
    layout["revenue_label"] = str(primary["revenue_label"])
    layout["primary_table_score"] = float(primary["score"])
    return layout


def _ui_detect_l0_quartile_block(
    rows: list[list[Any]],
    layout: dict[str, Any],
) -> dict[str, Any]:
    """Find the lower L0 benchmark rows sharing the detected quartile columns.

    The structural Aura pattern keeps L1 and L0 benchmark values in the same
    three consecutive quartile columns. L0 rows appear after the L1 FTE table
    and carry their L0 label somewhere to the left of those columns. Exact row
    numbers are intentionally not assumed.
    """
    qcols = sorted(int(value) for value in dict(layout["quartile_columns"]).values())
    if len(qcols) != 3 or qcols[-1] - qcols[0] != 2:
        raise ValueError(
            "The detected quartile headers are not three consecutive columns."
        )
    max_row = len(rows)

    embedded_start = int(layout.get("embedded_l0_start_row") or 0)
    embedded_end = int(layout.get("embedded_l0_end_row") or 0)
    if embedded_start and embedded_end >= embedded_start:
        function_column = int(layout["function_column"])
        embedded_rows: list[tuple[int, str]] = []
        for row in range(embedded_start, embedded_end + 1):
            label = _ui_clean_text(_ui_grid_value(rows, row, function_column))
            numeric_count = sum(
                1
                for column in qcols
                if _ui_number(_ui_grid_value(rows, row, column)) is not None
            )
            if _ui_looks_like_function(label) and numeric_count >= 2:
                embedded_rows.append((row, label))
        if len(embedded_rows) >= 2:
            return {
                "first_row": embedded_rows[0][0],
                "last_row": embedded_rows[-1][0],
                "rows": [row for row, _label in embedded_rows],
                "labels": [label for _row, label in embedded_rows],
                "header_row": int(layout["header_row"]),
                "first_column": qcols[0],
                "last_column": qcols[-1],
                "source": "embedded_aggregate_tail",
            }

    first_scan_row = int(layout["fte_end_row"]) + 1
    label_search_end = max(1, qcols[0] - 1)
    candidates: list[tuple[int, str]] = []
    started = False
    gap_after_start = 0
    ignored_keys = {
        "total", "overall", "aggregate", "aggregated",
        "totalexcludingoosandunmapped", "totalexcludingoosunmapped",
    }

    for row in range(first_scan_row, min(max_row, first_scan_row + 180) + 1):
        # A later repeated quartile header marks the next table.
        kinds = {
            _ui_quartile_kind(_ui_grid_value(rows, row, column))
            for column in qcols
        }
        if started and {"top", "median", "bottom"}.issubset(kinds):
            break

        numeric_count = sum(
            1
            for column in qcols
            if _ui_number(_ui_grid_value(rows, row, column)) is not None
        )
        label = ""
        if numeric_count >= 2:
            for column in range(label_search_end, 0, -1):
                value = _ui_grid_value(rows, row, column)
                if _ui_number(value) is None and _ui_looks_like_function(value):
                    candidate = _ui_clean_text(value)
                    if _ui_compact_text(candidate) not in ignored_keys:
                        label = candidate
                        break

        if numeric_count >= 2 and label:
            candidates.append((row, label))
            started = True
            gap_after_start = 0
            continue

        if started:
            gap_after_start += 1
            if gap_after_start >= 4:
                break

    if not candidates:
        raise ValueError(
            "Could not auto-detect the lower L0 quartile table in the three "
            "benchmark columns beneath the L1 FTE/$B table."
        )

    return {
        "first_row": candidates[0][0],
        "last_row": candidates[-1][0],
        "rows": [row for row, _label in candidates],
        "labels": [label for _row, label in candidates],
        "header_row": int(layout["header_row"]),
        "first_column": qcols[0],
        "last_column": qcols[-1],
    }


@st.cache_data(show_spinner=False, max_entries=24)
def aura_l1_auto_input_profile(
    path_text: str,
    cache_token: int,
    requested_sheet: str,
) -> dict[str, Any]:
    """Auto-detect the standardized non-taxonomy Aura ranges for L0/L1."""
    from openpyxl.utils import get_column_letter

    result: dict[str, Any] = {
        "sheet_name": "",
        "ranges": {},
        "layout": {},
        "error": "",
    }
    try:
        workbook = _analysis_workbook(path_text, cache_token)
    except Exception as exc:
        result["error"] = f"Aura Summary could not be opened for input detection: {exc}"
        return result

    try:
        sheet_names = list(workbook.sheetnames)
        selected = _match_sheet_name(sheet_names, requested_sheet)
        if selected is None:
            aura_candidates = [
                name for name in sheet_names if _preview_sheet_role(name) == "aura"
            ]
            selected = aura_candidates[0] if aura_candidates else None
        if selected is None:
            result["error"] = "No Aura Summary worksheet was detected for Aura input auto-detection."
            return result

        worksheet = workbook[selected]
        max_rows = max(1, min(int(worksheet.max_row or 1), 1600))
        max_columns = max(1, min(int(worksheet.max_column or 1), 180))
        rows = [
            list(row)
            for row in worksheet.iter_rows(
                min_row=1,
                max_row=max_rows,
                min_col=1,
                max_col=max_columns,
                values_only=True,
            )
        ]
        layout = _ui_detect_aura_layout(rows)
        fte_company_start = int(layout["company_start_column"])
        fte_company_end = int(layout["company_end_column"])
        company_header_row = int(layout["company_header_row"])
        canonical_company_start = int(layout["canonical_company_start_column"])
        canonical_company_end = int(layout["canonical_company_end_column"])
        header_row = int(layout["header_row"])
        first_row = int(layout["fte_start_row"])
        last_row = int(layout["fte_end_row"])
        qcols = dict(layout["quartile_columns"])

        def row_range(row: int, first_col: int, last_col: int) -> str:
            return (
                f"{get_column_letter(first_col)}{row}:"
                f"{get_column_letter(last_col)}{row}"
            )

        def column_range(column: int, first: int, last: int) -> str:
            letter = get_column_letter(column)
            return f"{letter}{first}:{letter}{last}"

        result["sheet_name"] = selected
        result["layout"] = layout
        function_column = int(layout["function_column"])
        result["function_labels_range"] = column_range(
            function_column,
            first_row,
            last_row,
        )
        revenue_label_column = int(layout["revenue_label_column"])
        expected_revenue_label_column = canonical_company_start - 1
        if expected_revenue_label_column < 1 or revenue_label_column != expected_revenue_label_column:
            raise ValueError(
                "The detected Revenue label is not immediately before the company "
                "values. Select the labeled Revenue row manually."
            )

        q_column_numbers = sorted(int(value) for value in qcols.values())
        if len(q_column_numbers) != 3 or q_column_numbers[-1] - q_column_numbers[0] != 2:
            raise ValueError(
                "Top, Median, and Bottom quartiles must occupy three consecutive columns."
            )
        quartile_first_col = q_column_numbers[0]
        quartile_last_col = q_column_numbers[-1]
        try:
            l0_quartile_block = _ui_detect_l0_quartile_block(rows, layout)
            layout["l0_quartile_block"] = l0_quartile_block
            result["l0_quartile_error"] = ""
        except Exception as exc:
            l0_quartile_block = None
            result["l0_quartile_error"] = str(exc)

        ranges = {
            # Company names are value-only. Revenue/FTE include their labels.
            # Quartile inputs are complete three-column tables, with semantic
            # column order resolved from the detected header row.
            "company_headers": row_range(
                company_header_row, canonical_company_start, canonical_company_end
            ),
            "revenue_values": row_range(
                int(layout["revenue_row"]),
                revenue_label_column,
                canonical_company_end,
            ),
            "fte_values": (
                f"{get_column_letter(function_column)}{header_row}:"
                f"{get_column_letter(fte_company_end)}{last_row}"
            ),
            "l1_quartile_table": (
                f"{get_column_letter(quartile_first_col)}{header_row}:"
                f"{get_column_letter(quartile_last_col)}{last_row}"
            ),
        }
        if l0_quartile_block is not None:
            ranges["l0_quartile_table"] = (
                f"{get_column_letter(quartile_first_col)}{int(l0_quartile_block['first_row'])}:"
                f"{get_column_letter(quartile_last_col)}{int(l0_quartile_block['last_row'])}"
            )
        result["ranges"] = ranges
        result["quartile_header_row"] = header_row
        result["quartile_columns"] = dict(qcols)
        return result
    except Exception as exc:
        result["sheet_name"] = selected or ""
        result["error"] = str(exc)
        return result
    finally:
        # Shared cached workbook; callers only read from it.
        pass




def _geocost_header_kind(value: Any) -> str:
    key = _ui_compact_text(value)
    if key == "region":
        return "region"
    if key in {"highlow", "highorlow", "costtier"}:
        return "high_low"
    if key in {"country", "state", "location", "geography"}:
        return "entity"
    if key in {
        "geocostanalysis", "geocostlocation", "displaycountry", "displaystate",
        "countryname", "statename",
    }:
        return "display"
    return ""


def _geocost_low_cost_row_score(value: Any) -> int:
    """Score labels for the row shown at the bottom of the Geo Cost chart."""
    raw = _ui_clean_text(value)
    text = re.sub(r"[^a-z0-9%]+", " ", raw.lower()).strip()
    compact = _ui_compact_text(raw)
    tokens = set(text.replace("%", " % ").split())
    if not text or "low" not in tokens or "cost" not in tokens:
        return 0

    score = 80
    if "%" in raw or "percent" in tokens or "percentage" in tokens or "pct" in tokens:
        score += 45
    if "fte" in compact or "ftes" in compact:
        score += 55
    if "region" in tokens or "regions" in tokens:
        score += 25
    if "location" in tokens or "locations" in tokens:
        score += 25
    if "workforce" in tokens or "employee" in tokens or "employees" in tokens:
        score += 20
    if compact in {
        "lowcost%", "lowcostpct", "lowcostpercentage", "lowcostfte%",
        "fteinlowcostlocations%", "ftesinlowcostlocations%",
        "percentofftesinlowcostregion", "percentageofftesinlowcostregion",
    }:
        score += 100
    return score


def _geocost_percentage_number(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        cleaned = (
            value.replace("%", "")
            .replace(",", "")
            .replace("\u00a0", "")
            .strip()
        )
        if cleaned.lower() in {"", "-", "na", "n/a", "nan", "none", "<no data>"}:
            return None
        try:
            number = float(cleaned)
        except ValueError:
            return None
        return number / 100.0 if "%" in value or abs(number) > 1.0000001 else number
    try:
        number = float(value)
    except Exception:
        return None
    return number / 100.0 if abs(number) > 1.0000001 else number


def _geocost_marker_candidates(
    worksheet,
    *,
    max_col: int,
    end_row: int,
    start_row: int = 1,
) -> list[dict[str, Any]]:
    rows = _stream_range_values(
        worksheet,
        min_row=max(1, start_row),
        max_row=max(1, end_row),
        min_col=1,
        max_col=max_col,
    )
    candidates: list[dict[str, Any]] = []
    for row_offset, values in enumerate(rows, start=max(1, start_row)):
        columns: list[int] = []
        functions: list[str] = []
        for index, value in enumerate(values, start=1):
            if _ui_clean_text(value) != ">>>":
                continue
            function_name = _ui_clean_text(values[index] if index < len(values) else None)
            if function_name:
                columns.append(index)
                functions.append(function_name)
        if columns:
            candidates.append(
                {
                    "row": row_offset,
                    "columns": columns,
                    "functions": functions,
                }
            )
    return candidates


def _geocost_layout_from_ranges(
    worksheet,
    *,
    table_range: str,
    data_row_range: str,
) -> dict[str, Any]:
    """Validate the visible Overall table and its low-cost percentage row."""
    from openpyxl.utils.cell import range_boundaries

    table_min_col, table_min_row, table_max_col, table_max_row = range_boundaries(
        table_range
    )
    data_min_col, data_min_row, data_max_col, data_max_row = range_boundaries(
        data_row_range
    )
    if data_min_row != data_max_row:
        raise ValueError("Geo Cost data row must be exactly one horizontal row.")
    if table_max_row <= table_min_row:
        raise ValueError(
            "Geo Cost table must include its company header row and at least one geography row."
        )
    if data_min_row <= table_max_row:
        raise ValueError(
            "Geo Cost data row must be below the final geography row of the Geo Cost table."
        )

    max_row = max(1, min(int(worksheet.max_row or 1), 5000))
    max_col = max(1, min(int(worksheet.max_column or 1), 2048))

    header_values = _stream_range_values(
        worksheet,
        min_row=table_min_row,
        max_row=table_min_row,
        min_col=1,
        max_col=max_col,
    )[0]
    found: dict[str, int] = {}
    for column, value in enumerate(header_values, start=1):
        kind = _geocost_header_kind(value)
        if kind and kind not in found:
            found[kind] = column
    required = {"region", "high_low", "entity"}
    missing_headers = sorted(required - set(found))
    if missing_headers:
        raise ValueError(
            "Could not infer Geo Cost helper column(s) to the left of the selected table: "
            + ", ".join(missing_headers)
            + "."
        )

    marker_candidates = _geocost_marker_candidates(
        worksheet,
        max_col=max_col,
        start_row=max(1, table_min_row - 48),
        end_row=table_min_row,
    )
    if not marker_candidates:
        raise ValueError(
            "Could not infer the >>> / function marker row above the selected Geo Cost table."
        )

    chosen_marker: dict[str, Any] | None = None
    chosen_block_index = -1
    chosen_score: tuple[int, int, int] | None = None
    for candidate in marker_candidates:
        columns = list(candidate["columns"])
        for index, start_col in enumerate(columns):
            next_col = columns[index + 1] if index + 1 < len(columns) else max_col + 1
            local_country_col = max(1, start_col - 1)
            intersects = not (
                table_max_col < local_country_col or table_min_col >= next_col
            )
            if not intersects:
                continue
            score = (
                1 if table_min_col <= local_country_col <= table_max_col else 0,
                len(columns),
                -abs(table_min_row - int(candidate["row"])),
            )
            if chosen_score is None or score > chosen_score:
                chosen_score = score
                chosen_marker = candidate
                chosen_block_index = index
    if chosen_marker is None:
        raise ValueError(
            "The selected Geo Cost table does not intersect any inferred >>> function block."
        )

    marker_columns = list(chosen_marker["columns"])
    functions = list(chosen_marker["functions"])
    marker_row = int(chosen_marker["row"])
    first_block_start = int(marker_columns[chosen_block_index])
    next_block_start = (
        int(marker_columns[chosen_block_index + 1])
        if chosen_block_index + 1 < len(marker_columns)
        else None
    )
    first_block_scan_end = next_block_start - 2 if next_block_start else max_col
    local_country_col = max(1, first_block_start - 1)

    company_columns = [
        column
        for column in range(first_block_start, first_block_scan_end + 1)
        if _ui_clean_text(header_values[column - 1] if column - 1 < len(header_values) else None)
    ]
    company_columns = [
        column for column in company_columns if table_min_col <= column <= table_max_col
    ]
    if not company_columns:
        raise ValueError(
            "The selected Geo Cost table contains no company headers for the Overall block."
        )
    company_start_col = min(company_columns)
    company_end_col = max(company_columns)
    if not (table_min_col <= local_country_col <= table_max_col):
        raise ValueError(
            "Geo Cost table must begin with the function-local Country/State column immediately "
            "to the left of the first company."
        )
    if table_max_col < company_end_col:
        raise ValueError("Geo Cost table does not include every Overall company column.")

    raw_country_col = int(found["entity"])
    display_country_col = int(found.get("display", raw_country_col))
    entity_header = _ui_compact_text(
        header_values[raw_country_col - 1] if raw_country_col - 1 < len(header_values) else None
    )
    entity_label = "State" if "state" in entity_header else "Country"

    scan_min_col = min(
        int(found["region"]),
        int(found["high_low"]),
        raw_country_col,
        display_country_col,
        local_country_col,
    )
    scan_max_col = max(
        int(found["region"]),
        int(found["high_low"]),
        raw_country_col,
        display_country_col,
        local_country_col,
    )
    data_rows = _stream_range_values(
        worksheet,
        min_row=table_min_row + 1,
        max_row=table_max_row,
        min_col=scan_min_col,
        max_col=scan_max_col,
    )
    valid_data_rows = 0
    for values in data_rows:
        def at(column: int) -> Any:
            index = column - scan_min_col
            return values[index] if 0 <= index < len(values) else None

        bucket = _ui_clean_text(at(int(found["region"])))
        high_low = _ui_clean_text(at(int(found["high_low"]))).casefold()
        entity = (
            _ui_clean_text(at(local_country_col))
            or _ui_clean_text(at(display_country_col))
            or _ui_clean_text(at(raw_country_col))
        )
        has_bucket = bool(bucket) and ("-" in bucket or high_low in {"high", "low"})
        if entity and has_bucket:
            valid_data_rows += 1
    if valid_data_rows < 1:
        raise ValueError(
            "Geo Cost table contains no geography rows aligned to Region and High/Low metadata."
        )

    selected_data_values = _stream_range_values(
        worksheet,
        min_row=data_min_row,
        max_row=data_min_row,
        min_col=data_min_col,
        max_col=data_max_col,
    )[0]
    label_candidates = [
        (data_min_col + offset, _geocost_low_cost_row_score(value))
        for offset, value in enumerate(selected_data_values)
    ]
    label_candidates = [item for item in label_candidates if item[1] > 0]
    if not label_candidates:
        raise ValueError(
            "Geo Cost data row must include a label such as '% of FTEs in Low Cost region', "
            "'FTEs in low-cost locations (%)', or 'Low cost %'."
        )
    low_cost_label_col = max(label_candidates, key=lambda item: item[1])[0]
    missing_company_cols = [
        column
        for column in company_columns
        if not (data_min_col <= column <= data_max_col)
    ]
    if missing_company_cols:
        raise ValueError(
            "Geo Cost data row must include percentage values for every Overall company."
        )
    numeric_values = [
        _geocost_percentage_number(
            selected_data_values[column - data_min_col]
        )
        for column in company_columns
    ]
    if not any(value is not None for value in numeric_values):
        raise ValueError(
            "Geo Cost data row contains no numeric low-cost percentages under the company columns."
        )

    return {
        "marker_row": marker_row,
        "marker_start_col": marker_columns[0],
        "marker_end_col": min(max_col, marker_columns[-1] + 1),
        "header_row": table_min_row,
        "data_start_row": table_min_row + 1,
        "data_end_row": table_max_row,
        "low_cost_row": data_min_row,
        "low_cost_label_col": low_cost_label_col,
        "table_start_col": table_min_col,
        "table_end_col": table_max_col,
        "segment_bucket_col": int(found["region"]),
        "high_low_col": int(found["high_low"]),
        "raw_country_col": raw_country_col,
        "display_country_col": display_country_col,
        "entity_label": entity_label,
        "first_block_start_col": first_block_start,
        "first_block_end_col": min(company_end_col, first_block_scan_end),
        "company_start_col": company_start_col,
        "company_end_col": company_end_col,
        "function_count": len(marker_columns),
        "function_names": functions,
        "marker_columns": marker_columns,
        "data_row_count": valid_data_rows,
    }


@st.cache_data(show_spinner=False, max_entries=24)
def geocost_auto_input_profile(
    path_text: str,
    cache_token: int,
    requested_sheet: str,
) -> dict[str, Any]:
    """Detect the Overall geography table and its low-cost percentage row."""
    from openpyxl.utils import get_column_letter

    result: dict[str, Any] = {
        "ready": False,
        "sheet_name": "",
        "ranges": {},
        "layout": {},
        "function_count": 0,
        "data_row_count": 0,
        "error": "",
    }
    try:
        workbook = _analysis_workbook(path_text, cache_token)
        selected = _match_sheet_name(list(workbook.sheetnames), requested_sheet)
        if selected is None:
            candidates = [
                name
                for name in workbook.sheetnames
                if str(_preview_sheet_role(name) or "").startswith("geocost")
            ]
            selected = candidates[0] if candidates else None
        if selected is None:
            raise ValueError("No Geo Cost worksheet was detected for input auto-detection.")

        worksheet = workbook[selected]
        max_row = max(1, min(int(worksheet.max_row or 1), 5000))
        max_col = max(1, min(int(worksheet.max_column or 1), 2048))
        top_rows = _stream_range_values(
            worksheet,
            min_row=1,
            max_row=min(max_row, 64),
            min_col=1,
            max_col=max_col,
        )
        marker_candidates: list[tuple[int, int, list[int], list[str]]] = []
        for row_number, values in enumerate(top_rows, start=1):
            columns: list[int] = []
            functions: list[str] = []
            for index, value in enumerate(values, start=1):
                if _ui_clean_text(value) != ">>>":
                    continue
                function_name = _ui_clean_text(values[index] if index < len(values) else None)
                if function_name:
                    columns.append(index)
                    functions.append(function_name)
            if columns:
                marker_candidates.append((len(columns), row_number, columns, functions))
        if not marker_candidates:
            raise ValueError(
                "No Geo Cost >>> function-marker row was found near the top of the worksheet."
            )
        _count, marker_row, marker_columns, function_names = max(
            marker_candidates,
            key=lambda item: (item[0], -item[1]),
        )

        best_header: tuple[int, int, dict[str, int]] | None = None
        header_values_matrix = _stream_range_values(
            worksheet,
            min_row=marker_row,
            max_row=min(max_row, marker_row + 18),
            min_col=1,
            max_col=max_col,
        )
        for offset, values in enumerate(header_values_matrix):
            row_number = marker_row + offset
            found: dict[str, int] = {}
            for column, value in enumerate(values, start=1):
                kind = _geocost_header_kind(value)
                if kind and kind not in found:
                    found[kind] = column
            score = sum(1 for key in ("region", "high_low", "entity") if key in found)
            if "display" in found:
                score += 1
            candidate = (score, row_number, found)
            if best_header is None or (score, -abs(row_number - (marker_row + 3))) > (
                best_header[0],
                -abs(best_header[1] - (marker_row + 3)),
            ):
                best_header = candidate
        if best_header is None or best_header[0] < 3:
            raise ValueError(
                "Geo Cost helper/header structure could not be found beneath the >>> marker row."
            )
        _score, header_row, found = best_header

        first_block_start = marker_columns[0]
        second_block_start = marker_columns[1] if len(marker_columns) > 1 else None
        block_scan_end = second_block_start - 2 if second_block_start else max_col
        header_row_values = _stream_range_values(
            worksheet,
            min_row=header_row,
            max_row=header_row,
            min_col=1,
            max_col=max_col,
        )[0]
        company_columns = [
            column
            for column in range(first_block_start, block_scan_end + 1)
            if _ui_clean_text(
                header_row_values[column - 1]
                if column - 1 < len(header_row_values)
                else None
            )
        ]
        if not company_columns:
            raise ValueError("No company header row was found in the Overall Geo Cost block.")
        company_start_col = min(company_columns)
        company_end_col = max(company_columns)
        local_country_col = max(1, first_block_start - 1)

        raw_country_col = int(found["entity"])
        display_country_col = int(found.get("display", raw_country_col))
        scan_min_col = min(
            int(found["region"]),
            int(found["high_low"]),
            raw_country_col,
            display_country_col,
            local_country_col,
        )
        scan_max_col = max(
            int(found["region"]),
            int(found["high_low"]),
            raw_country_col,
            display_country_col,
            local_country_col,
        )
        rows = _stream_range_values(
            worksheet,
            min_row=header_row + 1,
            max_row=max_row,
            min_col=scan_min_col,
            max_col=scan_max_col,
        )
        last_valid_row = 0
        blank_streak = 0
        started = False
        for row_number, values in enumerate(rows, start=header_row + 1):
            def at(column: int) -> Any:
                index = column - scan_min_col
                return values[index] if 0 <= index < len(values) else None

            bucket = _ui_clean_text(at(int(found["region"])))
            high_low = _ui_clean_text(at(int(found["high_low"]))).casefold()
            entity = (
                _ui_clean_text(at(local_country_col))
                or _ui_clean_text(at(display_country_col))
                or _ui_clean_text(at(raw_country_col))
            )
            has_bucket = bool(bucket) and ("-" in bucket or high_low in {"high", "low"})
            if entity and has_bucket:
                last_valid_row = row_number
                blank_streak = 0
                started = True
            elif started:
                blank_streak += 1
                if blank_streak >= 30:
                    break
        if last_valid_row <= header_row:
            raise ValueError(
                "No geography rows were detected beneath the Geo Cost company header row."
            )

        low_cost_scan_start = last_valid_row + 1
        low_cost_scan_end = min(max_row, last_valid_row + 120)
        low_cost_values = _stream_range_values(
            worksheet,
            min_row=low_cost_scan_start,
            max_row=low_cost_scan_end,
            min_col=local_country_col,
            max_col=company_end_col,
        )
        best_low_cost: tuple[int, int, int, int] | None = None
        for row_number, values in enumerate(low_cost_values, start=low_cost_scan_start):
            label_candidates = [
                (local_country_col + offset, _geocost_low_cost_row_score(value))
                for offset, value in enumerate(values)
            ]
            label_candidates = [item for item in label_candidates if item[1] > 0]
            if not label_candidates:
                continue
            label_col, label_score = max(label_candidates, key=lambda item: item[1])
            numeric_count = 0
            for column in company_columns:
                index = column - local_country_col
                value = values[index] if 0 <= index < len(values) else None
                if _geocost_percentage_number(value) is not None:
                    numeric_count += 1
            if numeric_count < 1:
                continue
            candidate = (label_score, numeric_count, -row_number, label_col)
            if best_low_cost is None or candidate > best_low_cost:
                best_low_cost = candidate
        if best_low_cost is None:
            raise ValueError(
                "Could not find the '% of FTEs in Low Cost region' row below the Overall Geo Cost table."
            )
        _label_score, _numeric_count, neg_low_cost_row, low_cost_label_col = best_low_cost
        low_cost_row = -neg_low_cost_row

        table_range = (
            f"{get_column_letter(local_country_col)}{header_row}:"
            f"{get_column_letter(company_end_col)}{last_valid_row}"
        )
        data_row_range = (
            f"{get_column_letter(low_cost_label_col)}{low_cost_row}:"
            f"{get_column_letter(company_end_col)}{low_cost_row}"
        )
        layout = _geocost_layout_from_ranges(
            worksheet,
            table_range=table_range,
            data_row_range=data_row_range,
        )
        result.update(
            {
                "ready": True,
                "sheet_name": selected,
                "ranges": {
                    "geocost_table": table_range,
                    "geocost_data_row": data_row_range,
                },
                "layout": layout,
                "function_count": int(layout.get("function_count") or 0),
                "data_row_count": int(layout.get("data_row_count") or 0),
                "functions": function_names,
                "error": "",
            }
        )
        return result
    except Exception as exc:
        result["error"] = str(exc)
        return result


@st.cache_data(show_spinner=False, max_entries=48)
def geocost_input_profile_from_mappings(
    path_text: str,
    cache_token: int,
    source_signature: str,
    table_sheet: str,
    table_range: str,
    data_row_sheet: str,
    data_row_range: str,
) -> dict[str, Any]:
    """Validate committed Geo Cost ranges and resolve their inferred source layout."""
    result: dict[str, Any] = {
        "ready": False,
        "source_signature": source_signature,
        "sheet_name": table_sheet,
        "layout": {},
        "function_count": 0,
        "data_row_count": 0,
        "error": "",
    }
    try:
        if not table_sheet or not table_range or not data_row_sheet or not data_row_range:
            raise ValueError("Map both Geo Cost table and Geo Cost data row.")
        if table_sheet.casefold() != data_row_sheet.casefold():
            raise ValueError("Geo Cost table and data row must come from the same worksheet.")
        workbook = _analysis_workbook(path_text, cache_token)
        selected = _match_sheet_name(list(workbook.sheetnames), table_sheet)
        if selected is None:
            raise ValueError(f"Geo Cost worksheet {table_sheet!r} was not found.")
        layout = _geocost_layout_from_ranges(
            workbook[selected],
            table_range=table_range,
            data_row_range=data_row_range,
        )
        result.update(
            {
                "ready": True,
                "sheet_name": selected,
                "layout": layout,
                "function_count": int(layout.get("function_count") or 0),
                "data_row_count": int(layout.get("data_row_count") or 0),
                "error": "",
            }
        )
        return result
    except Exception as exc:
        result["error"] = str(exc)
        return result



def _segment_fill_signature(cell: Any, theme_palette: list[str]) -> tuple[str, str]:
    """Return a stable fill signature and display color for one segment cell."""
    fill = getattr(cell, "fill", None)
    if fill is None:
        return "", ""
    fill_type = str(
        getattr(fill, "fill_type", None)
        or getattr(fill, "patternType", None)
        or ""
    ).strip().lower()
    if not fill_type or fill_type in {"none", "null"}:
        return "", ""
    foreground = _openpyxl_color_hex(
        getattr(fill, "fgColor", None),
        theme_palette,
    ) or ""
    background = _openpyxl_color_hex(
        getattr(fill, "bgColor", None),
        theme_palette,
    ) or ""
    color = foreground or background
    if not color:
        return "", ""
    return f"{fill_type}:{foreground}:{background}", color


def _segment_group_records_for_run(
    *,
    start_column: int,
    end_column: int,
    labels: list[tuple[int, str]],
    color: str,
    method: str,
) -> list[dict[str, Any]]:
    """Create one or more segment spans from a styled row run."""
    if not labels:
        return []
    normalized: list[tuple[int, str, str]] = []
    for column, label in labels:
        key = _ui_compact_text(label)
        if key:
            normalized.append((column, label, key))
    if not normalized:
        return []

    unique_keys: list[str] = []
    for _, _, key in normalized:
        if key not in unique_keys:
            unique_keys.append(key)
    if len(unique_keys) == 1:
        return [
            {
                "segment": normalized[0][1],
                "start_column": start_column,
                "end_column": end_column,
                "segment_color": color,
                "detection_method": method,
            }
        ]

    # A shared fill may contain multiple segment anchors. Split at each new
    # distinct label while treating repeated copies of the same label as one
    # visual group (for example repeated PES labels in Center Across Selection).
    anchors: list[tuple[int, str, str]] = []
    previous_key = ""
    for column, label, key in normalized:
        if key != previous_key:
            anchors.append((column, label, key))
            previous_key = key
    result: list[dict[str, Any]] = []
    for index, (column, label, _key) in enumerate(anchors):
        span_start = start_column if index == 0 else column
        span_end = (
            anchors[index + 1][0] - 1
            if index + 1 < len(anchors)
            else end_column
        )
        result.append(
            {
                "segment": label,
                "start_column": span_start,
                "end_column": span_end,
                "segment_color": color,
                "detection_method": method,
            }
        )
    return result


@st.cache_data(show_spinner=False, max_entries=48)
def l1_segment_detection_profile(
    path_text: str,
    cache_token: int,
    company_sheet: str,
    company_range: str,
    revenue_sheet: str,
    revenue_range: str,
) -> dict[str, Any]:
    """Detect the segment row immediately above the selected company row.

    Detection is structural and style-aware. True merged ranges are honored
    first; otherwise contiguous fill-color runs and Center Across Selection
    spans are used. The first company remains the client and is never assigned
    to a peer segment.
    """
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    result: dict[str, Any] = {
        "detected": False,
        "sheet_name": company_sheet,
        "company_range": company_range,
        "company_row": 0,
        "segment_row": 0,
        "segment_range": "",
        "company_count": 0,
        "segment_count": 0,
        "segment_order": [],
        "segments": [],
        "assignments": [],
        "error": "",
    }
    try:
        if not company_sheet or not company_range:
            raise ValueError("Company names must be mapped before segments can be detected.")
        min_col, min_row, max_col, max_row = range_boundaries(company_range)
        if min_row != max_row:
            raise ValueError("Company names must be a single row for segment detection.")
        if min_row <= 1:
            raise ValueError("The company row has no worksheet row above it for segments.")
        company_count = max_col - min_col + 1
        if company_count < 2:
            raise ValueError("Segment sorting requires the client plus at least one peer.")

        result["company_row"] = min_row
        result["segment_row"] = min_row - 1
        result["company_count"] = company_count
        result["segment_range"] = (
            f"{get_column_letter(min_col)}{min_row - 1}:"
            f"{get_column_letter(max_col)}{min_row - 1}"
        )

        workbook = _analysis_workbook(path_text, cache_token)
        try:
            if company_sheet not in workbook.sheetnames:
                raise ValueError(
                    f"Company names refers to missing worksheet {company_sheet!r}."
                )
            worksheet = workbook[company_sheet]
            theme_palette = _workbook_theme_palette(workbook)
            company_cells = [
                worksheet.cell(row=min_row, column=column)
                for column in range(min_col, max_col + 1)
            ]
            company_names = [
                _ui_clean_text(getattr(cell, "value", None))
                for cell in company_cells
            ]
            if any(not value for value in company_names):
                raise ValueError("The selected Company names row contains blank cells.")

            segment_row = min_row - 1
            segment_cells = [
                worksheet.cell(row=segment_row, column=column)
                for column in range(min_col, max_col + 1)
            ]

            inventory = workbook_sheet_inventory(path_text, cache_token)
            inventory_item = next(
                (item for item in inventory if item.get("name") == company_sheet),
                None,
            )
            if inventory_item is None:
                raise ValueError("The company worksheet could not be resolved in the workbook package.")
            sheet_part = str(inventory_item.get("sheet_part") or "")
            metadata = worksheet_render_metadata(path_text, cache_token, sheet_part)

            groups: list[dict[str, Any]] = []
            occupied: set[int] = set()

            # 1. True Excel merges.
            for reference in metadata.get("merges", []):
                try:
                    merge_min_col, merge_min_row, merge_max_col, merge_max_row = range_boundaries(reference)
                except Exception:
                    continue
                if not (merge_min_row <= segment_row <= merge_max_row):
                    continue
                start = max(min_col, merge_min_col)
                end = min(max_col, merge_max_col)
                if start > end:
                    continue
                anchor_value = _ui_clean_text(
                    worksheet.cell(row=merge_min_row, column=merge_min_col).value
                )
                if not anchor_value:
                    continue
                anchor_cell = worksheet.cell(row=merge_min_row, column=merge_min_col)
                _signature, color = _segment_fill_signature(anchor_cell, theme_palette)
                groups.append(
                    {
                        "segment": anchor_value,
                        "start_column": start,
                        "end_column": end,
                        "segment_color": color,
                        "detection_method": "merged_range",
                    }
                )
                occupied.update(range(start, end + 1))

            # 2. Contiguous styled runs. This is authoritative for the standard
            # Aura segment row and also handles Center Across Selection.
            column = min_col
            while column <= max_col:
                if column in occupied:
                    column += 1
                    continue
                cell = segment_cells[column - min_col]
                signature, color = _segment_fill_signature(cell, theme_palette)
                if not signature:
                    column += 1
                    continue
                run_start = column
                run_end = column
                while run_end + 1 <= max_col and run_end + 1 not in occupied:
                    next_cell = segment_cells[run_end + 1 - min_col]
                    next_signature, _next_color = _segment_fill_signature(
                        next_cell,
                        theme_palette,
                    )
                    if next_signature != signature:
                        break
                    run_end += 1
                labels = [
                    (
                        candidate_column,
                        _ui_clean_text(
                            getattr(
                                segment_cells[candidate_column - min_col],
                                "value",
                                None,
                            )
                        ),
                    )
                    for candidate_column in range(run_start, run_end + 1)
                    if _ui_clean_text(
                        getattr(
                            segment_cells[candidate_column - min_col],
                            "value",
                            None,
                        )
                    )
                ]
                detected_groups = _segment_group_records_for_run(
                    start_column=run_start,
                    end_column=run_end,
                    labels=labels,
                    color=color,
                    method="fill_run",
                )
                for group in detected_groups:
                    groups.append(group)
                    occupied.update(
                        range(
                            int(group["start_column"]),
                            int(group["end_column"]) + 1,
                        )
                    )
                column = run_end + 1

            # 3. Center Across Selection or repeated text without a useful fill.
            column_numbers = list(range(min_col, max_col + 1))
            merge_anchors: dict[tuple[int, int], tuple[int, int]] = {}
            merge_covered: set[tuple[int, int]] = set()
            center_anchors, _center_covered = _center_across_spans_for_row(
                row_number=segment_row,
                column_numbers=column_numbers,
                row_cells=segment_cells,
                actual_merge_anchors=merge_anchors,
                actual_merge_covered=merge_covered,
            )
            for (row_number, start), (_rowspan, colspan) in center_anchors.items():
                if row_number != segment_row:
                    continue
                end = start + colspan - 1
                if any(value in occupied for value in range(start, end + 1)):
                    continue
                label = _ui_clean_text(worksheet.cell(segment_row, start).value)
                if not label:
                    continue
                _signature, color = _segment_fill_signature(
                    worksheet.cell(segment_row, start),
                    theme_palette,
                )
                groups.append(
                    {
                        "segment": label,
                        "start_column": start,
                        "end_column": end,
                        "segment_color": color,
                        "detection_method": "center_across",
                    }
                )
                occupied.update(range(start, end + 1))

            # 4. Final fallback: contiguous repeated labels.
            column = min_col
            while column <= max_col:
                if column in occupied:
                    column += 1
                    continue
                label = _ui_clean_text(worksheet.cell(segment_row, column).value)
                if not label:
                    column += 1
                    continue
                key = _ui_compact_text(label)
                run_end = column
                while run_end + 1 <= max_col and run_end + 1 not in occupied:
                    next_label = _ui_clean_text(
                        worksheet.cell(segment_row, run_end + 1).value
                    )
                    if _ui_compact_text(next_label) != key:
                        break
                    run_end += 1
                _signature, color = _segment_fill_signature(
                    worksheet.cell(segment_row, column),
                    theme_palette,
                )
                groups.append(
                    {
                        "segment": label,
                        "start_column": column,
                        "end_column": run_end,
                        "segment_color": color,
                        "detection_method": "repeated_label",
                    }
                )
                occupied.update(range(column, run_end + 1))
                column = run_end + 1

            groups.sort(key=lambda item: int(item.get("start_column") or 0))
            group_by_column: dict[int, dict[str, Any]] = {}
            for group in groups:
                for group_column in range(
                    int(group["start_column"]),
                    int(group["end_column"]) + 1,
                ):
                    group_by_column.setdefault(group_column, group)

            # Revenue is carried with the assignment so the third rail can show
            # the exact final order before PowerPoint generation.
            revenues: list[float | None] = [None] * company_count
            if revenue_sheet and revenue_range and revenue_sheet in workbook.sheetnames:
                rev_min_col, rev_min_row, rev_max_col, rev_max_row = range_boundaries(
                    revenue_range
                )
                if rev_min_row == rev_max_row:
                    raw_values = [
                        workbook[revenue_sheet].cell(rev_min_row, col).value
                        for col in range(rev_min_col, rev_max_col + 1)
                    ]
                    if len(raw_values) == company_count + 1:
                        raw_values = raw_values[1:]
                    if len(raw_values) == company_count:
                        revenues = [_ui_number(value) for value in raw_values]

            assignments: list[dict[str, Any]] = []
            unassigned: list[str] = []
            for index, (column_number, company_name) in enumerate(
                zip(range(min_col, max_col + 1), company_names)
            ):
                group = group_by_column.get(column_number)
                if index == 0:
                    segment_name = "Client"
                    segment_color = (
                        str((group or {}).get("segment_color") or "")
                        if group
                        else ""
                    )
                elif group is None:
                    segment_name = "Unassigned"
                    segment_color = ""
                    unassigned.append(company_name)
                else:
                    segment_name = _ui_clean_text(group.get("segment")) or "Unassigned"
                    segment_color = str(group.get("segment_color") or "")
                    if segment_name == "Unassigned":
                        unassigned.append(company_name)
                assignments.append(
                    {
                        "company_index": index,
                        "company": company_name,
                        "segment": segment_name,
                        "segment_color": segment_color,
                        "source_column": column_number,
                        "revenue": revenues[index],
                    }
                )

            segment_order: list[str] = []
            for assignment in assignments[1:]:
                segment_name = str(assignment.get("segment") or "").strip()
                if (
                    segment_name
                    and segment_name != "Unassigned"
                    and segment_name not in segment_order
                ):
                    segment_order.append(segment_name)

            result.update(
                {
                    "segments": groups,
                    "assignments": assignments,
                    "segment_order": segment_order,
                    "segment_count": len(segment_order),
                }
            )
            if unassigned:
                raise ValueError(
                    "No segment could be detected for: " + ", ".join(unassigned)
                )
            if not segment_order:
                raise ValueError(
                    "No peer segments were detected in the row above Company names."
                )
            result["detected"] = True
            return result
        finally:
            # Shared cached workbook; callers only read from it.
            pass
    except Exception as exc:
        result["error"] = str(exc)
        return result


def _l1_mapping_ranges_json(mappings: dict[str, dict[str, Any]]) -> str:
    """Return a stable, cache-friendly representation of mapped Aura ranges."""
    payload = {
        key: {
            "sheet_name": str((mappings.get(key) or {}).get("sheet_name") or ""),
            "range_a1": str((mappings.get(key) or {}).get("range_a1") or ""),
        }
        for key in L1_INPUT_KEYS
        if mappings.get(key)
    }
    return json.dumps(payload, sort_keys=True, ensure_ascii=False)


@st.cache_data(show_spinner=False, max_entries=48)
def l1_master_label_profile(
    path_text: str,
    cache_token: int,
    mappings_json: str,
) -> dict[str, Any]:
    """Validate the L1 range contract and the labels used by non-company inputs."""
    from openpyxl.utils.cell import range_boundaries

    try:
        mappings = json.loads(mappings_json or "{}")
    except Exception:
        mappings = {}

    result: dict[str, Any] = {
        "attempted": bool(mappings),
        "mapping_signature": hashlib.sha256(
            repr(
                sorted(
                    (
                        key,
                        str((mappings.get(key) or {}).get("sheet_name") or ""),
                        str((mappings.get(key) or {}).get("range_a1") or ""),
                    )
                    for key in L1_INPUT_KEYS
                    if mappings.get(key)
                )
            ).encode("utf-8")
        ).hexdigest(),
        "ready": False,
        "labels": {},
        "errors": [],
        "company_names": [],
        "fte_company_headers": [],
        "quartile_tables": {},
    }
    if not mappings:
        return result

    try:
        workbook = _analysis_workbook(path_text, cache_token)
    except Exception as exc:
        result["errors"].append(
            f"Workbook could not be opened for master-label validation: {exc}"
        )
        return result

    def read_matrix(key: str, display: str) -> list[list[Any]] | None:
        mapping = mappings.get(key)
        if not isinstance(mapping, dict):
            return None
        sheet_name = str(mapping.get("sheet_name") or "").strip()
        range_a1 = str(mapping.get("range_a1") or "").strip()
        if not sheet_name or not range_a1:
            result["errors"].append(
                f"{display} mapping must contain a worksheet and A1 range."
            )
            return None
        if sheet_name not in workbook.sheetnames:
            result["errors"].append(
                f"{display} refers to missing worksheet {sheet_name!r}."
            )
            return None
        try:
            min_col, min_row, max_col, max_row = range_boundaries(range_a1)
        except Exception:
            result["errors"].append(
                f"{display} has an invalid range: {range_a1!r}."
            )
            return None
        worksheet = workbook[sheet_name]
        return _stream_range_values(
            worksheet,
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col,
        )

    def clean(value: Any) -> str:
        return "" if value is None else str(value).strip()

    try:
        taxonomy = read_matrix("taxonomy", "Taxonomy")
        if taxonomy and taxonomy[0]:
            result["labels"]["taxonomy"] = [
                clean(value) for value in taxonomy[0]
            ]

        companies_matrix = read_matrix("company_headers", "Company names")
        company_names: list[str] = []
        if companies_matrix is not None:
            if len(companies_matrix) != 1 or len(companies_matrix[0]) < 2:
                result["errors"].append(
                    "Company names must be one horizontal value-only row containing "
                    "the client followed by at least one peer company."
                )
            else:
                company_names = [clean(value) for value in companies_matrix[0]]
                if any(not value for value in company_names):
                    result["errors"].append(
                        "Company names contain one or more blank cells."
                    )
                result["company_names"] = company_names

        revenue_matrix = read_matrix("revenue_values", "Revenue row")
        if revenue_matrix is not None:
            if len(revenue_matrix) != 1 or len(revenue_matrix[0]) < 2:
                result["errors"].append(
                    "Revenue row must begin with one Revenue master label followed "
                    "by the company revenue values."
                )
            else:
                revenue_label = clean(revenue_matrix[0][0])
                result["labels"]["revenue_values"] = revenue_label
                if _ui_revenue_label_score(revenue_label) <= 0:
                    result["errors"].append(
                        f"Revenue row begins with {revenue_label!r}; the first selected "
                        "cell must be the Revenue label immediately before the values."
                    )

        fte_matrix = read_matrix("fte_values", "Function + FTE / $B table")
        if fte_matrix is not None:
            if len(fte_matrix) < 2 or not fte_matrix[0]:
                result["errors"].append(
                    "Function + FTE / $B table must include its header row and at "
                    "least one function data row."
                )
            elif company_names:
                header = [clean(value) for value in fte_matrix[0]]
                company_count = len(company_names)
                descriptor_count = len(header) - company_count
                if descriptor_count != 1:
                    result["errors"].append(
                        "Function + FTE / $B must start with exactly one L1/function "
                        "column followed by the company columns. Do not include L0."
                    )
                else:
                    function_header = header[0]
                    fte_company_headers = header[1:]
                    result["labels"]["fte_values"] = function_header
                    result["fte_company_headers"] = fte_company_headers
                    descriptor = _taxonomy_header_descriptor(function_header)
                    if not (
                        (descriptor is not None and descriptor[0] == 1)
                        or _ui_function_header_score(function_header) > 0
                    ):
                        result["errors"].append(
                            f"Function + FTE / $B begins with {function_header!r}; "
                            "the first selected column must be L1/Function."
                        )
                    expected = [_normalize_sheet_key(value) for value in company_names]
                    observed = [
                        _normalize_sheet_key(value) for value in fte_company_headers
                    ]
                    if observed != expected:
                        result["errors"].append(
                            "The company headers inside Function + FTE / $B do not "
                            "match the selected Company names in the same order."
                        )

        def inspect_quartile_table(key: str, display: str, *, require_l0_labels: bool) -> None:
            mapping = mappings.get(key)
            if not isinstance(mapping, dict):
                return
            sheet_name = str(mapping.get("sheet_name") or "").strip()
            range_a1 = str(mapping.get("range_a1") or "").strip()
            profile: dict[str, Any] = {
                "sheet_name": sheet_name,
                "range_a1": range_a1,
                "header_row": 0,
                "headers": {},
                "data_row_count": 0,
                "data_rows": [],
                "row_labels": [],
                "error": "",
            }
            result["quartile_tables"][key] = profile
            if not sheet_name or not range_a1 or sheet_name not in workbook.sheetnames:
                profile["error"] = f"{display} must refer to a valid worksheet range."
                result["errors"].append(profile["error"])
                return
            try:
                min_col, min_row, max_col, max_row = range_boundaries(range_a1)
            except Exception:
                profile["error"] = f"{display} has an invalid range: {range_a1!r}."
                result["errors"].append(profile["error"])
                return
            if max_col - min_col + 1 != 3:
                profile["error"] = (
                    f"{display} must contain exactly three consecutive columns; "
                    f"the selected range contains {max_col - min_col + 1}."
                )
                result["errors"].append(profile["error"])
                return

            worksheet = workbook[sheet_name]
            scan_start_row = max(1, min_row - 80)
            scan_start_col = 1 if require_l0_labels else min_col
            matrix = _stream_range_values(
                worksheet,
                min_row=scan_start_row,
                max_row=max_row,
                min_col=scan_start_col,
                max_col=max_col,
            )

            def value_at(row_number: int, column_number: int) -> Any:
                row_offset = row_number - scan_start_row
                col_offset = column_number - scan_start_col
                if row_offset < 0 or row_offset >= len(matrix):
                    return None
                row_values = matrix[row_offset]
                if col_offset < 0 or col_offset >= len(row_values):
                    return None
                return row_values[col_offset]

            def header_map(row_number: int) -> dict[str, int]:
                found: dict[str, int] = {}
                for column in range(min_col, max_col + 1):
                    kind = _ui_quartile_kind(value_at(row_number, column))
                    if kind and kind not in found:
                        found[kind] = column
                return found

            header_row = 0
            headers = header_map(min_row)
            header_in_selection = {"top", "median", "bottom"}.issubset(headers)
            if header_in_selection:
                header_row = min_row
                data_start_row = min_row + 1
            else:
                data_start_row = min_row
                for row_number in range(min_row - 1, scan_start_row - 1, -1):
                    candidate = header_map(row_number)
                    if {"top", "median", "bottom"}.issubset(candidate):
                        header_row = row_number
                        headers = candidate
                        break
            if not header_row:
                profile["error"] = (
                    f"{display} could not find Top, Median, and Bottom headers in the "
                    "selected three columns or immediately above them."
                )
                result["errors"].append(profile["error"])
                return

            profile["header_row"] = header_row
            profile["header_in_selection"] = header_in_selection
            profile["headers"] = {
                kind: {
                    "column": int(column),
                    "label": clean(value_at(header_row, column)),
                }
                for kind, column in headers.items()
            }
            result["labels"][key] = {
                kind: value["label"] for kind, value in profile["headers"].items()
            }

            data_rows: list[int] = []
            row_labels: list[str] = []
            unlabeled_rows: list[int] = []
            for row_number in range(data_start_row, max_row + 1):
                numeric_count = sum(
                    1
                    for column in range(min_col, max_col + 1)
                    if _ui_number(value_at(row_number, column)) is not None
                )
                if numeric_count < 2:
                    continue
                label = ""
                if require_l0_labels:
                    for column in range(min_col - 1, 0, -1):
                        value = value_at(row_number, column)
                        if _ui_number(value) is None and _ui_looks_like_function(value):
                            candidate = clean(value)
                            if _ui_compact_text(candidate) not in {
                                "total", "overall", "aggregate", "aggregated",
                                "totalexcludingoosandunmapped",
                            }:
                                label = candidate
                                break
                    if not label:
                        unlabeled_rows.append(row_number)
                        continue
                data_rows.append(row_number)
                if label:
                    row_labels.append(label)

            profile["data_rows"] = data_rows
            profile["data_row_count"] = len(data_rows)
            profile["row_labels"] = row_labels
            if not data_rows:
                profile["error"] = f"{display} contains no numeric benchmark rows."
                result["errors"].append(profile["error"])
            elif require_l0_labels and unlabeled_rows:
                profile["ignored_unlabeled_rows"] = unlabeled_rows

        inspect_quartile_table(
            "l0_quartile_table",
            "L0 quartile table detector",
            require_l0_labels=True,
        )
        inspect_quartile_table(
            "l1_quartile_table",
            "L1 quartile table detector",
            require_l0_labels=False,
        )

        result["ready"] = not result["errors"]
        return result
    finally:
        # Shared cached workbook; callers only read from it.
        pass


def _current_taxonomy_sheet_request(source_signature: str) -> str:
    selector_key = f"source_taxonomy_{(source_signature or 'workbook')[:12]}"
    option = str(st.session_state.get(selector_key) or "")
    selected = _sheet_name_from_option(option) if option else ""
    return selected or str(
        st.session_state.get("selected_taxonomy_sheet_override") or ""
    ).strip()


@st.cache_data(show_spinner=False, max_entries=24)
def worksheet_render_metadata(
    path_text: str,
    cache_token: int,
    sheet_part: str,
) -> dict[str, Any]:
    """Read only layout metadata for the selected worksheet XML part."""
    del cache_token
    metadata: dict[str, Any] = {
        "show_gridlines": True,
        "default_column_width": 8.43,
        "default_row_height": 15.0,
        "columns": [],
        "rows": {},
        "merges": [],
    }

    with zipfile.ZipFile(path_text, "r") as archive:
        with archive.open(sheet_part) as stream:
            for event, element in ET.iterparse(stream, events=("start", "end")):
                if event != "start":
                    element.clear()
                    continue
                local_name = element.tag.rsplit("}", 1)[-1]
                attrs = element.attrib

                if local_name == "sheetView":
                    metadata["show_gridlines"] = str(
                        attrs.get("showGridLines", "1")
                    ) not in {"0", "false", "False"}
                elif local_name == "sheetFormatPr":
                    try:
                        metadata["default_column_width"] = float(
                            attrs.get("defaultColWidth") or 8.43
                        )
                    except Exception:
                        pass
                    try:
                        metadata["default_row_height"] = float(
                            attrs.get("defaultRowHeight") or 15.0
                        )
                    except Exception:
                        pass
                elif local_name == "col":
                    try:
                        metadata["columns"].append(
                            {
                                "min": int(attrs.get("min") or 1),
                                "max": int(attrs.get("max") or attrs.get("min") or 1),
                                "width": float(
                                    attrs.get("width")
                                    or metadata["default_column_width"]
                                ),
                                "hidden": str(attrs.get("hidden", "0"))
                                in {"1", "true", "True"},
                            }
                        )
                    except Exception:
                        pass
                elif local_name == "row":
                    try:
                        row_number = int(attrs.get("r") or 0)
                    except Exception:
                        row_number = 0
                    if row_number > 0:
                        row_record: dict[str, Any] = {
                            "hidden": str(attrs.get("hidden", "0"))
                            in {"1", "true", "True"}
                        }
                        if attrs.get("ht") is not None:
                            try:
                                row_record["height"] = float(attrs["ht"])
                            except Exception:
                                pass
                        metadata["rows"][row_number] = row_record
                elif local_name == "mergeCell":
                    reference = str(attrs.get("ref") or "").strip()
                    if reference:
                        metadata["merges"].append(reference)

    return metadata


def _normalize_rgb_hex(value: Any) -> str | None:
    if value is None:
        return None
    raw = str(value).strip().replace("#", "").upper()
    if len(raw) == 8:
        # openpyxl commonly serializes six-digit Excel colors as 00RRGGBB.
        # Treat only the all-zero sentinel as empty; otherwise retain RRGGBB.
        if raw == "00000000":
            return None
        raw = raw[2:]
    if len(raw) == 6 and re.fullmatch(r"[0-9A-F]{6}", raw):
        return raw
    return None


def _apply_excel_tint(hex_color: str, tint: float | None) -> str:
    """Approximate Excel theme tinting without opening Excel COM."""
    tint_value = float(tint or 0.0)
    if abs(tint_value) < 1e-9:
        return hex_color

    channels = [int(hex_color[index:index + 2], 16) for index in (0, 2, 4)]
    adjusted: list[int] = []
    for channel in channels:
        if tint_value < 0:
            value = channel * (1.0 + tint_value)
        else:
            value = channel + (255.0 - channel) * tint_value
        adjusted.append(max(0, min(255, int(round(value)))))
    return "".join(f"{channel:02X}" for channel in adjusted)


def _workbook_theme_palette(workbook) -> list[str]:
    """Return Excel theme colors in the index order used by cell styles."""
    defaults = [
        "FFFFFF",  # Light 1
        "000000",  # Dark 1
        "E7E6E6",  # Light 2
        "44546A",  # Dark 2
        "4472C4",  # Accent 1
        "ED7D31",  # Accent 2
        "A5A5A5",  # Accent 3
        "FFC000",  # Accent 4
        "5B9BD5",  # Accent 5
        "70AD47",  # Accent 6
        "0563C1",  # Hyperlink
        "954F72",  # Followed hyperlink
    ]
    loaded_theme = getattr(workbook, "loaded_theme", None)
    if not loaded_theme:
        return defaults

    try:
        root = ET.fromstring(loaded_theme)
        namespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
        scheme = root.find(f".//{{{namespace}}}clrScheme")
        if scheme is None:
            return defaults

        found: dict[str, str] = {}
        for scheme_color in list(scheme):
            tag = scheme_color.tag.rsplit("}", 1)[-1]
            if not list(scheme_color):
                continue
            color_node = list(scheme_color)[0]
            color_value = (
                color_node.attrib.get("lastClr")
                or color_node.attrib.get("val")
                or ""
            )
            normalized = _normalize_rgb_hex(color_value)
            if normalized:
                found[tag] = normalized

        # Excel's zero-based theme indexes place light colors before dark
        # colors even though the underlying OOXML color scheme is stored in a
        # different element order.
        tags = (
            "lt1",
            "dk1",
            "lt2",
            "dk2",
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hlink",
            "folHlink",
        )
        return [found.get(tag, defaults[index]) for index, tag in enumerate(tags)]
    except Exception:
        return defaults


def _openpyxl_color_hex(
    color: Any,
    theme_palette: list[str],
    default: str | None = None,
) -> str | None:
    if color is None:
        return default

    color_type = str(getattr(color, "type", "") or "").lower()
    result: str | None = None

    if color_type == "rgb":
        raw = getattr(color, "rgb", None)
        if isinstance(raw, str):
            result = _normalize_rgb_hex(raw)
    elif color_type == "theme":
        try:
            theme_index = int(getattr(color, "theme"))
            if 0 <= theme_index < len(theme_palette):
                result = theme_palette[theme_index]
        except Exception:
            result = None
    elif color_type == "indexed":
        try:
            from openpyxl.styles.colors import COLOR_INDEX

            indexed = int(getattr(color, "indexed"))
            if 0 <= indexed < len(COLOR_INDEX):
                result = _normalize_rgb_hex(COLOR_INDEX[indexed])
        except Exception:
            result = None
    elif color_type == "auto":
        result = default

    if result is None:
        return default

    try:
        tint = float(getattr(color, "tint", 0.0) or 0.0)
    except Exception:
        tint = 0.0
    return _apply_excel_tint(result, tint)


def _split_excel_format_sections(format_code: str) -> list[str]:
    sections: list[str] = []
    current: list[str] = []
    quoted = False
    escaped = False
    for character in str(format_code or "General"):
        if escaped:
            current.append(character)
            escaped = False
            continue
        if character == "\\":
            current.append(character)
            escaped = True
            continue
        if character == '"':
            quoted = not quoted
            current.append(character)
            continue
        if character == ";" and not quoted:
            sections.append("".join(current))
            current = []
            continue
        current.append(character)
    sections.append("".join(current))
    return sections


def _number_format_section(format_code: str, value: float) -> str:
    sections = _split_excel_format_sections(format_code)
    if len(sections) == 1:
        return sections[0]
    if value > 0:
        return sections[0]
    if value < 0:
        return sections[1] if len(sections) > 1 else sections[0]
    return sections[2] if len(sections) > 2 else sections[0]


def _decimal_places_from_format(format_code: str) -> int:
    cleaned = re.sub(r'"[^"]*"', "", format_code)
    match = re.search(r"[0#?]+\.([0#?]+)", cleaned)
    return len(match.group(1)) if match else 0


def _format_excel_datetime(value: Any, format_code: str) -> str:
    lower = str(format_code or "").lower()
    has_date = any(token in lower for token in ("yy", "dd", "mmm"))
    has_time = any(token in lower for token in ("h:", "hh", ":mm", "ss"))

    if isinstance(value, datetime):
        if has_date and has_time:
            if "mmm" in lower:
                return value.strftime("%d-%b-%Y %H:%M")
            return value.strftime("%m/%d/%Y %H:%M")
        if has_time and not has_date:
            return value.strftime("%H:%M:%S" if "ss" in lower else "%H:%M")
        if "mmm" in lower:
            return value.strftime("%d-%b-%Y")
        return value.strftime("%m/%d/%Y")
    if isinstance(value, date):
        return value.strftime("%d-%b-%Y" if "mmm" in lower else "%m/%d/%Y")
    if isinstance(value, datetime_time):
        return value.strftime("%H:%M:%S" if "ss" in lower else "%H:%M")
    return str(value)


def _format_excel_number(value: float, format_code: str) -> str:
    if not isinstance(value, (int, float)):
        return str(value)

    section = _number_format_section(format_code or "General", float(value))
    cleaned = re.sub(r"\[[^\]]+\]", "", section)
    cleaned = cleaned.replace("_", "").replace("*", "")
    literal_parts = re.findall(r'"([^"]*)"', cleaned)
    unquoted = re.sub(r'"[^"]*"', "", cleaned)

    if float(value) == 0 and unquoted.strip() in {"-", "--"}:
        return "-"

    if "%" in unquoted:
        decimals = _decimal_places_from_format(unquoted)
        return f"{float(value) * 100:,.{decimals}f}%"

    if re.search(r"[Ee][+-]?0+", unquoted):
        decimals = _decimal_places_from_format(unquoted)
        return f"{float(value):.{decimals}E}"

    decimals = _decimal_places_from_format(unquoted)
    use_grouping = "," in unquoted.split(".", 1)[0]

    # Excel uses trailing commas after a numeric placeholder to scale by 1,000.
    scaling_commas = len(re.findall(r",(?=\s*$)", unquoted.strip()))
    scaled_value = float(value) / (1000 ** scaling_commas) if scaling_commas else float(value)

    absolute_value = abs(scaled_value)
    if use_grouping:
        number_text = f"{absolute_value:,.{decimals}f}"
    elif decimals > 0:
        number_text = f"{absolute_value:.{decimals}f}"
    elif isinstance(value, int) or float(value).is_integer():
        number_text = str(int(round(absolute_value)))
    else:
        number_text = f"{absolute_value:.12g}"

    currency_match = re.search(r"(?:\[\$([^\]-]+)[^\]]*\]|([$€£¥₹]))", cleaned)
    currency = ""
    if currency_match:
        currency = (currency_match.group(1) or currency_match.group(2) or "").strip()
        if currency and len(currency) > 4:
            currency = currency[:4]

    negative = float(value) < 0
    parentheses = negative and "(" in section and ")" in section
    sign = "" if not negative or parentheses else "-"

    suffix = "".join(literal_parts)
    # Common suffixes that are not quoted, such as 0.0x.
    suffix_match = re.search(r"[0#?](?:\.[0#?]+)?\s*([A-Za-z]+)$", unquoted)
    if suffix_match:
        suffix += suffix_match.group(1)

    rendered = f"{sign}{currency}{number_text}{suffix}"
    if parentheses:
        rendered = f"({currency}{number_text}{suffix})"
    return rendered


def _excel_display_text(value: Any, number_format: str) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date, datetime_time)):
        return _format_excel_datetime(value, number_format)
    if isinstance(value, (int, float)):
        return _format_excel_number(value, number_format)
    return str(value)


def _css_border_from_side(
    side: Any,
    theme_palette: list[str],
) -> str | None:
    style_name = str(getattr(side, "style", "") or "")
    if not style_name:
        return None

    width, line_style = {
        "hair": ("1px", "solid"),
        "thin": ("1px", "solid"),
        "medium": ("2px", "solid"),
        "thick": ("3px", "solid"),
        "double": ("3px", "double"),
        "dotted": ("1px", "dotted"),
        "dashed": ("1px", "dashed"),
        "dashDot": ("1px", "dashed"),
        "dashDotDot": ("1px", "dashed"),
        "mediumDashed": ("2px", "dashed"),
        "mediumDashDot": ("2px", "dashed"),
        "mediumDashDotDot": ("2px", "dashed"),
        "slantDashDot": ("2px", "dashed"),
    }.get(style_name, ("1px", "solid"))
    color = _openpyxl_color_hex(
        getattr(side, "color", None),
        theme_palette,
        default="000000",
    ) or "000000"
    return f"{width} {line_style} #{color}"


def _cell_fill_css(cell: Any, theme_palette: list[str]) -> list[str]:
    fill = cell.fill
    pattern_type = str(getattr(fill, "patternType", "") or "")
    if not pattern_type:
        return []

    foreground = _openpyxl_color_hex(
        getattr(fill, "fgColor", None),
        theme_palette,
        default=None,
    )
    background = _openpyxl_color_hex(
        getattr(fill, "bgColor", None),
        theme_palette,
        default="FFFFFF",
    ) or "FFFFFF"

    if pattern_type == "solid":
        if foreground:
            return [f"background-color:#{foreground}"]
        return []

    foreground = foreground or "B8BDC5"
    result = [f"background-color:#{background}"]
    if "horizontal" in pattern_type.lower():
        result.append(
            f"background-image:repeating-linear-gradient(0deg,#{foreground} 0 1px,transparent 1px 4px)"
        )
    elif "vertical" in pattern_type.lower():
        result.append(
            f"background-image:repeating-linear-gradient(90deg,#{foreground} 0 1px,transparent 1px 4px)"
        )
    elif "grid" in pattern_type.lower() or "trellis" in pattern_type.lower():
        result.append(
            "background-image:"
            f"repeating-linear-gradient(0deg,#{foreground} 0 1px,transparent 1px 4px),"
            f"repeating-linear-gradient(90deg,#{foreground} 0 1px,transparent 1px 4px)"
        )
    elif "down" in pattern_type.lower():
        result.append(
            f"background-image:repeating-linear-gradient(45deg,#{foreground} 0 1px,transparent 1px 5px)"
        )
    else:
        result.append(
            f"background-image:repeating-linear-gradient(135deg,#{foreground} 0 1px,transparent 1px 5px)"
        )
    return result


def _cell_style_css(
    cell: Any,
    value: Any,
    theme_palette: list[str],
) -> str:
    css: list[str] = []
    css.extend(_cell_fill_css(cell, theme_palette))

    font = cell.font
    font_color = _openpyxl_color_hex(
        getattr(font, "color", None),
        theme_palette,
        default="000000",
    ) or "000000"
    css.append(f"color:#{font_color}")
    if getattr(font, "bold", False):
        css.append("font-weight:700")
    if getattr(font, "italic", False):
        css.append("font-style:italic")

    decorations: list[str] = []
    if getattr(font, "underline", None):
        decorations.append("underline")
    if getattr(font, "strike", False):
        decorations.append("line-through")
    if decorations:
        css.append(f"text-decoration:{' '.join(decorations)}")

    font_name = str(getattr(font, "name", "") or "Calibri").replace('"', "")
    css.append(f'font-family:"{font_name}",Calibri,"Segoe UI",sans-serif')
    try:
        font_size = float(getattr(font, "sz", 11.0) or 11.0)
    except Exception:
        font_size = 11.0
    css.append(f"font-size:{max(7.0, min(24.0, font_size)):.1f}pt")

    alignment = cell.alignment
    horizontal = str(getattr(alignment, "horizontal", "") or "").lower()
    if horizontal in {"center", "centercontinuous", "distributed"}:
        css.append("text-align:center")
    elif horizontal in {"right"}:
        css.append("text-align:right")
    elif horizontal in {"justify", "fill"}:
        css.append("text-align:justify")
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        css.append("text-align:right")
    else:
        css.append("text-align:left")

    vertical = str(getattr(alignment, "vertical", "") or "").lower()
    css.append(
        "vertical-align:top"
        if vertical == "top"
        else "vertical-align:middle"
        if vertical in {"center", "distributed", "justify"}
        else "vertical-align:bottom"
    )

    if bool(getattr(alignment, "wrap_text", False)):
        css.extend(["white-space:pre-wrap", "overflow-wrap:anywhere"])
    else:
        css.extend(["white-space:nowrap", "overflow:hidden", "text-overflow:clip"])

    try:
        indent = int(getattr(alignment, "indent", 0) or 0)
    except Exception:
        indent = 0
    if indent > 0:
        css.append(f"padding-left:{6 + indent * 10}px")

    try:
        rotation = int(getattr(alignment, "textRotation", 0) or 0)
    except Exception:
        rotation = 0
    if rotation in {90, 255}:
        css.append("writing-mode:vertical-rl")

    border = cell.border
    for edge in ("top", "right", "bottom", "left"):
        side = getattr(border, edge, None)
        css_border = _css_border_from_side(side, theme_palette)
        if css_border:
            css.append(f"border-{edge}:{css_border}")

    return ";".join(css)



def _column_width_pixels_from_metadata(
    metadata: dict[str, Any],
    column_number: int,
) -> int:
    default_width = float(metadata.get("default_column_width") or 8.43)
    width = default_width
    hidden = False
    for record in metadata.get("columns", []):
        if int(record.get("min", 1)) <= column_number <= int(record.get("max", 1)):
            width = float(record.get("width") or default_width)
            hidden = bool(record.get("hidden", False))
    if hidden:
        return 0
    return max(28, min(360, int(round(width * 7.0 + 12))))


def _row_height_pixels_from_metadata(
    metadata: dict[str, Any],
    row_number: int,
) -> int:
    record = dict(metadata.get("rows", {}).get(row_number, {}) or {})
    if bool(record.get("hidden", False)):
        return 0
    height = float(record.get("height") or metadata.get("default_row_height") or 15.0)
    return max(18, min(240, int(round(height * 96.0 / 72.0))))


def _blank_preview_cell(coordinate: str):
    from types import SimpleNamespace
    from openpyxl.styles import Alignment, Border, Font, PatternFill

    return SimpleNamespace(
        coordinate=coordinate,
        value=None,
        number_format="General",
        fill=PatternFill(),
        font=Font(name="Calibri", size=11),
        border=Border(),
        alignment=Alignment(),
        comment=None,
        hyperlink=None,
    )


def _cell_has_display_value(cell: Any) -> bool:
    """Return whether a preview cell contains a value Excel can center."""
    value = getattr(cell, "value", None)
    if value is None:
        return False
    if isinstance(value, str):
        return value != ""
    return True


def _is_center_across_cell(cell: Any) -> bool:
    """Return whether the cell uses Excel's Center Across Selection alignment."""
    alignment = getattr(cell, "alignment", None)
    horizontal = str(getattr(alignment, "horizontal", "") or "").lower()
    return horizontal == "centercontinuous"


def _center_across_spans_for_row(
    *,
    row_number: int,
    column_numbers: list[int],
    row_cells: list[Any],
    actual_merge_anchors: dict[tuple[int, int], tuple[int, int]],
    actual_merge_covered: set[tuple[int, int]],
) -> tuple[dict[tuple[int, int], tuple[int, int]], set[tuple[int, int]]]:
    """Infer Excel ``Center Across Selection`` groups for one rendered row.

    Excel's OOXML does not encode these as ``mergeCell`` ranges. Instead, the
    anchor cell and the blank cells to its right all use horizontal alignment
    ``centerContinuous``. Excel paints the anchor value across that run. The
    preview converts the run to a visual colspan so it matches Excel while the
    underlying workbook remains untouched.
    """
    anchors: dict[tuple[int, int], tuple[int, int]] = {}
    covered: set[tuple[int, int]] = set()
    if not column_numbers:
        return anchors, covered

    def cell_at(offset: int) -> Any:
        column_number = column_numbers[offset]
        raw_cell = row_cells[offset] if offset < len(row_cells) else None
        return (
            raw_cell
            if raw_cell is not None and hasattr(raw_cell, "alignment")
            else _blank_preview_cell(f"{column_number}:{row_number}")
        )

    offset = 0
    while offset < len(column_numbers):
        column_number = column_numbers[offset]
        coordinate_key = (row_number, column_number)

        # Real Excel merges always win. Never build a synthetic span through
        # any coordinate already owned by a mergeCell range.
        if (
            coordinate_key in actual_merge_anchors
            or coordinate_key in actual_merge_covered
        ):
            offset += 1
            continue

        anchor_cell = cell_at(offset)
        if not (_is_center_across_cell(anchor_cell) and _cell_has_display_value(anchor_cell)):
            offset += 1
            continue

        end_offset = offset
        scan_offset = offset + 1
        while scan_offset < len(column_numbers):
            scan_column = column_numbers[scan_offset]
            scan_key = (row_number, scan_column)
            if scan_key in actual_merge_anchors or scan_key in actual_merge_covered:
                break

            scan_cell = cell_at(scan_offset)
            if not _is_center_across_cell(scan_cell):
                break

            # A second populated centerContinuous cell begins a new Center
            # Across Selection group; it is not part of the current one.
            if _cell_has_display_value(scan_cell):
                break

            end_offset = scan_offset
            scan_offset += 1

        if end_offset > offset:
            colspan = end_offset - offset + 1
            anchors[coordinate_key] = (1, colspan)
            for covered_offset in range(offset + 1, end_offset + 1):
                covered.add((row_number, column_numbers[covered_offset]))
            offset = end_offset + 1
        else:
            offset += 1

    return anchors, covered


@st.cache_data(show_spinner=False, max_entries=24)
def workbook_sheet_formatted_preview(
    path_text: str,
    cache_token: int,
    sheet_name: str,
    sheet_part: str,
    start_row: int,
    row_count: int,
    start_column: int,
    column_count: int,
) -> dict[str, Any]:
    """Stream one relevant worksheet range with cached values and cell styles."""
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    layout = worksheet_render_metadata(path_text, cache_token, sheet_part)
    workbook = _analysis_workbook(path_text, cache_token)
    try:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Worksheet not found: {sheet_name}")

        worksheet = workbook[sheet_name]
        worksheet_max_row = max(1, min(EXCEL_MAX_ROWS, int(worksheet.max_row or 1)))
        worksheet_max_column = max(
            1,
            min(EXCEL_MAX_COLUMNS, int(worksheet.max_column or 1)),
        )
        start_row = min(max(1, int(start_row)), EXCEL_MAX_ROWS)
        start_column = min(max(1, int(start_column)), worksheet_max_column)
        end_row = min(
            EXCEL_MAX_ROWS,
            start_row + max(1, int(row_count)) - 1,
        )
        end_column = min(
            worksheet_max_column,
            start_column + max(1, int(column_count)) - 1,
        )

        theme_palette = _workbook_theme_palette(workbook)
        gridlines_visible = bool(layout.get("show_gridlines", True))
        gridline_color = "#d9dde3" if gridlines_visible else "transparent"

        merge_anchors: dict[tuple[int, int], tuple[int, int]] = {}
        merged_covered: set[tuple[int, int]] = set()
        for reference in layout.get("merges", []):
            try:
                min_col, min_row, max_col, max_row_for_merge = range_boundaries(reference)
            except Exception:
                continue
            # A merge is rendered only when its true Excel anchor is visible.
            # This avoids inventing a new anchor when the preview begins in the
            # middle of a merged range.
            if not (
                start_row <= min_row <= end_row
                and start_column <= min_col <= end_column
            ):
                continue
            if max_row_for_merge < start_row or max_col < start_column:
                continue
            clipped_max_row = min(end_row, max_row_for_merge)
            clipped_max_col = min(end_column, max_col)
            anchor = (min_row, min_col)
            merge_anchors[anchor] = (
                clipped_max_row - min_row + 1,
                clipped_max_col - min_col + 1,
            )
            for row_number in range(min_row, clipped_max_row + 1):
                for column_number in range(min_col, clipped_max_col + 1):
                    if (row_number, column_number) != anchor:
                        merged_covered.add((row_number, column_number))

        column_numbers = list(range(start_column, end_column + 1))
        row_numbers = list(range(start_row, end_row + 1))
        column_labels = [get_column_letter(column) for column in column_numbers]

        colgroup = ['<col style="width:52px">']
        for column_number in column_numbers:
            width = _column_width_pixels_from_metadata(layout, column_number)
            if width <= 0:
                colgroup.append('<col style="width:0;visibility:collapse">')
            else:
                colgroup.append(f'<col style="width:{width}px">')

        header_cells = [
            '<th class="corner-cell" title="Select the loaded preview window"></th>'
        ]
        for column_number, column_label in zip(column_numbers, column_labels):
            header_cells.append(
                f'<th class="column-header" data-col="{column_number}" '
                f'title="Select visible column {column_label}">{column_label}</th>'
            )

        body_rows: list[str] = []
        existing_end_row = min(end_row, worksheet_max_row)
        streamed_rows = (
            iter(
                worksheet.iter_rows(
                    min_row=start_row,
                    max_row=existing_end_row,
                    min_col=start_column,
                    max_col=end_column,
                )
            )
            if start_row <= existing_end_row
            else iter(())
        )
        for row_number in row_numbers:
            row_cells = list(
                next(streamed_rows, ()) if row_number <= existing_end_row else ()
            )
            center_across_anchors, center_across_covered = _center_across_spans_for_row(
                row_number=row_number,
                column_numbers=column_numbers,
                row_cells=row_cells,
                actual_merge_anchors=merge_anchors,
                actual_merge_covered=merged_covered,
            )
            row_height = _row_height_pixels_from_metadata(layout, row_number)
            source_hidden = row_height <= 0
            row_style = "display:none" if source_hidden else f"height:{row_height}px"
            cells = [
                f'<th class="row-header" data-row="{row_number}" scope="row" '
                f'title="Select visible row {row_number}">{row_number}</th>'
            ]

            for offset, column_number in enumerate(column_numbers):
                coordinate_key = (row_number, column_number)
                if (
                    coordinate_key in merged_covered
                    or coordinate_key in center_across_covered
                ):
                    continue

                coordinate = f"{get_column_letter(column_number)}{row_number}"
                raw_cell = row_cells[offset] if offset < len(row_cells) else None
                cell = (
                    raw_cell
                    if raw_cell is not None and hasattr(raw_cell, "font")
                    else _blank_preview_cell(coordinate)
                )

                rowspan = 1
                colspan = 1
                if coordinate_key in merge_anchors:
                    rowspan, colspan = merge_anchors[coordinate_key]
                elif coordinate_key in center_across_anchors:
                    rowspan, colspan = center_across_anchors[coordinate_key]

                value = getattr(cell, "value", None)
                number_format = str(getattr(cell, "number_format", "General") or "General")
                display = _excel_display_text(value, number_format)
                escaped_display = html.escape(display).replace("\n", "<br>")
                cell_css = _cell_style_css(cell, value, theme_palette)

                end_coordinate = f"{get_column_letter(column_number + colspan - 1)}{row_number + rowspan - 1}"
                coordinate_label = (
                    coordinate
                    if end_coordinate == coordinate
                    else f"{coordinate}:{end_coordinate}"
                )
                attributes = [
                    f'style="{html.escape(cell_css, quote=True)}"',
                    f'title="{html.escape(coordinate_label, quote=True)}"',
                    f'data-row="{row_number}"',
                    f'data-col="{column_number}"',
                    f'data-row-end="{row_number + rowspan - 1}"',
                    f'data-col-end="{column_number + colspan - 1}"',
                    f'data-coordinate="{html.escape(coordinate_label, quote=True)}"',
                    'role="gridcell"',
                    'aria-selected="false"',
                ]
                if rowspan > 1:
                    attributes.append(f'rowspan="{rowspan}"')
                if colspan > 1:
                    attributes.append(f'colspan="{colspan}"')

                classes = ["excel-cell"]
                if coordinate_key in center_across_anchors:
                    classes.append("center-across-cell")
                if getattr(cell, "comment", None) is not None:
                    classes.append("has-comment")
                cells.append(
                    f'<td class="{" ".join(classes)}" {" ".join(attributes)}>{escaped_display}</td>'
                )

            body_rows.append(
                f'<tr data-row="{row_number}" data-source-hidden="{1 if source_hidden else 0}" '
                f'style="{row_style}">{"".join(cells)}</tr>'
            )

        table_html = f"""
          <style>
            :root {{
              --gridline: {gridline_color};
              --header-bg: #f3f5f8;
              --header-line: #cfd5dc;
              --header-text: #657080;
            }}
            .excel-preview-root,
            .excel-preview-root * {{ box-sizing: border-box; }}
            .excel-preview-root {{
              width: 100%; margin: 0; padding: 0; color: #111827;
              background: transparent;
              font-family: Calibri, "Segoe UI", Arial, sans-serif;
            }}
            .excel-viewport {{
              height: calc(100vh - 100px); min-height: 470px; width: 100%; overflow: auto;
              border: 1px solid #e3e6ea; border-radius: 14px;
              background: #ffffff;
              box-shadow: 0 4px 14px rgba(15, 23, 42, 0.035);
            }}
            .excel-table {{
              border-collapse: separate; border-spacing: 0; table-layout: fixed;
              width: max-content; min-width: 100%; background: #ffffff;
            }}
            .excel-table th, .excel-table td {{
              margin: 0; padding: 2px 6px; line-height: 1.2;
            }}
            .excel-cell {{
              position: relative; border-right: 1px solid var(--gridline);
              border-bottom: 1px solid var(--gridline); background-clip: padding-box;
            }}
            .column-header {{
              position: sticky; top: 0; z-index: 4; height: 32px;
              color: var(--header-text); background: var(--header-bg);
              border-right: 1px solid var(--header-line);
              border-bottom: 1px solid var(--header-line); text-align: center;
              font-family: "Segoe UI", Arial, sans-serif; font-size: 11px; font-weight: 600;
            }}
            .row-header {{
              position: sticky; left: 0; z-index: 3; min-width: 52px; width: 52px;
              color: var(--header-text); background: var(--header-bg);
              border-right: 1px solid var(--header-line);
              border-bottom: 1px solid var(--header-line); text-align: center;
              vertical-align: middle; font-family: "Segoe UI", Arial, sans-serif;
              font-size: 11px; font-weight: 500;
            }}
            .corner-cell {{
              position: sticky; top: 0; left: 0; z-index: 6; min-width: 52px; width: 52px;
              background: #edf0f4; border-right: 1px solid var(--header-line);
              border-bottom: 1px solid var(--header-line);
            }}
            tr.selected-row > .row-header {{
              color: #b00000; background: #fee2e2;
              box-shadow: inset 4px 0 0 #cc0000; font-weight: 700;
            }}
            tr.selected-row > td {{
              box-shadow: inset 0 1px 0 rgba(204,0,0,.45),
                          inset 0 -1px 0 rgba(204,0,0,.45);
            }}
            .has-comment::after {{
              content: ""; position: absolute; top: 0; right: 0; width: 0; height: 0;
              border-top: 7px solid #cc0000; border-left: 7px solid transparent;
            }}
          </style>
          <div class="excel-preview-root">
            <div class="excel-viewport">
              <table class="excel-table" role="grid" aria-label="Formatted Excel preview with mouse range selection">
                <colgroup>{''.join(colgroup)}</colgroup>
                <thead><tr>{''.join(header_cells)}</tr></thead>
                <tbody>{''.join(body_rows)}</tbody>
              </table>
            </div>
          </div>
        """

        return {
            "html": table_html,
            "row_numbers": row_numbers,
            "start_row": start_row,
            "end_row": end_row,
            "start_column": start_column,
            "end_column": end_column,
            "gridlines_visible": gridlines_visible,
        }
    finally:
        # Shared cached workbook; callers only read from it.
        pass


def _excel_column_label(column_number: int) -> str:
    column = max(1, int(column_number))
    label = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        label = chr(65 + remainder) + label
    return label


def _quote_excel_sheet_name(sheet_name: str) -> str:
    return "'" + str(sheet_name).replace("'", "''") + "'"


def _range_selection_store_key(source_signature: str, sheet_name: str) -> str:
    digest = hashlib.sha1(sheet_name.encode("utf-8", errors="replace")).hexdigest()[:12]
    return f"{source_signature}:{digest}"


def _safe_int(value: Any, fallback: int) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return int(fallback)


def _normalize_range_selection(
    raw_value: Any,
    *,
    sheet_name: str,
    source_signature: str,
    max_row: int,
    max_column: int,
) -> dict[str, Any] | None:
    """Validate the browser component result before storing it in session state."""
    if not isinstance(raw_value, dict):
        return None
    if bool(raw_value.get("cleared")) or not str(raw_value.get("range_a1") or ""):
        return None
    if str(raw_value.get("sheet_name") or "") != str(sheet_name):
        return None

    try:
        start_row = int(raw_value["start_row"])
        start_column = int(raw_value["start_column"])
        end_row = int(raw_value["end_row"])
        end_column = int(raw_value["end_column"])
    except (KeyError, TypeError, ValueError):
        return None

    start_row, end_row = sorted((start_row, end_row))
    start_column, end_column = sorted((start_column, end_column))
    if (
        start_row < 1
        or start_column < 1
        or end_row > max(1, int(max_row))
        or end_column > max(1, int(max_column))
    ):
        return None

    first_cell = f"{_excel_column_label(start_column)}{start_row}"
    last_cell = f"{_excel_column_label(end_column)}{end_row}"
    range_a1 = first_cell if first_cell == last_cell else f"{first_cell}:{last_cell}"
    row_count = end_row - start_row + 1
    column_count = end_column - start_column + 1

    try:
        scroll_left = max(0, int(float(raw_value.get("scroll_left") or 0)))
    except (TypeError, ValueError):
        scroll_left = 0
    try:
        scroll_top = max(0, int(float(raw_value.get("scroll_top") or 0)))
    except (TypeError, ValueError):
        scroll_top = 0

    return {
        "version": 1,
        "sheet_name": sheet_name,
        "source_signature": source_signature,
        "range_a1": range_a1,
        "qualified_range": f"{_quote_excel_sheet_name(sheet_name)}!{range_a1}",
        "start_row": start_row,
        "start_column": start_column,
        "end_row": end_row,
        "end_column": end_column,
        "row_count": row_count,
        "column_count": column_count,
        "cell_count": row_count * column_count,
        "scroll_left": scroll_left,
        "scroll_top": scroll_top,
        "anchor_row": max(start_row, min(end_row, _safe_int(raw_value.get("anchor_row"), start_row))),
        "anchor_column": max(
            start_column,
            min(end_column, _safe_int(raw_value.get("anchor_column"), start_column)),
        ),
        "active_row": max(start_row, min(end_row, _safe_int(raw_value.get("active_row"), end_row))),
        "active_column": max(
            start_column,
            min(end_column, _safe_int(raw_value.get("active_column"), end_column)),
        ),
        "zoom_percent": max(40, min(180, _safe_int(raw_value.get("zoom_percent"), 100))),
        "visible_window": dict(raw_value.get("visible_window") or {}),
        "selected_at": str(raw_value.get("selected_at") or ""),
        "cleared": False,
    }



def _mapping_from_a1_range(
    *,
    sheet_name: str,
    range_a1: str,
    source_signature: str,
    mapping_origin: str = "auto",
    detection_label: str = "",
) -> dict[str, Any]:
    """Create the same normalized selection payload used by the range picker."""
    from openpyxl.utils.cell import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    first_cell = f"{_excel_column_label(min_col)}{min_row}"
    last_cell = f"{_excel_column_label(max_col)}{max_row}"
    normalized_range = first_cell if first_cell == last_cell else f"{first_cell}:{last_cell}"
    return {
        "version": 1,
        "sheet_name": sheet_name,
        "source_signature": source_signature,
        "range_a1": normalized_range,
        "qualified_range": f"{_quote_excel_sheet_name(sheet_name)}!{normalized_range}",
        "start_row": min_row,
        "start_column": min_col,
        "end_row": max_row,
        "end_column": max_col,
        "row_count": max_row - min_row + 1,
        "column_count": max_col - min_col + 1,
        "cell_count": (max_row - min_row + 1) * (max_col - min_col + 1),
        "scroll_left": 0,
        "scroll_top": 0,
        "anchor_row": min_row,
        "anchor_column": min_col,
        "active_row": min_row,
        "active_column": min_col,
        "zoom_percent": 100,
        "visible_window": {},
        "selected_at": datetime.now().isoformat(timespec="seconds"),
        "cleared": False,
        "mapping_origin": mapping_origin,
        "detection_label": detection_label,
        "mapping_contract_version": L1_INPUT_MAP_SCHEMA_VERSION,
    }


def _mapped_vector_length(mapping: dict[str, Any] | None) -> int:
    """Return values in the value-only horizontal Company names row."""
    if not isinstance(mapping, dict):
        return 0
    try:
        rows = int(mapping.get("row_count") or 0)
        columns = int(mapping.get("column_count") or 0)
    except (TypeError, ValueError):
        return 0
    if rows == 1 and columns >= 1:
        return columns
    return 0


def _same_mapping_range(left: dict[str, Any] | None, right: dict[str, Any] | None) -> bool:
    if not left or not right:
        return False
    return (
        str(left.get("sheet_name") or "") == str(right.get("sheet_name") or "")
        and str(left.get("range_a1") or "") == str(right.get("range_a1") or "")
        and str(left.get("source_signature") or "") == str(right.get("source_signature") or "")
    )


def _ensure_l1_auto_mappings(
    *,
    workbook_path: Path,
    cache_token: int,
    source_signature: str,
    aura_sheet: str,
    taxonomy_sheet: str,
    geocost_sheet: str = "",
) -> tuple[dict[str, dict[str, Any]], dict[str, Any], bool]:
    """Fill all unmapped standardized Aura and Geo Cost inputs.

    Manual mappings are never overwritten. Auto-detected mappings refresh when
    the selected source worksheet or detected structural range changes.
    """
    current = dict(_get_l1_input_mappings(source_signature))
    detected: dict[str, dict[str, Any]] = {}
    diagnostics: dict[str, Any] = {
        "aura_sheet": aura_sheet,
        "taxonomy_sheet": taxonomy_sheet,
        "geocost_sheet": geocost_sheet,
        "aura_error": "",
        "l0_quartile_error": "",
        "taxonomy_error": "",
        "geocost_error": "",
        "detected_keys": [],
    }

    if taxonomy_sheet or aura_sheet:
        taxonomy_profile = taxonomy_sheet_profile(
            str(workbook_path),
            cache_token,
            taxonomy_sheet,
        )
        diagnostics["taxonomy_profile"] = taxonomy_profile
        diagnostics["taxonomy_error"] = str(taxonomy_profile.get("error") or "")
        if not diagnostics["taxonomy_error"] and taxonomy_profile.get("range_a1"):
            detected["taxonomy"] = _mapping_from_a1_range(
                sheet_name=str(taxonomy_profile.get("sheet_name") or taxonomy_sheet),
                range_a1=str(taxonomy_profile["range_a1"]),
                source_signature=source_signature,
                mapping_origin="auto",
                detection_label="Taxonomy header pair",
            )

        aura_profile = aura_l1_auto_input_profile(
            str(workbook_path),
            cache_token,
            aura_sheet,
        )
        diagnostics["aura_profile"] = aura_profile
        diagnostics["aura_error"] = str(aura_profile.get("error") or "")
        diagnostics["l0_quartile_error"] = str(
            aura_profile.get("l0_quartile_error") or ""
        )
        if not diagnostics["aura_error"]:
            resolved_aura_sheet = str(aura_profile.get("sheet_name") or aura_sheet)
            for key, range_a1 in dict(aura_profile.get("ranges") or {}).items():
                if key not in L1_INPUT_KEYS or key == "taxonomy" or not range_a1:
                    continue
                candidate = _mapping_from_a1_range(
                    sheet_name=resolved_aura_sheet,
                    range_a1=str(range_a1),
                    source_signature=source_signature,
                    mapping_origin="auto",
                    detection_label=(
                        "L0 quartile table"
                        if key == "l0_quartile_table"
                        else (
                            "L1 quartile table"
                            if key == "l1_quartile_table"
                            else "Aura table"
                        )
                    ),
                )
                if key == "fte_values":
                    function_labels_range = str(
                        aura_profile.get("function_labels_range") or ""
                    ).strip()
                    if function_labels_range:
                        candidate["row_label_sheet"] = resolved_aura_sheet
                        candidate["row_label_range"] = function_labels_range
                        candidate["row_label_column"] = int(
                            dict(aura_profile.get("layout") or {}).get(
                                "function_column"
                            )
                            or 0
                        )
                detected[key] = candidate

    geocost_profile = geocost_auto_input_profile(
        str(workbook_path),
        cache_token,
        geocost_sheet,
    )
    diagnostics["geocost_profile"] = geocost_profile
    diagnostics["geocost_error"] = str(geocost_profile.get("error") or "")
    if bool(geocost_profile.get("ready")):
        resolved_geocost_sheet = str(
            geocost_profile.get("sheet_name") or geocost_sheet
        )
        for key, range_a1 in dict(geocost_profile.get("ranges") or {}).items():
            if key not in GEOCOST_INPUT_KEYS or not range_a1:
                continue
            candidate = _mapping_from_a1_range(
                sheet_name=resolved_geocost_sheet,
                range_a1=str(range_a1),
                source_signature=source_signature,
                mapping_origin="auto",
                detection_label=(
                    "Overall Geo Cost table"
                    if key == "geocost_table"
                    else "Geo Cost low-cost FTE percentage row"
                ),
            )
            candidate["mapping_contract_version"] = GEOCOST_INPUT_MAP_SCHEMA_VERSION
            detected[key] = candidate

    changed = False
    for key in STANDARD_INPUT_KEYS:
        existing = current.get(key)
        candidate = detected.get(key)
        existing_origin = str((existing or {}).get("mapping_origin") or "manual")
        if existing and existing_origin != "auto":
            continue
        if candidate is None:
            if existing and existing_origin == "auto":
                current.pop(key, None)
                changed = True
            continue
        metadata_changed = False
        if key == "fte_values":
            metadata_changed = any(
                str((existing or {}).get(field) or "")
                != str(candidate.get(field) or "")
                for field in (
                    "row_label_sheet",
                    "row_label_range",
                    "row_label_column",
                )
            )
        if (
            not _same_mapping_range(existing, candidate)
            or existing_origin != "auto"
            or metadata_changed
        ):
            current[key] = candidate
            changed = True

    diagnostics["detected_keys"] = [
        key for key in STANDARD_INPUT_KEYS if key in detected
    ]
    diagnostics["mapped_count"] = sum(
        1 for key in STANDARD_INPUT_KEYS if current.get(key)
    )
    st.session_state["l1_auto_detection"] = diagnostics
    st.session_state["geocost_auto_detection"] = {
        "error": diagnostics["geocost_error"],
        "profile": geocost_profile,
        "detected_keys": [key for key in GEOCOST_INPUT_KEYS if key in detected],
    }
    if changed:
        _set_l1_input_mappings(source_signature, current)
    return current, diagnostics, changed


def _taxonomy_profile_from_mapping(
    *,
    workbook_path: Path,
    cache_token: int,
    mapping: dict[str, Any] | None,
    fallback_sheet: str,
) -> dict[str, Any]:
    if mapping:
        sheet_name = str(mapping.get("sheet_name") or "").strip()
        range_a1 = str(mapping.get("range_a1") or "").strip()
        if sheet_name and range_a1:
            return taxonomy_range_profile(
                str(workbook_path),
                cache_token,
                sheet_name,
                range_a1,
            )
    return taxonomy_sheet_profile(
        str(workbook_path),
        cache_token,
        fallback_sheet,
    )

def _render_active_input_validation_banner(
    *,
    source_signature: str,
    status: dict[str, Any],
    ignore_mismatch_name: str = "",
    restore_mismatch_name: str = "",
) -> None:
    """Render active-input status and exclusion actions outside the custom component.

    State-changing Ignore/Re-enable actions must remain native Streamlit buttons.
    Custom-component values persist across rerenders, so using them as commands can
    replay the same action and cause recursive reruns.
    """
    level = str(status.get("level") or "info").strip().lower()
    if level not in {"ok", "warning", "error", "info"}:
        level = "info"
    message = str(status.get("message") or "").strip()
    if not message:
        return

    palette = {
        "ok": ("#eaf7ef", "#b9e1c7", "#1f8a4c"),
        "warning": ("#fff8df", "#ead38a", "#d69a00"),
        "error": ("#fff0f0", "#efb6b6", "#c90000"),
        "info": ("#f6f7f9", "#d9dde3", "#59616d"),
    }
    background, border, dot = palette[level]
    action_name = ""
    action_label = ""
    action_kind = ""
    if ignore_mismatch_name:
        action_name = " ".join(str(ignore_mismatch_name).split()).strip()
        action_label = f"Ignore {action_name}"
        action_kind = "ignore"
    elif restore_mismatch_name:
        action_name = " ".join(str(restore_mismatch_name).split()).strip()
        action_label = f"Re-enable {action_name}"
        action_kind = "restore"

    # Keep the status message and its native Ignore/Re-enable action in one
    # visual alert row. The action remains a real Streamlit button (rather
    # than a custom-component command), so it still fires once per click.
    st.markdown(
        f"""
        <style>
        .st-key-active_input_validation_banner {{
            border: 1px solid {border};
            border-radius: 8px;
            background: {background};
            min-height: 36px;
            margin: 0 0 6px 0;
            padding: 4px 8px;
            box-sizing: border-box;
        }}
        .st-key-active_input_validation_banner [data-testid="stHorizontalBlock"] {{
            align-items: center;
            gap: 8px;
        }}
        .st-key-active_input_validation_banner [data-testid="stColumn"] {{
            min-width: 0;
        }}
        .st-key-active_input_validation_banner .stMarkdown {{
            margin: 0 !important;
        }}
        .st-key-active_input_validation_banner .stMarkdown p {{
            margin: 0 !important;
        }}
        .st-key-active_input_validation_banner .stButton {{
            display: flex;
            justify-content: flex-end;
            margin: 0 !important;
        }}
        .st-key-active_input_validation_banner .stButton > button {{
            min-height: 26px !important;
            height: 26px !important;
            padding: 0 10px !important;
            border-radius: 6px !important;
            font-size: 11px !important;
            line-height: 1 !important;
            white-space: nowrap !important;
            margin: 0 !important;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    has_action = bool(action_name and action_kind and source_signature)
    with st.container(key="active_input_validation_banner"):
        if has_action:
            message_col, action_col = st.columns([8.5, 1.5], gap="small")
        else:
            message_col = st.container()
            action_col = None

        with message_col:
            st.markdown(
                f'<div style="display:flex;align-items:center;gap:8px;min-height:26px;'
                f'font-size:12px;line-height:1.3;color:#30343a;box-sizing:border-box;">'
                f'<span style="width:8px;height:8px;border-radius:999px;background:{dot};'
                f'flex:0 0 8px;"></span><span>{html.escape(message)}</span></div>',
                unsafe_allow_html=True,
            )

        if not has_action or action_col is None:
            return

        action_key = hashlib.sha1(
            f"{source_signature}|{action_kind}|{action_name.casefold()}".encode(
                "utf-8", errors="replace"
            )
        ).hexdigest()[:14]
        with action_col:
            action_clicked = st.button(
                action_label,
                key=f"taxonomy_exclusion_{action_key}",
                type="secondary",
                use_container_width=True,
            )
        if action_clicked:
            if action_kind == "ignore":
                _add_ignored_aura_taxonomy_function(source_signature, action_name)
            else:
                _remove_ignored_aura_taxonomy_function(source_signature, action_name)
            # Native Streamlit buttons return True for one pass only, so this
            # single rerun cannot replay recursively.
            st.session_state["workbook_review_finalized"] = False
            st.session_state["canvas_stage"] = "review"
            st.rerun()


def _render_interactive_range_selector(
    *,
    preview: dict[str, Any],
    sheet_name: str,
    source_signature: str,
    max_row: int,
    max_column: int,
    mapping_section: str = "",
    active_input_key: str = "",
    active_input_label: str = "",
    active_input_status: str = "",
    active_input_status_message: str = "",
    active_input_ignore_mismatch: str = "",
    active_input_restore_mismatch: str = "",
) -> dict[str, Any] | None:
    """Render the mouse-selectable grid and optionally bind it to one named input."""
    selection_store = dict(st.session_state.get("workbook_range_selections") or {})
    store_key = _range_selection_store_key(source_signature, sheet_name)
    mapping_mode = (
        mapping_section in {"l1", "aura", "inputs"}
        and active_input_key in STANDARD_INPUT_KEYS
    )
    mappings = _get_l1_input_mappings(source_signature) if mapping_mode else {}

    if mapping_mode:
        mapped_selection = mappings.get(active_input_key)
        initial_selection = (
            mapped_selection
            if isinstance(mapped_selection, dict)
            and str(mapped_selection.get("sheet_name") or "") == sheet_name
            else None
        )
    else:
        initial_selection = selection_store.get(store_key)

    sheet_digest = hashlib.sha1(
        sheet_name.encode("utf-8", errors="replace")
    ).hexdigest()[:12]
    mapping_revision = ""
    if mapping_mode and isinstance(initial_selection, dict):
        mapping_revision = hashlib.sha1(
            repr(
                (
                    str(initial_selection.get("sheet_name") or ""),
                    str(initial_selection.get("range_a1") or ""),
                    int(initial_selection.get("start_row") or 0),
                    int(initial_selection.get("start_column") or 0),
                    int(initial_selection.get("end_row") or 0),
                    int(initial_selection.get("end_column") or 0),
                    str(initial_selection.get("mapping_origin") or ""),
                )
            ).encode("utf-8", errors="replace")
        ).hexdigest()[:10]
    component_key_suffix = (
        f"_{mapping_section}_{active_input_key}_{mapping_revision or 'empty'}"
        if mapping_mode
        else ""
    )

    try:
        selector = _get_range_selector_component()
        raw_selection = selector(
            table_html=preview["html"],
            sheet_name=sheet_name,
            source_signature=source_signature,
            active_input_key=active_input_key if mapping_mode else "",
            active_input_label=active_input_label if mapping_mode else "",
            # Status text and Ignore/Re-enable actions are rendered by native
            # Streamlit controls above the custom grid. Never use a persistent
            # custom-component value as a state-changing command.
            active_input_status="",
            active_input_status_message="",
            active_input_ignore_mismatch="",
            active_input_restore_mismatch="",
            initial_selection=initial_selection or {},
            visible_start_row=int(preview["start_row"]),
            visible_end_row=int(preview["end_row"]),
            visible_start_column=int(preview["start_column"]),
            visible_end_column=int(preview["end_column"]),
            component_height=940,
            key=(
                f"excel_range_selector_{source_signature[:12]}_{sheet_digest}"
                f"{component_key_suffix}"
            ),
            default=initial_selection,
        )
    except Exception as exc:
        components.html(preview["html"], height=900, scrolling=False)
        st.warning(
            "Interactive range selection could not be initialized in this "
            f"Streamlit runtime. The preview remains read-only. Details: {exc}"
        )
        st.session_state["active_range_selection"] = initial_selection
        return initial_selection

    def selection_identity(value: dict[str, Any] | None) -> tuple[Any, ...] | None:
        if not value:
            return None
        return (
            value.get("sheet_name"),
            value.get("range_a1"),
            int(value.get("start_row") or 0),
            int(value.get("start_column") or 0),
            int(value.get("end_row") or 0),
            int(value.get("end_column") or 0),
            int(value.get("scroll_left") or 0),
            int(value.get("scroll_top") or 0),
        )

    initial_identity = selection_identity(initial_selection)
    mapping_changed = False

    if isinstance(raw_selection, dict):
        emitted_input_key = str(raw_selection.get("input_key") or "")
        stale_mapping_event = (
            mapping_mode
            and emitted_input_key
            and emitted_input_key != active_input_key
        )
        if stale_mapping_event:
            selection = initial_selection
        elif bool(raw_selection.get("cleared")) or not str(
            raw_selection.get("range_a1") or ""
        ):
            selection_store.pop(store_key, None)
            selection = None
            if mapping_mode:
                previous = mappings.get(active_input_key)
                if (
                    isinstance(previous, dict)
                    and str(previous.get("sheet_name") or "") == sheet_name
                ):
                    mappings.pop(active_input_key, None)
                    _set_l1_input_mappings(source_signature, mappings)
                    mapping_changed = True
        else:
            selection = _normalize_range_selection(
                raw_selection,
                sheet_name=sheet_name,
                source_signature=source_signature,
                max_row=max_row,
                max_column=max_column,
            )
            if selection is not None:
                selection_store[store_key] = selection
                if mapping_mode:
                    selection["mapping_origin"] = "manual"
                    selection["detection_label"] = ""
                    selection["mapping_contract_version"] = (
                        GEOCOST_INPUT_MAP_SCHEMA_VERSION
                        if active_input_key in GEOCOST_INPUT_KEYS
                        else AURA_INPUT_MAP_SCHEMA_VERSION
                    )
                    previous_identity = selection_identity(
                        mappings.get(active_input_key)
                    )
                    mappings[active_input_key] = selection
                    _set_l1_input_mappings(source_signature, mappings)
                    mapping_changed = (
                        selection_identity(selection) != previous_identity
                    )
    else:
        selection = initial_selection

    st.session_state["workbook_range_selections"] = selection_store
    st.session_state["active_range_selection"] = selection

    rerun = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)
    if (
        selection_identity(selection) != initial_identity or mapping_changed
    ) and callable(rerun):
        rerun()

    return selection



def _selected_section_set(
    selected_sections: Iterable[str] | None,
    workflow_mode: str = "",
) -> set[str]:
    """Normalize the new multi-select while tolerating older callers."""
    valid = {"summary", "l0", "l1", "geocost"}
    selected = {
        str(item)
        for item in (selected_sections or ())
        if str(item) in valid
    }
    if selected:
        return selected
    legacy = {
        "complete": valid,
        "summary": {"summary"},
        "l0": {"l0"},
        "l1": {"l1"},
        "geocost": {"geocost"},
    }
    return set(legacy.get(str(workflow_mode or ""), set()))


def _preferred_preview_sheet(manifest: list[dict[str, Any]]) -> str:
    if not manifest:
        return ""
    return str(manifest[0]["name"])


def _bounded_formatted_preview_window(
    *,
    sheet_name: str,
    sheet_role: str,
    auto_end_row: int,
    last_content_column: int,
    active_mapping: dict[str, Any] | None,
    requested_start_row: int | None = None,
) -> dict[str, int | bool]:
    """Return a responsive worksheet window for the HTML preview.

    The committed input range is authoritative.  When that range is larger
    than the bounded HTML window (common for the horizontal Geo Cost marker
    row), keep the *start* of the range in view rather than jumping to its
    bottom/right edge.  The range selector can then highlight the visible
    portion while its toolbar continues to show the complete A1 reference.
    """
    from openpyxl.utils.cell import range_boundaries

    sheet_rows = max(1, min(EXCEL_MAX_ROWS, int(auto_end_row or 1)))
    sheet_columns = max(
        1, min(EXCEL_MAX_COLUMNS, int(last_content_column or 1))
    )
    target: tuple[int, int, int, int] | None = None
    mapping = active_mapping or {}
    mapping_sheet = str(mapping.get("sheet_name") or "").strip()
    if (
        mapping_sheet
        and mapping_sheet.casefold() == str(sheet_name or "").strip().casefold()
        and str(mapping.get("range_a1") or "").strip()
    ):
        try:
            target = tuple(
                int(value)
                for value in range_boundaries(str(mapping["range_a1"]))
            )
        except Exception:
            target = None

    is_geocost = str(sheet_role or "").startswith("geocost")
    cell_budget = (
        GEOCOST_FORMATTED_PREVIEW_CELLS
        if is_geocost
        else MAX_FORMATTED_PREVIEW_CELLS
    )

    # Normal worksheets keep the established 140-column cap.  A Geo Cost
    # marker/function row can be much wider than that, so when such a one-row
    # mapping is active, temporarily trade some vertical rows for enough
    # columns to display the complete mapped row whenever practical.
    column_limit = MAX_FORMATTED_PREVIEW_COLUMNS
    target_width = 0
    target_height = 0
    if target is not None:
        target_min_col, target_min_row, target_max_col, target_max_row = target
        target_width = target_max_col - target_min_col + 1
        target_height = target_max_row - target_min_row + 1
        if is_geocost and target_height <= 3 and target_width > column_limit:
            wide_row_minimum_rows = 110
            budget_column_limit = max(
                column_limit,
                cell_budget // wide_row_minimum_rows,
            )
            column_limit = min(
                512,
                budget_column_limit,
                max(column_limit, target_width + 8),
            )

    column_count = min(sheet_columns, column_limit)
    start_column = 1
    if sheet_columns > column_count and target is not None:
        target_min_col, _target_min_row, target_max_col, _target_max_row = target
        target_fits = target_width <= column_count
        desired = max(1, target_min_col - 5)
        start_column = min(desired, sheet_columns - column_count + 1)
        # Only shift toward the right edge when the complete target can fit.
        # If it is wider than the preview window, anchoring on the first cell
        # is the only deterministic way to reveal the correct mapped range.
        if target_fits and target_max_col > start_column + column_count - 1:
            start_column = min(
                max(1, target_max_col - column_count + 1),
                sheet_columns - column_count + 1,
            )

    minimum_rows = 220 if is_geocost else 120
    if is_geocost and target is not None and target_height <= 3 and column_count > MAX_FORMATTED_PREVIEW_COLUMNS:
        minimum_rows = 110
    row_budget = max(minimum_rows, cell_budget // max(1, column_count))
    row_limit = min(MAX_FORMATTED_PREVIEW_ROWS, row_budget)
    row_count = min(sheet_rows, row_limit)
    start_row = 1
    if sheet_rows > row_count and target is not None:
        _target_min_col, target_min_row, _target_max_col, target_max_row = target
        target_fits = target_height <= row_count
        desired = max(1, target_min_row - PREVIEW_MAPPING_CONTEXT_ROWS)
        start_row = min(desired, sheet_rows - row_count + 1)
        # As with columns, do not jump to the bottom of an oversized committed
        # mapping.  Keep its first row visible and let the user scroll through
        # the highlighted range naturally.
        if target_fits and target_max_row > start_row + row_count - 1:
            start_row = min(
                max(1, target_max_row - row_count + 1),
                sheet_rows - row_count + 1,
            )
    elif sheet_rows > row_count and requested_start_row is not None:
        start_row = min(
            max(1, int(requested_start_row)),
            sheet_rows - row_count + 1,
        )

    end_row = min(sheet_rows, start_row + row_count - 1)
    end_column = min(sheet_columns, start_column + column_count - 1)
    return {
        "start_row": start_row,
        "row_count": end_row - start_row + 1,
        "end_row": end_row,
        "start_column": start_column,
        "column_count": end_column - start_column + 1,
        "end_column": end_column,
        "limited": bool(
            start_row > 1
            or end_row < sheet_rows
            or start_column > 1
            or end_column < sheet_columns
        ),
        "target_clipped": bool(
            target is not None
            and (
                target[0] < start_column
                or target[2] > end_column
                or target[1] < start_row
                or target[3] > end_row
            )
        ),
    }


def _preview_row_window_key(source_signature: str, sheet_name: str) -> str:
    identity = re.sub(
        r"[^a-z0-9]+",
        "_",
        f"{source_signature}:{sheet_name}".casefold(),
    ).strip("_")
    return f"preview_row_window_{identity}"


def _render_preview_row_navigation(
    *,
    state_key: str,
    start_row: int,
    end_row: int,
    total_rows: int,
    row_count: int,
) -> None:
    """Render compact paging only when a bounded preview cannot show all rows."""
    if total_rows <= row_count:
        return

    max_start = max(1, total_rows - row_count + 1)
    page_step = max(1, row_count - 20)
    previous_start = max(1, start_row - page_step)
    next_start = min(max_start, start_row + page_step)

    label_col, top_col, previous_col, next_col, bottom_col = st.columns(
        [3.3, 1, 1, 1, 1]
    )
    with label_col:
        st.caption(
            f"Preview rows {start_row:,}–{end_row:,} of {total_rows:,}. "
            "Use the row controls to browse the rest of this wide worksheet."
        )

    rerun = getattr(st, "rerun", None) or getattr(st, "experimental_rerun", None)

    def move_to(value: int) -> None:
        st.session_state[state_key] = int(value)
        if callable(rerun):
            rerun()

    with top_col:
        if st.button(
            "Top",
            key=f"{state_key}_top",
            disabled=start_row <= 1,
            use_container_width=True,
        ):
            move_to(1)
    with previous_col:
        if st.button(
            "Previous",
            key=f"{state_key}_previous",
            disabled=start_row <= 1,
            use_container_width=True,
        ):
            move_to(previous_start)
    with next_col:
        if st.button(
            "Next",
            key=f"{state_key}_next",
            disabled=end_row >= total_rows,
            use_container_width=True,
        ):
            move_to(next_start)
    with bottom_col:
        if st.button(
            "Bottom",
            key=f"{state_key}_bottom",
            disabled=end_row >= total_rows,
            use_container_width=True,
        ):
            move_to(max_start)


def render_workbook_preview(
    preview_workbook_path: Path | None,
    workbook_display_name: str,
    preview_metadata: dict[str, Any] | None = None,
    preview_error: str | None = None,
    workflow_mode: str = "complete",
    selected_sections: Iterable[str] | None = None,
    input_section: str = "",
    load_progress: WorkbookLoadProgress | None = None,
) -> None:
    """Render a source-mapped workbook preview with maximum sheet viewport."""
    if load_progress is not None and load_progress.active:
        load_progress.update(
            64,
            "Reading workbook",
            "Building a worksheet inventory and identifying relevant source tabs.",
            phase="detect",
        )
    source_signature = str(st.session_state.get("uploaded_signature") or "")
    section_set = _selected_section_set(selected_sections, workflow_mode)
    required_aura_keys = aura_input_keys_for_sections(tuple(section_set))
    required_geocost_keys = (
        GEOCOST_INPUT_KEYS if "geocost" in section_set else ()
    )
    required_input_keys = standardized_input_keys_for_sections(tuple(section_set))
    show_input_mapping = bool(required_input_keys)
    sort_enabled = bool(section_set.intersection(SORTABLE_DECK_SECTIONS))
    l1_mappings = _get_l1_input_mappings(source_signature)

    preview_available = bool(
        preview_workbook_path is not None
        and preview_workbook_path.is_file()
    )
    token = _workbook_cache_token(preview_workbook_path) if preview_available else 0
    inventory: list[dict[str, Any]] = []
    inventory_error = ""
    if preview_available and preview_workbook_path is not None:
        try:
            with ui_stage(
                "workbook.inventory",
                warn_after=8,
                stack_after=25,
                path=str(preview_workbook_path),
            ):
                inventory = workbook_sheet_inventory(
                    str(preview_workbook_path),
                    token,
                )
            ui_log(
                "WORKBOOK INVENTORY",
                sheets=len(inventory),
                names=[item.get("name") for item in inventory],
            )
            if load_progress is not None and load_progress.active:
                load_progress.update(
                    68,
                    "Worksheets identified",
                    f"Found {len(inventory)} worksheet{'s' if len(inventory) != 1 else ''}. Detecting slide inputs.",
                    phase="detect",
                )
        except Exception as exc:
            inventory_error = str(exc)
            ui_log(
                "WORKBOOK INVENTORY FAILED",
                level="ERROR",
                error=f"{type(exc).__name__}: {exc}",
            )

    # Resolve the current source choices before rendering the controls so every
    # Standardized Aura inputs can be prefilled on the first review pass. Manual mappings are
    # preserved; only prior auto-detected mappings are refreshed.
    actual_aura_sheet = ""
    actual_taxonomy_sheet = ""
    actual_geocost_sheet = ""
    if inventory:
        selected_aura_override = str(
            st.session_state.get("selected_aura_sheet_override") or ""
        ).strip()
        selected_taxonomy_override = str(
            st.session_state.get("selected_taxonomy_sheet_override") or ""
        ).strip()
        selected_geocost_override = str(
            st.session_state.get("selected_geocost_sheet_override") or ""
        ).strip()
        actual_aura_sheet = selected_aura_override or _auto_detected_source_sheet(
            inventory, "aura"
        )
        actual_taxonomy_sheet = (
            selected_taxonomy_override
            or _auto_detected_source_sheet(inventory, "taxonomy")
        )
        actual_geocost_sheet = (
            selected_geocost_override
            or _auto_detected_source_sheet(inventory, "geocost")
        )
        st.session_state["resolved_aura_sheet"] = actual_aura_sheet
        st.session_state["resolved_taxonomy_sheet"] = actual_taxonomy_sheet
        st.session_state["resolved_geocost_sheet"] = actual_geocost_sheet

    if (
        preview_available
        and preview_workbook_path is not None
        and source_signature
        and (show_input_mapping or sort_enabled)
        and not inventory_error
    ):
        if load_progress is not None and load_progress.active:
            load_progress.update(
                71,
                "Detecting slide inputs",
                "Locating Aura and Geo Cost source ranges.",
                phase="detect",
            )
        with ui_stage(
            "workbook.auto_detect_inputs",
            warn_after=12,
            stack_after=35,
            aura_sheet=actual_aura_sheet,
            taxonomy_sheet=actual_taxonomy_sheet,
            geocost_sheet=actual_geocost_sheet,
        ):
            l1_mappings, _auto_diagnostics, auto_changed = _ensure_l1_auto_mappings(
                workbook_path=preview_workbook_path,
                cache_token=token,
                source_signature=source_signature,
                aura_sheet=actual_aura_sheet,
                taxonomy_sheet=actual_taxonomy_sheet,
                geocost_sheet=actual_geocost_sheet,
            )
        ui_log(
            "AUTO INPUTS",
            changed=auto_changed,
            mapped={
                key: str((l1_mappings.get(key) or {}).get("range_a1") or "")
                for key in required_input_keys
            },
            diagnostics={
                "aura_error": str(_auto_diagnostics.get("aura_error") or ""),
                "taxonomy_error": str(_auto_diagnostics.get("taxonomy_error") or ""),
                "l0_quartile_error": str(_auto_diagnostics.get("l0_quartile_error") or ""),
                "geocost_error": str(_auto_diagnostics.get("geocost_error") or ""),
            },
        )
        if load_progress is not None and load_progress.active:
            mapped_count = sum(1 for key in required_input_keys if l1_mappings.get(key))
            load_progress.update(
                77,
                "Inputs detected",
                f"Detected {mapped_count} of {len(required_input_keys)} required input range{'s' if len(required_input_keys) != 1 else ''}.",
                phase="detect",
            )
        if auto_changed:
            # Seed the validated Taxonomy profile before one controlled rerun so
            # the standardized Inputs rail immediately shows its detected ranges.
            # The signature guard prevents a detector/cache edge case from
            # trapping the uploader in a permanent rerun loop.
            detected_taxonomy_profile = _taxonomy_profile_from_mapping(
                workbook_path=preview_workbook_path,
                cache_token=token,
                mapping=l1_mappings.get("taxonomy"),
                fallback_sheet=actual_taxonomy_sheet,
            )
            st.session_state["taxonomy_sheet_profile"] = detected_taxonomy_profile
            st.session_state["resolved_taxonomy_sheet"] = str(
                detected_taxonomy_profile.get("sheet_name") or actual_taxonomy_sheet
            )
            # Do not force the visible source tab here. The active-input
            # navigation block below performs that synchronization once per
            # input change. Keeping it out of auto-detection is what allows a
            # user to browse Aura Summary, Taxonomy, or Geo Cost freely after
            # the mappings have been detected.

            mapping_signature = hashlib.sha256(
                json.dumps(
                    {
                        "source": source_signature,
                        "aura": actual_aura_sheet,
                        "taxonomy": actual_taxonomy_sheet,
                        "geocost": actual_geocost_sheet,
                        "ranges": {
                            key: {
                                "sheet": str((l1_mappings.get(key) or {}).get("sheet_name") or ""),
                                "range": str((l1_mappings.get(key) or {}).get("range_a1") or ""),
                            }
                            for key in STANDARD_INPUT_KEYS
                            if l1_mappings.get(key)
                        },
                    },
                    sort_keys=True,
                    ensure_ascii=False,
                ).encode("utf-8")
            ).hexdigest()
            # Keep the signature for diagnostics, but continue this render.
            # The Inputs rail is rendered after the canvas in app.py, so it can
            # consume the new mappings immediately without forcing a second
            # full workbook pass.
            st.session_state["auto_mapping_rerun_signature"] = mapping_signature

    taxonomy_profile: dict[str, Any]
    if (
        preview_available
        and preview_workbook_path is not None
        and bool(required_aura_keys)
    ):
        if load_progress is not None and load_progress.active:
            load_progress.update(
                79,
                "Validating taxonomy",
                "Checking the selected L0/L1 taxonomy against the detected function table.",
                phase="detect",
            )
        with ui_stage(
            "workbook.validate_taxonomy_mapping",
            warn_after=8,
            stack_after=25,
            sheet=actual_taxonomy_sheet,
        ):
            taxonomy_profile = _taxonomy_profile_from_mapping(
                workbook_path=preview_workbook_path,
                cache_token=token,
                mapping=l1_mappings.get("taxonomy"),
                fallback_sheet=actual_taxonomy_sheet,
            )
    else:
        taxonomy_profile = {
            "sheet_name": actual_taxonomy_sheet,
            "row_count": 0,
            "error": (
                "Upload a workbook before configuring the Taxonomy source."
                if required_aura_keys and not preview_available
                else ""
            ),
        }
    st.session_state["taxonomy_sheet_profile"] = taxonomy_profile
    st.session_state["resolved_taxonomy_sheet"] = str(
        taxonomy_profile.get("sheet_name") or actual_taxonomy_sheet
    )

    taxonomy_mapping = dict(l1_mappings.get("taxonomy") or {})
    fte_mapping = dict(l1_mappings.get("fte_values") or {})
    if (
        preview_available
        and preview_workbook_path is not None
        and bool(required_aura_keys)
    ):
        validation_mappings = {
            key: value
            for key, value in l1_mappings.items()
            if key in required_aura_keys
        }
        with ui_stage(
            "workbook.validate_master_labels",
            warn_after=10,
            stack_after=30,
            mapping_count=len(validation_mappings),
        ):
            master_label_profile = l1_master_label_profile(
                str(preview_workbook_path),
                token,
                _l1_mapping_ranges_json(validation_mappings),
            )
        if load_progress is not None and load_progress.active:
            load_progress.update(
                82,
                "Validating mapped tables",
                "Checking labels, dimensions, and benchmark-table alignment.",
                phase="detect",
            )
        company_count = _mapped_vector_length(
            dict(l1_mappings.get("company_headers") or {})
        )
        with ui_stage(
            "workbook.match_taxonomy_to_fte",
            warn_after=10,
            stack_after=30,
            company_count=company_count,
        ):
            function_lookup_profile = l1_function_lookup_profile(
                str(preview_workbook_path),
                token,
                str(taxonomy_mapping.get("sheet_name") or ""),
                str(taxonomy_mapping.get("range_a1") or ""),
                str(fte_mapping.get("sheet_name") or ""),
                str(fte_mapping.get("range_a1") or ""),
                str(fte_mapping.get("row_label_sheet") or ""),
                str(fte_mapping.get("row_label_range") or ""),
                company_count,
            )
        if load_progress is not None and load_progress.active:
            load_progress.update(
                85,
                "Matching functions",
                "Matching Taxonomy functions to their FTE and quartile source rows.",
                phase="detect",
            )
    else:
        master_label_profile = {
            "attempted": False,
            "ready": False,
            "labels": {},
            "errors": [],
        }
        function_lookup_profile = {
            "attempted": False,
            "ready": False,
            "matches": [],
            "missing": [],
            "ambiguous": [],
            "error": "",
        }
    st.session_state["l1_master_label_profile"] = master_label_profile
    st.session_state["l1_function_lookup_profile"] = function_lookup_profile

    geocost_table_mapping = dict(l1_mappings.get("geocost_table") or {})
    geocost_data_row_mapping = dict(l1_mappings.get("geocost_data_row") or {})
    if (
        preview_available
        and preview_workbook_path is not None
        and bool(required_geocost_keys)
    ):
        with ui_stage(
            "workbook.validate_geocost_mapping",
            warn_after=10,
            stack_after=30,
            sheet=str(geocost_table_mapping.get("sheet_name") or actual_geocost_sheet),
        ):
            geocost_profile = geocost_input_profile_from_mappings(
                str(preview_workbook_path),
                token,
                source_signature,
                str(geocost_table_mapping.get("sheet_name") or ""),
                str(geocost_table_mapping.get("range_a1") or ""),
                str(geocost_data_row_mapping.get("sheet_name") or ""),
                str(geocost_data_row_mapping.get("range_a1") or ""),
            )
    else:
        geocost_profile = {
            "ready": False,
            "source_signature": source_signature,
            "sheet_name": actual_geocost_sheet,
            "layout": {},
            "function_count": 0,
            "data_row_count": 0,
            "error": (
                "Upload a workbook before configuring Geo Cost inputs."
                if required_geocost_keys and not preview_available
                else ""
            ),
        }
    st.session_state["geocost_input_profile"] = geocost_profile

    company_mapping = dict(l1_mappings.get("company_headers") or {})
    revenue_mapping = dict(l1_mappings.get("revenue_values") or {})
    if (
        sort_enabled
        and preview_available
        and preview_workbook_path is not None
    ):
        if load_progress is not None and load_progress.active:
            load_progress.update(
                87,
                "Detecting company segments",
                "Reading the segment row above the company headers for universal sorting.",
                phase="detect",
            )
        with ui_stage(
            "workbook.detect_company_segments",
            warn_after=10,
            stack_after=30,
            company_range=str(company_mapping.get("range_a1") or ""),
        ):
            detected_segment_profile = l1_segment_detection_profile(
                str(preview_workbook_path),
                token,
                str(company_mapping.get("sheet_name") or actual_aura_sheet),
                str(company_mapping.get("range_a1") or ""),
                str(revenue_mapping.get("sheet_name") or actual_aura_sheet),
                str(revenue_mapping.get("range_a1") or ""),
            )
        segment_profile = _apply_l1_segment_assignment_overrides(
            detected_segment_profile,
            _get_l1_segment_assignment_overrides(source_signature),
        )
    else:
        segment_profile = {
            "detected": False,
            "segment_count": 0,
            "company_count": 0,
            "segment_order": [],
            "detected_segment_order": [],
            "segment_name_overrides": {},
            "segment_assignment_overrides": {},
            "assignments": [],
            "error": "Upload a workbook before detecting company segments.",
        }
    segment_profile_signature = hashlib.sha256(
        json.dumps(
            segment_profile,
            sort_keys=True,
            ensure_ascii=False,
            default=str,
        ).encode("utf-8")
    ).hexdigest()
    previous_segment_signature = str(
        st.session_state.get("l1_segment_profile_signature") or ""
    )
    st.session_state["l1_segment_profile"] = segment_profile
    st.session_state["l1_segment_profile_signature"] = segment_profile_signature
    if (
        str(st.session_state.get("l1_sort_order") or "") == "segment"
        and not bool(segment_profile.get("detected"))
    ):
        st.session_state["l1_sort_order"] = L1_SORT_ORDER_DEFAULT
    if (
        sort_enabled
        and segment_profile_signature != previous_segment_signature
    ):
        # Refresh the visible Company order status once, but never allow a
        # profile that oscillates because of workbook/style metadata to create
        # an endless Streamlit rerun cycle.
        # The Segment rail is rendered after the canvas and therefore sees
        # this profile in the same Streamlit pass. Avoid a second full workbook
        # render solely to refresh fixed-sidebar text.
        st.session_state["segment_profile_rerun_signature"] = (
            segment_profile_signature
        )

    input_mapping_required = bool(required_input_keys)
    standard_validation = validate_standard_input_mappings(
        l1_mappings,
        selected_sections=tuple(section_set),
        source_signature=source_signature,
        taxonomy_profile=taxonomy_profile,
        function_lookup_profile=function_lookup_profile,
        master_label_profile=master_label_profile,
        geocost_profile=geocost_profile,
        ignored_taxonomy_functions=_get_ignored_aura_taxonomy_functions(source_signature),
    )
    mapping_blocks_finalize = (
        input_mapping_required and not bool(standard_validation["ready"])
    )
    preview_blocks_finalize = not preview_available or bool(inventory_error) or not inventory

    finalize_blocker_summary = ""
    if preview_blocks_finalize:
        if not preview_available:
            finalize_help = "Cannot finalize yet: upload a workbook and wait for the workbook preview to become available."
            finalize_blocker_summary = "Workbook preview is not available yet."
        elif inventory_error:
            finalize_help = f"Cannot finalize yet: workbook inventory failed. {inventory_error}"
            finalize_blocker_summary = f"Workbook inventory failed: {inventory_error}"
        else:
            finalize_help = "Cannot finalize yet: the workbook contains no readable worksheets."
            finalize_blocker_summary = "No readable worksheets were found."
    elif mapping_blocks_finalize:
        missing_labels = [
            STANDARD_INPUT_BY_KEY[key]["label"]
            for key in (standard_validation.get("missing") or [])
            if key in STANDARD_INPUT_BY_KEY
        ]
        validation_errors = [
            str(value).strip()
            for value in (standard_validation.get("errors") or [])
            if str(value).strip()
        ]
        help_parts = ["Cannot finalize yet."]
        summary_parts: list[str] = []
        if missing_labels:
            missing_text = "Missing inputs: " + ", ".join(missing_labels) + "."
            help_parts.append(missing_text)
            summary_parts.append(missing_text)
        if validation_errors:
            issue_text = "Validation issue" + ("s" if len(validation_errors) > 1 else "") + ": " + " | ".join(validation_errors[:4])
            help_parts.append(issue_text)
            summary_parts.append("Validation: " + validation_errors[0])
        help_parts.append("Hover any Inputs card on the left for field-specific details, or click a card to review its selected range.")
        finalize_help = "\n\n".join(help_parts)
        finalize_blocker_summary = "  •  ".join(summary_parts) or "One or more required inputs have not passed validation."
    else:
        finalize_help = "All required workbook inputs pass validation. Finalize the workbook and continue to the build console."

    header_column, action_column = st.columns([7.0, 1.35])
    with header_column:
        st.markdown(
            """
            <span class="tc-workbook-mode" aria-hidden="true"></span>
            <div class="tc-workspace-header">
              <div class="tc-main-title">Workbook preview</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with action_column:
        with st.container(key="workbook_review_action"):
            finalize_review = st.button(
                "Finalize workbook",
                type="primary",
                use_container_width=True,
                disabled=preview_blocks_finalize or mapping_blocks_finalize,
                key="finalize_workbook_review",
            )

    if finalize_review:
        st.session_state["workbook_review_finalized"] = True
        st.session_state["canvas_stage"] = "console"
        st.rerun()

    if preview_error:
        st.warning(preview_error)

    if not preview_available:
        if load_progress is not None and load_progress.active:
            load_progress.fail("The uploaded workbook copy is not available for preview.")
        st.markdown(
            """
            <div class="tc-preview-empty">
              <div class="tc-preview-empty-icon">XLSX</div>
              <div class="tc-preview-empty-title">Upload a source workbook</div>
              <div class="tc-preview-empty-copy">Aura Summary, Taxonomy, and Geo Cost sources will appear here after upload.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    if inventory_error:
        if load_progress is not None and load_progress.active:
            load_progress.fail(f"Workbook inventory failed: {inventory_error}")
        st.error(f"Workbook preview could not be opened: {inventory_error}")
        return

    if not inventory:
        if load_progress is not None and load_progress.active:
            load_progress.fail("The workbook contains no readable worksheets.")
        st.warning("The workbook contains no readable worksheets.")
        return

    # Resolve the active destination before rendering the source controls.
    # Input buttons should still take the user to the relevant source once, but
    # they must not continuously own the preview. Otherwise every Streamlit
    # rerun immediately cancels a manual click on Aura Summary, Taxonomy, or
    # Geo Cost. The token below changes only when the active input or its
    # committed mapping changes. After that one synchronization, the source
    # controls are free-standing workbook tabs until another input is clicked.
    active_l1_input_key = (
        _active_l1_input_key(source_signature)
        if show_input_mapping and source_signature
        else ""
    )
    active_mapping = dict(l1_mappings.get(active_l1_input_key) or {})
    input_navigation_token = _active_input_navigation_token(
        source_signature=source_signature,
        input_key=active_l1_input_key,
        mapping=active_mapping,
    )
    previous_input_navigation_token = str(
        st.session_state.get("preview_input_navigation_token") or ""
    )
    if (
        input_navigation_token
        and input_navigation_token != previous_input_navigation_token
    ):
        preferred_role = _preview_role_for_active_input(
            input_key=active_l1_input_key,
            mapping=active_mapping,
            aura_sheet=actual_aura_sheet,
            taxonomy_sheet=actual_taxonomy_sheet,
            geocost_sheet=actual_geocost_sheet,
        )
        if preferred_role:
            st.session_state["preview_source_role"] = preferred_role
        mapping_sheet = str(active_mapping.get("sheet_name") or "").strip()
        if mapping_sheet and any(
            str(item.get("name") or "").casefold() == mapping_sheet.casefold()
            for item in inventory
        ):
            # Consume this exact-sheet reveal once.  This is especially
            # important when a workbook contains both Global and US Geo Cost
            # tabs: the generic Geo Cost role alone is not enough to guarantee
            # that the preview opens the worksheet that owns the committed
            # range.  Source-tab browsing remains free after this one reveal.
            st.session_state["pending_input_preview_sheet"] = mapping_sheet
            st.session_state["pending_input_preview_token"] = input_navigation_token
        st.session_state["preview_input_navigation_token"] = (
            input_navigation_token
        )

    inventory, selected_sheet = _render_source_selectors(
        preview_workbook_path,
        workflow_mode,
        selected_sections=selected_sections,
        preview_enabled=True,
        inventory=inventory,
    )
    taxonomy_profile = dict(st.session_state.get("taxonomy_sheet_profile") or taxonomy_profile)
    # The standardized Aura and Geo Cost inputs are rendered once in the adjacent Inputs
    # rail. The workbook canvas only consumes the active range destination.

    pending_preview_sheet = str(
        st.session_state.pop("pending_input_preview_sheet", "") or ""
    ).strip()
    st.session_state.pop("pending_input_preview_token", None)
    if pending_preview_sheet:
        exact_match = next(
            (
                str(item.get("name") or "")
                for item in inventory
                if str(item.get("name") or "").casefold()
                == pending_preview_sheet.casefold()
            ),
            "",
        )
        if exact_match:
            selected_sheet = exact_match
            st.session_state["preview_sheet_name"] = exact_match

    inventory_by_name = {
        str(item.get("name") or ""): item
        for item in inventory
        if item.get("name")
    }
    if selected_sheet not in inventory_by_name:
        selected_sheet = str(inventory[0].get("name") or "")
        st.session_state["preview_sheet_name"] = selected_sheet
    if not selected_sheet:
        if load_progress is not None and load_progress.active:
            load_progress.fail("No worksheet is available for preview.")
        st.warning("No worksheet is available for preview.")
        return

    sheet_info = inventory_by_name[selected_sheet]
    dimension_max_row = max(1, int(sheet_info["rows"]))
    dimension_max_column = max(1, int(sheet_info["columns"]))
    try:
        if load_progress is not None and load_progress.active:
            load_progress.update(
                91,
                "Sizing the preview",
                f"Finding the final table and meaningful used range on {selected_sheet}.",
                phase="preview",
            )
        with ui_stage(
            "workbook.scan_preview_extent",
            warn_after=10,
            stack_after=30,
            sheet=selected_sheet,
            sheet_part=str(sheet_info["sheet_part"]),
        ):
            content_extent = worksheet_content_extent(
                str(preview_workbook_path),
                token,
                str(sheet_info["sheet_part"]),
            )
    except Exception as exc:
        ui_log(
            "PREVIEW EXTENT FALLBACK",
            level="WARNING",
            sheet=selected_sheet,
            error=f"{type(exc).__name__}: {exc}",
        )
        content_extent = {
            "last_content_row": dimension_max_row,
            "last_content_column": dimension_max_column,
            "last_table_row": dimension_max_row,
            "preview_end_row": min(
                EXCEL_MAX_ROWS,
                dimension_max_row + PREVIEW_ROWS_AFTER_LAST_TABLE,
            ),
        }

    last_content_row = max(1, int(content_extent.get("last_content_row") or 1))
    last_content_column = max(
        1,
        int(content_extent.get("last_content_column") or dimension_max_column),
    )
    last_table_row = max(
        1,
        int(content_extent.get("last_table_row") or last_content_row),
    )
    auto_preview_end_row = max(
        1,
        min(
            EXCEL_MAX_ROWS,
            int(
                content_extent.get("preview_end_row")
                or (last_table_row + PREVIEW_ROWS_AFTER_LAST_TABLE)
            ),
        ),
    )
    preview_row_state_key = _preview_row_window_key(
        source_signature or "workbook",
        selected_sheet,
    )
    requested_preview_start_row = int(
        st.session_state.get(preview_row_state_key, 1) or 1
    )
    preview_window = _bounded_formatted_preview_window(
        sheet_name=selected_sheet,
        sheet_role=str(sheet_info.get("role") or ""),
        auto_end_row=auto_preview_end_row,
        last_content_column=last_content_column,
        active_mapping=active_mapping,
        requested_start_row=requested_preview_start_row,
    )
    start_row = int(preview_window["start_row"])
    row_count = int(preview_window["row_count"])
    start_column = int(preview_window["start_column"])
    column_count = int(preview_window["column_count"])
    preview_end_row = int(preview_window["end_row"])
    preview_max_column = int(preview_window["end_column"])
    preview_limited = bool(preview_window["limited"])
    st.session_state[preview_row_state_key] = start_row

    if preview_end_row < auto_preview_end_row or start_row > 1:
        _render_preview_row_navigation(
            state_key=preview_row_state_key,
            start_row=start_row,
            end_row=preview_end_row,
            total_rows=auto_preview_end_row,
            row_count=row_count,
        )

    try:
        if load_progress is not None and load_progress.active:
            load_progress.update(
                95,
                "Rendering workbook preview",
                f"Applying Excel formatting for {selected_sheet} and loading the interactive range selector.",
                phase="preview",
            )
        with ui_stage(
            "workbook.render_formatted_preview",
            warn_after=12,
            stack_after=35,
            sheet=selected_sheet,
            rows=f"{start_row}:{preview_end_row}",
            columns=f"{start_column}:{preview_max_column}",
            requested_cells=row_count * column_count,
        ):
            preview = workbook_sheet_formatted_preview(
                str(preview_workbook_path),
                token,
                selected_sheet,
                str(sheet_info["sheet_part"]),
                start_row,
                row_count,
                start_column,
                column_count,
            )
    except Exception as exc:
        ui_log(
            "FORMATTED PREVIEW FAILED",
            level="ERROR",
            sheet=selected_sheet,
            error=f"{type(exc).__name__}: {exc}",
        )
        if load_progress is not None and load_progress.active:
            load_progress.fail(f"Formatted preview failed: {exc}")
        st.error(f"Could not read worksheet preview: {exc}")
        return

    active_l1_label = (
        STANDARD_INPUT_BY_KEY[active_l1_input_key]["label"]
        if active_l1_input_key in STANDARD_INPUT_BY_KEY
        else ""
    )
    active_input_status = {"level": "info", "message": ""}
    active_input_ignore_mismatch = ""
    active_input_restore_mismatch = ""
    if active_l1_input_key in STANDARD_INPUT_BY_KEY:
        active_input_status = standard_input_status(
            field=STANDARD_INPUT_BY_KEY[active_l1_input_key],
            mapping=(
                dict(l1_mappings.get(active_l1_input_key) or {})
                if l1_mappings.get(active_l1_input_key)
                else None
            ),
            workbook_ready=preview_available,
            validation=standard_validation,
            auto_error="",
        )
        if active_l1_input_key in {"taxonomy", "fte_values"}:
            ignored_names = _get_ignored_aura_taxonomy_functions(source_signature)
            ignored_keys = {" ".join(value.split()).casefold() for value in ignored_names}
            missing_names = [
                " ".join(str(value or "").split()).strip()
                for value in (function_lookup_profile.get("missing") or [])
                if " ".join(str(value or "").split()).strip()
            ]
            active_missing = [
                value for value in missing_names
                if value.casefold() not in ignored_keys
            ]
            ignored_missing = [
                value for value in missing_names
                if value.casefold() in ignored_keys
            ]
            if str(active_input_status.get("level") or "") == "error" and active_missing:
                active_input_ignore_mismatch = active_missing[0]
            elif ignored_missing:
                active_input_restore_mismatch = ignored_missing[0]
    _render_active_input_validation_banner(
        source_signature=source_signature,
        status=active_input_status,
        ignore_mismatch_name=active_input_ignore_mismatch,
        restore_mismatch_name=active_input_restore_mismatch,
    )
    _render_interactive_range_selector(
        preview=preview,
        sheet_name=selected_sheet,
        source_signature=source_signature or "workbook",
        max_row=max(auto_preview_end_row, dimension_max_row),
        max_column=max(last_content_column, dimension_max_column),
        mapping_section="inputs" if show_input_mapping else "",
        active_input_key=active_l1_input_key,
        active_input_label=active_l1_label,
        active_input_status=str(active_input_status.get("level") or "info"),
        active_input_status_message=str(active_input_status.get("message") or ""),
        active_input_ignore_mismatch=active_input_ignore_mismatch,
        active_input_restore_mismatch=active_input_restore_mismatch,
    )
    if load_progress is not None and load_progress.active:
        load_progress.complete(
            f"{selected_sheet} is ready for review and range selection."
        )
        load_progress.clear()



def inspect_uploaded_workbook(workbook_path: Path) -> UploadInspection:
    """Prepare metadata used by the safe, relevant-tab-only preview."""
    resolved = workbook_path.resolve()
    metadata = {
        "source_size": int(resolved.stat().st_size),
        "source_untouched": True,
        "cached_values": True,
        "links_disabled": True,
    }
    ui_log(
        "WORKBOOK INSPECTION REQUEST",
        path=str(resolved),
        size_bytes=metadata["source_size"],
    )
    try:
        with ui_stage(
            "workbook.inspect_manifest",
            warn_after=8,
            stack_after=25,
            path=str(resolved),
        ):
            relevant_tabs = workbook_manifest(
                str(resolved),
                _workbook_cache_token(resolved),
            )
        preview_error = None
        if not relevant_tabs:
            preview_error = (
                "No Aura Summary, Taxonomy, or Geo Cost tab was identified by name. "
                "Slide generation remains available because the backend performs "
                "its own structural validation."
            )
        ui_log(
            "WORKBOOK INSPECTION READY",
            path=str(resolved),
            relevant_tabs=[item.get("name") for item in relevant_tabs],
        )
    except Exception as exc:
        relevant_tabs = []
        preview_error = (
            "The uploaded copy is available for slide generation, but the "
            f"relevant-tab preview could not be prepared: {exc}"
        )
        ui_log(
            "WORKBOOK INSPECTION DEGRADED",
            level="ERROR",
            path=str(resolved),
            error=f"{type(exc).__name__}: {exc}",
        )

    return UploadInspection(
        preview_path=resolved,
        metadata=metadata,
        preview_error=preview_error,
        relevant_tab_count=len(relevant_tabs),
    )


# -----------------------------------------------------------------------------
# Build-console source overrides
# -----------------------------------------------------------------------------


_AUTO_SOURCE_OPTION = "__auto_detect__"
_SHEET_OPTION_PREFIX = "sheet::"
_US_STATE_KEYS = {
    re.sub(r"[^a-z0-9]+", "", value.lower())
    for value in (
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


def _sheet_option(sheet_name: str) -> str:
    return f"{_SHEET_OPTION_PREFIX}{sheet_name}"


def _sheet_name_from_option(option: str) -> str:
    value = str(option or "")
    return value[len(_SHEET_OPTION_PREFIX):] if value.startswith(_SHEET_OPTION_PREFIX) else ""


def _normalized_location_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


@st.cache_data(show_spinner=False, max_entries=24)
def _infer_geocost_scope_from_sheet(
    path_text: str,
    cache_token: int,
    sheet_name: str,
) -> str:
    """Infer Global versus US for an explicitly selected, generically named tab."""
    del cache_token
    from openpyxl import load_workbook

    try:
        workbook = load_workbook(
            path_text,
            read_only=True,
            data_only=True,
            keep_links=False,
        )
    except TypeError:
        workbook = load_workbook(
            path_text,
            read_only=True,
            data_only=True,
        )
    try:
        if sheet_name not in workbook.sheetnames:
            return "global"
        worksheet = workbook[sheet_name]
        max_row = max(1, min(int(worksheet.max_row or 1), 240))
        max_column = max(1, min(int(worksheet.max_column or 1), 100))
        rows = [
            list(row)
            for row in worksheet.iter_rows(
                min_row=1,
                max_row=max_row,
                min_col=1,
                max_col=max_column,
                values_only=True,
            )
        ]

        best_header_row = -1
        best_entity_column = -1
        best_score = -1
        entity_header = ""
        for row_index, row in enumerate(rows[:35]):
            keys = [_normalized_location_key(value) for value in row]
            has_region = any(key == "region" for key in keys)
            has_high_low = any(
                key in {"highlow", "highorlow", "costtier"}
                for key in keys
            )
            entity_columns = [
                index
                for index, key in enumerate(keys)
                if key in {"country", "state", "location", "geography"}
            ]
            score = int(has_region) + int(has_high_low) + int(bool(entity_columns))
            if score > best_score and entity_columns:
                best_score = score
                best_header_row = row_index
                best_entity_column = entity_columns[0]
                entity_header = keys[best_entity_column]

        if entity_header == "state":
            return "us"

        if best_header_row >= 0 and best_entity_column >= 0:
            location_keys: list[str] = []
            blank_streak = 0
            for row in rows[best_header_row + 1:]:
                value = row[best_entity_column] if best_entity_column < len(row) else None
                key = _normalized_location_key(value)
                if not key:
                    blank_streak += 1
                    if blank_streak >= 30 and location_keys:
                        break
                    continue
                blank_streak = 0
                location_keys.append(key)
            state_count = sum(key in _US_STATE_KEYS for key in location_keys)
            if len(location_keys) >= 5 and state_count / len(location_keys) >= 0.6:
                return "us"

        return "global"
    finally:
        workbook.close()


def _selector_inventory(
    inventory: list[dict[str, Any]],
    preferred_roles: set[str],
) -> list[dict[str, Any]]:
    preferred = [item for item in inventory if str(item.get("role")) in preferred_roles]
    other = [item for item in inventory if item not in preferred]
    return preferred + other


def _auto_source_label(
    label: str,
    detected: list[dict[str, Any]],
) -> str:
    if len(detected) == 1:
        return f"Auto-detect · {detected[0]['name']}"
    if len(detected) > 1:
        return f"Auto-detect · {len(detected)} candidate tabs"
    return f"Auto-detect · {label}"


def _source_option_labels(
    inventory: list[dict[str, Any]],
    *,
    selector: str,
) -> tuple[list[str], dict[str, str]]:
    if selector == "aura":
        detected = [item for item in inventory if item.get("role") == "aura"]
        ordered = _selector_inventory(inventory, {"aura"})
        auto_label = _auto_source_label("Aura Summary", detected)
    elif selector == "taxonomy":
        detected = [item for item in inventory if item.get("role") == "taxonomy"]
        ordered = _selector_inventory(inventory, {"taxonomy"})
        auto_label = _auto_source_label("Taxonomy", detected)
    else:
        detected = [
            item
            for item in inventory
            if str(item.get("role") or "").startswith("geocost")
        ]
        ordered = _selector_inventory(
            inventory,
            {"geocost_global", "geocost_us", "geocost"},
        )
        auto_label = _auto_source_label("Geo Cost", detected)

    options = [_AUTO_SOURCE_OPTION]
    labels = {_AUTO_SOURCE_OPTION: auto_label}
    for item in ordered:
        name = str(item.get("name") or "")
        if not name:
            continue
        option = _sheet_option(name)
        options.append(option)
        role = str(item.get("role") or "")
        suffix = ""
        if selector == "aura" and role == "aura":
            suffix = "  · detected Aura Summary"
        elif selector == "taxonomy" and role == "taxonomy":
            suffix = "  · detected Taxonomy"
        elif selector == "geocost" and role == "geocost_global":
            suffix = "  · detected Global"
        elif selector == "geocost" and role == "geocost_us":
            suffix = "  · detected US"
        elif selector == "geocost" and role == "geocost":
            suffix = "  · detected Geo Cost"
        if str(item.get("state") or "visible") != "visible":
            suffix += f"  · {item['state']}"
        labels[option] = f"{name}{suffix}"
    return options, labels


def _initialize_selector_key(
    key: str,
    options: list[str],
    current_sheet: str,
) -> None:
    if key in st.session_state and st.session_state[key] in options:
        return
    requested = _sheet_option(current_sheet) if current_sheet else _AUTO_SOURCE_OPTION
    st.session_state[key] = requested if requested in options else _AUTO_SOURCE_OPTION


def _auto_detected_source_sheet(
    inventory: list[dict[str, Any]],
    selector: str,
) -> str:
    """Return the strongest name-based source hint for preview purposes."""
    if selector == "aura":
        role_priority = {"aura": 0}
    elif selector == "taxonomy":
        role_priority = {"taxonomy": 0}
    else:
        role_priority = {
            "geocost_global": 0,
            "geocost_us": 1,
            "geocost": 2,
        }

    candidates = [
        item
        for item in inventory
        if str(item.get("role") or "") in role_priority
    ]
    candidates.sort(
        key=lambda item: (
            role_priority.get(str(item.get("role") or ""), 99),
            _taxonomy_sheet_sort_key(
                str(item.get("name") or ""),
                int(item.get("workbook_order", 9999)),
            )
            if selector == "taxonomy"
            else (0, 0, int(item.get("workbook_order", 9999))),
        )
    )
    return str(candidates[0].get("name") or "") if candidates else ""


def _set_preview_source_role(role: str) -> None:
    st.session_state["preview_source_role"] = role


def _active_input_navigation_token(
    *,
    source_signature: str,
    input_key: str,
    mapping: dict[str, Any] | None,
) -> str:
    """Return a stable token for one input-driven preview navigation event.

    The token intentionally excludes the current preview role. A source-tab
    click therefore survives unrelated Streamlit reruns. It changes only when
    the user chooses a different input or that input is assigned a different
    committed range.
    """
    if not source_signature or not input_key:
        return ""
    value = mapping or {}
    payload = (
        str(source_signature),
        str(input_key),
        str(value.get("sheet_name") or "").strip(),
        str(value.get("range_a1") or "").strip(),
    )
    return hashlib.sha256(repr(payload).encode("utf-8")).hexdigest()


def _preview_role_for_active_input(
    *,
    input_key: str,
    mapping: dict[str, Any] | None,
    aura_sheet: str,
    taxonomy_sheet: str,
    geocost_sheet: str = "",
) -> str:
    """Resolve the source tab used when a standardized input is selected."""
    if not input_key:
        return ""
    if input_key == "taxonomy":
        return "taxonomy"
    if input_key in GEOCOST_INPUT_KEYS:
        return "geocost"

    mapping_sheet = str((mapping or {}).get("sheet_name") or "").strip()
    if (
        mapping_sheet
        and taxonomy_sheet
        and mapping_sheet.casefold() == str(taxonomy_sheet).strip().casefold()
    ):
        return "taxonomy"
    if (
        mapping_sheet
        and geocost_sheet
        and mapping_sheet.casefold() == str(geocost_sheet).strip().casefold()
    ):
        return "geocost"
    if aura_sheet:
        return "aura"
    return ""


def _preview_role_sheet_map(
    actual_aura: str,
    actual_taxonomy: str,
    actual_geocost: str,
) -> dict[str, str]:
    return {
        "aura": str(actual_aura or "").strip(),
        "taxonomy": str(actual_taxonomy or "").strip(),
        "geocost": str(actual_geocost or "").strip(),
    }


def _fallback_preview_role(
    role_sheets: dict[str, str],
    selected_sections: Iterable[str] | None = None,
    workflow_mode: str = "",
) -> str:
    section_set = _selected_section_set(selected_sections, workflow_mode)
    preferred_order = (
        ("geocost", "aura", "taxonomy")
        if section_set == {"geocost"}
        else ("aura", "taxonomy", "geocost")
    )
    for role in preferred_order:
        if role_sheets.get(role):
            return role
    return "aura"


def _source_control_payload(
    *,
    role: str,
    title: str,
    choice: str,
    options: list[str],
    labels: dict[str, str],
    actual_sheet: str,
) -> dict[str, Any]:
    return {
        "role": role,
        "title": title,
        "choice": choice,
        "status": labels.get(choice, choice),
        "actual_sheet": actual_sheet,
        "options": [
            {
                "value": option,
                "label": labels.get(option, option),
            }
            for option in options
        ],
    }


def _render_combined_source_controls(
    *,
    signature: str,
    preview_enabled: bool,
    active_role: str,
    controls: list[dict[str, Any]],
) -> dict[str, Any] | None:
    """Render one compact split control per source.

    The left side switches the visible preview. The integrated chevron opens
    the Auto-detect/manual worksheet menu. This replaces the former two-row
    selectbox-plus-toggle layout.
    """
    component = _get_source_selector_component()
    return component(
        controls=controls,
        preview_enabled=bool(preview_enabled),
        active_role=active_role if preview_enabled else "",
        key=(
            f"source_split_selector_{signature}_"
            f"{'review' if preview_enabled else 'console'}"
        ),
        default=None,
    )


def _render_source_selectors(
    source_path: Path | None,
    workflow_mode: str,
    *,
    selected_sections: Iterable[str] | None = None,
    preview_enabled: bool,
    inventory: list[dict[str, Any]] | None = None,
) -> tuple[list[dict[str, Any]], str]:
    """Render combined source/view controls for Aura, Taxonomy, and Geo Cost.

    Each source is now a single split control: click its main area to preview
    that source, or click its integrated chevron to keep Auto-detect or choose
    a worksheet manually. Taxonomy Auto-detect remains name-based and requires
    a tab name containing ``Taxonomy`` case-insensitively.
    """
    resolved_inventory = list(inventory or [])
    cache_token = 0
    signature = str(st.session_state.get("uploaded_signature") or "workbook")[:12]
    inventory_error = ""

    if source_path is not None and source_path.is_file():
        cache_token = _workbook_cache_token(source_path)
        if inventory is None:
            try:
                resolved_inventory = workbook_sheet_inventory(
                    str(source_path),
                    cache_token,
                )
            except Exception as exc:
                inventory_error = str(exc)
                resolved_inventory = []

    aura_options, aura_labels = _source_option_labels(
        resolved_inventory,
        selector="aura",
    )
    taxonomy_options, taxonomy_labels = _source_option_labels(
        resolved_inventory,
        selector="taxonomy",
    )
    geo_options, geo_labels = _source_option_labels(
        resolved_inventory,
        selector="geocost",
    )

    aura_key = f"source_aura_{signature}"
    taxonomy_key = f"source_taxonomy_{signature}"
    geo_key = f"source_geocost_{signature}"
    _initialize_selector_key(
        aura_key,
        aura_options,
        str(st.session_state.get("selected_aura_sheet_override") or ""),
    )
    _initialize_selector_key(
        taxonomy_key,
        taxonomy_options,
        str(st.session_state.get("selected_taxonomy_sheet_override") or ""),
    )
    _initialize_selector_key(
        geo_key,
        geo_options,
        str(st.session_state.get("selected_geocost_sheet_override") or ""),
    )

    def current_source_state() -> dict[str, Any]:
        aura_choice = str(st.session_state.get(aura_key) or _AUTO_SOURCE_OPTION)
        taxonomy_choice = str(st.session_state.get(taxonomy_key) or _AUTO_SOURCE_OPTION)
        geocost_choice = str(st.session_state.get(geo_key) or _AUTO_SOURCE_OPTION)

        selected_aura = _sheet_name_from_option(aura_choice)
        selected_taxonomy = _sheet_name_from_option(taxonomy_choice)
        selected_geocost = _sheet_name_from_option(geocost_choice)

        actual_aura = selected_aura or _auto_detected_source_sheet(
            resolved_inventory,
            "aura",
        )
        actual_taxonomy = selected_taxonomy or _auto_detected_source_sheet(
            resolved_inventory,
            "taxonomy",
        )
        actual_geocost = selected_geocost or _auto_detected_source_sheet(
            resolved_inventory,
            "geocost",
        )

        return {
            "aura_choice": aura_choice,
            "taxonomy_choice": taxonomy_choice,
            "geocost_choice": geocost_choice,
            "selected_aura": selected_aura,
            "selected_taxonomy": selected_taxonomy,
            "selected_geocost": selected_geocost,
            "actual_aura": actual_aura,
            "actual_taxonomy": actual_taxonomy,
            "actual_geocost": actual_geocost,
        }

    source_state = current_source_state()
    role_sheets = _preview_role_sheet_map(
        source_state["actual_aura"],
        source_state["actual_taxonomy"],
        source_state["actual_geocost"],
    )
    active_role = str(
        st.session_state.get("preview_source_role") or ""
    ).strip().lower()
    if active_role not in role_sheets or not role_sheets.get(active_role):
        active_role = _fallback_preview_role(
            role_sheets, selected_sections, workflow_mode
        )
        st.session_state["preview_source_role"] = active_role

    controls = [
        _source_control_payload(
            role="aura",
            title="Aura Summary",
            choice=source_state["aura_choice"],
            options=aura_options,
            labels=aura_labels,
            actual_sheet=source_state["actual_aura"],
        ),
        _source_control_payload(
            role="taxonomy",
            title="Taxonomy",
            choice=source_state["taxonomy_choice"],
            options=taxonomy_options,
            labels=taxonomy_labels,
            actual_sheet=source_state["actual_taxonomy"],
        ),
        _source_control_payload(
            role="geocost",
            title="Geo Cost",
            choice=source_state["geocost_choice"],
            options=geo_options,
            labels=geo_labels,
            actual_sheet=source_state["actual_geocost"],
        ),
    ]

    with st.container(key="source_selector_row"):
        try:
            event = _render_combined_source_controls(
                signature=signature,
                preview_enabled=preview_enabled,
                active_role=active_role,
                controls=controls,
            )
        except Exception as exc:
            event = None
            st.warning(
                "The compact source controls could not be initialized in this "
                f"Streamlit runtime. Details: {exc}"
            )

    event_state_key = (
        f"source_split_last_event_{signature}_"
        f"{'review' if preview_enabled else 'console'}"
    )
    if isinstance(event, dict):
        event_id = str(event.get("event_id") or "")
        last_event_id = str(st.session_state.get(event_state_key) or "")
        if event_id and event_id != last_event_id:
            st.session_state[event_state_key] = event_id
            action = str(event.get("action") or "").strip().lower()
            role = str(event.get("role") or "").strip().lower()
            value = str(event.get("value") or "")

            role_config = {
                "aura": (aura_key, aura_options),
                "taxonomy": (taxonomy_key, taxonomy_options),
                "geocost": (geo_key, geo_options),
            }
            if action == "preview" and preview_enabled:
                if role in role_sheets and role_sheets.get(role):
                    st.session_state["preview_source_role"] = role
                    ui_log(
                        "PREVIEW SOURCE TAB",
                        role=role,
                        sheet=str(role_sheets.get(role) or ""),
                    )
            elif action == "select" and role in role_config:
                state_key, valid_options = role_config[role]
                if value in valid_options:
                    st.session_state[state_key] = value
                    if preview_enabled:
                        st.session_state["preview_source_role"] = role

            rerun = getattr(st, "rerun", None) or getattr(
                st,
                "experimental_rerun",
                None,
            )
            if callable(rerun):
                rerun()

    # Read the final synchronized choices after any component event.
    source_state = current_source_state()
    selected_aura = source_state["selected_aura"]
    selected_taxonomy = source_state["selected_taxonomy"]
    selected_geocost = source_state["selected_geocost"]
    actual_aura = source_state["actual_aura"]
    actual_taxonomy = source_state["actual_taxonomy"]
    actual_geocost = source_state["actual_geocost"]

    st.session_state["selected_aura_sheet_override"] = selected_aura
    st.session_state["selected_taxonomy_sheet_override"] = selected_taxonomy
    st.session_state["selected_geocost_sheet_override"] = selected_geocost
    st.session_state["selected_geocost_scope_override"] = ""
    st.session_state["selected_geocost_global_sheet_override"] = ""
    st.session_state["selected_geocost_us_sheet_override"] = ""

    if selected_geocost and source_path is not None and source_path.is_file():
        item = next(
            (
                entry
                for entry in resolved_inventory
                if entry.get("name") == selected_geocost
            ),
            {},
        )
        role = str(item.get("role") or "")
        if role == "geocost_us":
            scope = "us"
        elif role == "geocost_global":
            scope = "global"
        else:
            scope = _infer_geocost_scope_from_sheet(
                str(source_path),
                cache_token,
                selected_geocost,
            )
        st.session_state["selected_geocost_scope_override"] = scope
        if scope == "us":
            st.session_state["selected_geocost_us_sheet_override"] = selected_geocost
        else:
            st.session_state["selected_geocost_global_sheet_override"] = selected_geocost

    st.session_state["resolved_aura_sheet"] = actual_aura
    st.session_state["resolved_taxonomy_sheet"] = actual_taxonomy
    st.session_state["resolved_geocost_sheet"] = actual_geocost

    if source_path is not None and source_path.is_file():
        full_signature = str(st.session_state.get("uploaded_signature") or "")
        taxonomy_mapping = _get_l1_input_mappings(full_signature).get("taxonomy")
        taxonomy_profile = _taxonomy_profile_from_mapping(
            workbook_path=source_path,
            cache_token=cache_token,
            mapping=taxonomy_mapping,
            fallback_sheet=actual_taxonomy,
        )
    else:
        taxonomy_profile = {
            "sheet_name": actual_taxonomy,
            "row_count": 0,
            "error": "No workbook is available for Taxonomy validation.",
        }
    st.session_state["taxonomy_sheet_profile"] = taxonomy_profile
    st.session_state["resolved_taxonomy_sheet"] = str(
        taxonomy_profile.get("sheet_name") or actual_taxonomy
    )

    role_sheets = _preview_role_sheet_map(
        actual_aura,
        actual_taxonomy,
        actual_geocost,
    )
    active_role = str(
        st.session_state.get("preview_source_role") or ""
    ).strip().lower()
    if active_role not in role_sheets or not role_sheets.get(active_role):
        active_role = _fallback_preview_role(
            role_sheets, selected_sections, workflow_mode
        )
        st.session_state["preview_source_role"] = active_role

    preview_sheet = str(role_sheets.get(active_role) or "").strip()
    if not preview_sheet and resolved_inventory:
        preview_sheet = str(resolved_inventory[0].get("name") or "")

    if preview_sheet:
        st.session_state["preview_sheet_name"] = preview_sheet

    if inventory_error:
        st.caption(
            "Workbook tabs could not be listed for manual source selection; "
            f"backend auto-detection remains active. ({inventory_error})"
        )

    return resolved_inventory, preview_sheet


# -----------------------------------------------------------------------------
# Backend command and output discovery
# -----------------------------------------------------------------------------


def _write_geocost_input_manifest(options: BuildOptions) -> Path:
    """Persist the two-range Geo Cost wrapper contract for backend parsing."""
    mappings = dict(options.geocost_input_mappings or {})
    table = dict(mappings.get("geocost_table") or {})
    data_row = dict(mappings.get("geocost_data_row") or {})
    source_signature = str(
        table.get("source_signature")
        or data_row.get("source_signature")
        or st.session_state.get("uploaded_signature")
        or ""
    )
    profile = geocost_input_profile_from_mappings(
        str(options.workbook),
        _workbook_cache_token(options.workbook),
        source_signature,
        str(table.get("sheet_name") or ""),
        str(table.get("range_a1") or ""),
        str(data_row.get("sheet_name") or ""),
        str(data_row.get("range_a1") or ""),
    )
    validation = validate_geocost_input_mappings(
        mappings,
        source_signature=source_signature,
        profile=profile,
        required_keys=GEOCOST_INPUT_KEYS,
    )
    if not validation["ready"]:
        details = list(validation.get("errors") or [])
        if validation.get("missing"):
            details.append(
                "Missing: "
                + ", ".join(
                    GEOCOST_INPUT_BY_KEY[key]["label"]
                    for key in validation["missing"]
                    if key in GEOCOST_INPUT_BY_KEY
                )
            )
        raise ValueError(
            "Geo Cost input mapping is incomplete or invalid. "
            + (" ".join(details) if details else "Review the mapped ranges.")
        )

    sheet_name = str(profile.get("sheet_name") or table.get("sheet_name") or "")
    scope_hint = _infer_geocost_scope_from_sheet(
        str(options.workbook),
        _workbook_cache_token(options.workbook),
        sheet_name,
    )
    inputs = {
        key: mappings[key]
        for key in GEOCOST_INPUT_KEYS
        if isinstance(mappings.get(key), dict)
    }
    payload = {
        "schema_version": GEOCOST_INPUT_MAP_SCHEMA_VERSION,
        "slide_family": "geocost",
        "workbook_path": str(options.workbook.resolve()),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "sheet_name": sheet_name,
        "scope_hint": scope_hint,
        "inputs": inputs,
        "layout": dict(profile.get("layout") or {}),
        "function_count": int(profile.get("function_count") or 0),
        "data_row_count": int(profile.get("data_row_count") or 0),
    }
    identity = json.dumps(payload, sort_keys=True, ensure_ascii=False, default=str)
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:12]
    manifest_dir = options.workbook.parent.parent / "input_maps"
    manifest_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = manifest_dir / f"geocost_input_map_{digest}.json"
    temporary_path = manifest_path.with_suffix(".json.writing")
    temporary_path.write_text(
        json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False),
        encoding="utf-8",
    )
    temporary_path.replace(manifest_path)
    return manifest_path.resolve()


def _write_aura_input_manifest(options: BuildOptions) -> Path:
    """Persist the standardized Aura range contract for L0 and/or L1."""
    mappings = dict(
        options.aura_input_mappings
        or options.l1_input_mappings
        or {}
    )
    required_keys = aura_input_keys_for_sections(options.selected_sections)
    if not required_keys:
        raise ValueError("The selected output does not currently use mapped Aura inputs.")

    taxonomy_mapping = mappings.get("taxonomy") or {}
    taxonomy_profile = taxonomy_range_profile(
        str(options.workbook),
        _workbook_cache_token(options.workbook),
        str(taxonomy_mapping.get("sheet_name") or options.taxonomy_sheet),
        str(taxonomy_mapping.get("range_a1") or ""),
    )
    fte_mapping = dict(mappings.get("fte_values") or {})
    company_count = _mapped_vector_length(
        dict(mappings.get("company_headers") or {})
    )
    validation_mappings = {
        key: value for key, value in mappings.items() if key in required_keys
    }
    master_label_profile = l1_master_label_profile(
        str(options.workbook),
        _workbook_cache_token(options.workbook),
        _l1_mapping_ranges_json(validation_mappings),
    )
    function_lookup_profile = l1_function_lookup_profile(
        str(options.workbook),
        _workbook_cache_token(options.workbook),
        str(taxonomy_mapping.get("sheet_name") or options.taxonomy_sheet),
        str(taxonomy_mapping.get("range_a1") or ""),
        str(fte_mapping.get("sheet_name") or ""),
        str(fte_mapping.get("range_a1") or ""),
        str(fte_mapping.get("row_label_sheet") or ""),
        str(fte_mapping.get("row_label_range") or ""),
        company_count,
    )
    validation = validate_l1_input_mappings(
        mappings,
        taxonomy_profile=taxonomy_profile,
        function_lookup_profile=function_lookup_profile,
        master_label_profile=master_label_profile,
        required_keys=required_keys,
        ignored_taxonomy_functions=options.ignored_aura_taxonomy_functions,
    )
    if not validation["ready"]:
        details = list(validation.get("errors") or [])
        if validation.get("missing"):
            missing_labels = [
                AURA_INPUT_BY_KEY[key]["label"]
                for key in validation["missing"]
                if key in AURA_INPUT_BY_KEY
            ]
            details.append("Missing: " + ", ".join(missing_labels))
        raise ValueError(
            "Aura input mapping is incomplete or invalid. "
            + (" ".join(details) if details else "Review the mapped ranges.")
        )

    taxonomy_sheet = str(taxonomy_profile.get("sheet_name") or "").strip()
    taxonomy_range = str(taxonomy_mapping.get("range_a1") or "").strip()
    manifest_dir = options.workbook.parent.parent / "input_maps"
    manifest_dir.mkdir(parents=True, exist_ok=True)
    segment_profile = dict(options.l1_segment_profile or {})
    segment_payload = {
        "detected": bool(segment_profile.get("detected")),
        "sheet_name": str(segment_profile.get("sheet_name") or ""),
        "company_range": str(segment_profile.get("company_range") or ""),
        "segment_range": str(segment_profile.get("segment_range") or ""),
        "company_row": int(segment_profile.get("company_row") or 0),
        "segment_row": int(segment_profile.get("segment_row") or 0),
        "company_count": int(segment_profile.get("company_count") or 0),
        "segment_count": int(segment_profile.get("segment_count") or 0),
        "detected_segment_order": list(
            segment_profile.get("detected_segment_order") or []
        ),
        "segment_order": list(segment_profile.get("segment_order") or []),
        "segment_name_overrides": dict(
            segment_profile.get("segment_name_overrides") or {}
        ),
        "segment_assignment_overrides": dict(
            segment_profile.get("segment_assignment_overrides") or {}
        ),
        "segments": list(segment_profile.get("segments") or []),
        "assignments": list(segment_profile.get("assignments") or []),
        "error": str(segment_profile.get("error") or ""),
    }
    manifest_inputs = {
        key: mappings[key]
        for key in required_keys
        if isinstance(mappings.get(key), dict)
    }
    current_missing_keys = {
        " ".join(str(value or "").split()).casefold()
        for value in (function_lookup_profile.get("missing") or [])
        if " ".join(str(value or "").split()).strip()
    }
    active_ignored_taxonomy_functions = [
        " ".join(str(value or "").split()).strip()
        for value in options.ignored_aura_taxonomy_functions
        if " ".join(str(value or "").split()).strip().casefold() in current_missing_keys
    ]
    identity_payload = {
        "selected_sections": list(options.selected_sections),
        "taxonomy_sheet": taxonomy_sheet,
        "ignored_taxonomy_functions": active_ignored_taxonomy_functions,
        "segments": segment_payload,
        "function_lookup": {
            "function_label_sheet": function_lookup_profile.get("function_label_sheet"),
            "function_label_range": function_lookup_profile.get("function_label_range"),
            "fte_value_range": function_lookup_profile.get("fte_value_range"),
            "descriptor_column_count": function_lookup_profile.get("descriptor_column_count"),
            "matches": function_lookup_profile.get("matches"),
        },
        "master_labels": dict(master_label_profile.get("labels") or {}),
        "inputs": manifest_inputs,
    }
    identity_text = json.dumps(
        identity_payload,
        sort_keys=True,
        ensure_ascii=False,
    )
    digest = hashlib.sha256(identity_text.encode("utf-8")).hexdigest()[:12]
    manifest_path = manifest_dir / f"aura_input_map_{digest}.json"
    payload = {
        "schema_version": AURA_INPUT_MAP_SCHEMA_VERSION,
        "slide_family": "aura",
        "selected_sections": list(options.selected_sections),
        "requires_l0_quartiles": "l0" in options.selected_sections,
        "requires_l1_benchmarks": "l1" in options.selected_sections,
        "workbook_path": str(options.workbook.resolve()),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "taxonomy_sheet": taxonomy_sheet,
        "ignored_taxonomy_functions": active_ignored_taxonomy_functions,
        "taxonomy": {
            "sheet_name": taxonomy_sheet,
            "range_a1": taxonomy_range,
            "header_row": taxonomy_profile.get("header_row"),
            "l0_column": taxonomy_profile.get("l0_column"),
            "l1_column": taxonomy_profile.get("l1_column"),
            "header_prefix": taxonomy_profile.get("header_prefix"),
            "l0_header": taxonomy_profile.get("l0_header"),
            "l1_header": taxonomy_profile.get("l1_header"),
            "row_count": taxonomy_profile.get("row_count"),
        },
        "function_lookup": {
            "function_label_sheet": function_lookup_profile.get("function_label_sheet"),
            "function_label_range": function_lookup_profile.get("function_label_range"),
            "function_label_column": function_lookup_profile.get("function_label_column"),
            "fte_value_range": function_lookup_profile.get("fte_value_range"),
            "descriptor_column_count": function_lookup_profile.get("descriptor_column_count"),
            "value_column_count": function_lookup_profile.get("value_column_count"),
            "source_row_count": function_lookup_profile.get("source_row_count"),
            "matched_count": function_lookup_profile.get("matched_count"),
            "matches": function_lookup_profile.get("matches"),
            "extra_source": function_lookup_profile.get("extra_source"),
        },
        "master_labels": dict(master_label_profile.get("labels") or {}),
        "quartile_tables": dict(master_label_profile.get("quartile_tables") or {}),
        "segments": segment_payload,
        "inputs": manifest_inputs,
    }
    temporary_path = manifest_path.with_suffix(".json.writing")
    temporary_path.write_text(
        json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False),
        encoding="utf-8",
    )
    temporary_path.replace(manifest_path)
    return manifest_path.resolve()


# Legacy internal name retained for any external imports.
def _write_l1_input_manifest(options: BuildOptions) -> Path:
    return _write_aura_input_manifest(options)


def _read_sort_vector_from_mapping(
    workbook,
    mapping: dict[str, Any],
) -> list[Any]:
    from openpyxl.utils.cell import range_boundaries

    sheet_name = str(mapping.get("sheet_name") or "").strip()
    range_a1 = str(mapping.get("range_a1") or "").strip()
    if not sheet_name or not range_a1 or sheet_name not in workbook.sheetnames:
        return []
    min_col, min_row, max_col, max_row = range_boundaries(range_a1)
    if min_row != max_row:
        return []
    worksheet = workbook[sheet_name]
    return [
        worksheet.cell(row=min_row, column=column).value
        for column in range(min_col, max_col + 1)
    ]


def _write_company_sort_manifest(options: BuildOptions) -> Path:
    """Persist one universal company order reference for every bar-chart builder."""
    sort_order = str(options.l1_sort_order or L1_SORT_ORDER_DEFAULT).strip()
    if sort_order not in L1_SORT_ORDER_LABELS:
        raise ValueError(
            f"Unsupported company sort order {sort_order!r}. "
            f"Expected one of {list(L1_SORT_ORDER_LABELS)}."
        )
    direction = str(
        options.l1_segment_sort_direction
        or L1_SEGMENT_SORT_DIRECTION_DEFAULT
    ).strip()
    if direction not in L1_SEGMENT_SORT_DIRECTION_LABELS:
        raise ValueError(
            f"Unsupported within-segment order {direction!r}. "
            f"Expected one of {list(L1_SEGMENT_SORT_DIRECTION_LABELS)}."
        )

    mappings = dict(
        options.aura_input_mappings
        or options.l1_input_mappings
        or {}
    )
    company_mapping = dict(mappings.get("company_headers") or {})
    revenue_mapping = dict(mappings.get("revenue_values") or {})
    company_names: list[str] = []
    revenue_values: list[float | None] = []

    if company_mapping and revenue_mapping:
        from openpyxl import load_workbook

        try:
            workbook = load_workbook(
                options.workbook,
                read_only=True,
                data_only=True,
                keep_links=False,
            )
        except TypeError:
            workbook = load_workbook(
                options.workbook,
                read_only=True,
                data_only=True,
            )
        try:
            raw_companies = _read_sort_vector_from_mapping(
                workbook,
                company_mapping,
            )
            raw_revenue = _read_sort_vector_from_mapping(
                workbook,
                revenue_mapping,
            )
        finally:
            workbook.close()

        company_names = [str(value or "").strip() for value in raw_companies]
        if raw_revenue and len(raw_revenue) == len(company_names) + 1:
            raw_revenue = raw_revenue[1:]
        if len(raw_revenue) == len(company_names):
            for value in raw_revenue:
                try:
                    if value is None or str(value).strip() in {"", "-"}:
                        revenue_values.append(None)
                    else:
                        revenue_values.append(float(value))
                except (TypeError, ValueError):
                    revenue_values.append(None)
        else:
            revenue_values = [None] * len(company_names)

    profile = dict(options.l1_segment_profile or {})
    profile_assignments = [
        dict(item)
        for item in (profile.get("assignments") or [])
        if isinstance(item, dict)
    ]
    by_company = {
        re.sub(r"[^a-z0-9]+", "", str(item.get("company") or "").lower()): item
        for item in profile_assignments
        if str(item.get("company") or "").strip()
    }

    assignments: list[dict[str, Any]] = []
    for index, company in enumerate(company_names):
        key = re.sub(r"[^a-z0-9]+", "", company.lower())
        detected = dict(by_company.get(key) or {})
        assignments.append(
            {
                "company_index": index,
                "company": company,
                "revenue": (
                    revenue_values[index]
                    if index < len(revenue_values)
                    else detected.get("revenue")
                ),
                "segment": (
                    "Client"
                    if index == 0
                    else str(detected.get("segment") or "Unassigned")
                ),
                "segment_color": str(detected.get("segment_color") or ""),
            }
        )

    if sort_order in {"revenue_ascending", "revenue_descending", "segment"}:
        if not assignments or len(assignments) < 2:
            raise ValueError(
                "Company order requires the auto-detected Company names and Revenue row. "
                "Review the Aura Summary source or choose Sheet order."
            )
        if all(item.get("revenue") is None for item in assignments[1:]):
            raise ValueError(
                "Revenue-based company ordering was selected, but no peer revenue values "
                "could be read from the mapped Revenue row."
            )
    if sort_order == "segment":
        missing = [
            str(item.get("company") or "")
            for item in assignments[1:]
            if str(item.get("segment") or "").strip().lower()
            in {"", "client", "unassigned"}
        ]
        if missing:
            raise ValueError(
                "Sort by segment is selected, but no segment was detected for: "
                + ", ".join(missing)
            )

    payload = {
        "schema_version": 1,
        "sort_order": sort_order,
        "segment_sort_direction": direction,
        "selected_sections": list(options.selected_sections),
        "client_fixed_first": True,
        "source": {
            "company_sheet": str(company_mapping.get("sheet_name") or ""),
            "company_range": str(company_mapping.get("range_a1") or ""),
            "revenue_sheet": str(revenue_mapping.get("sheet_name") or ""),
            "revenue_range": str(revenue_mapping.get("range_a1") or ""),
            "segment_sheet": str(profile.get("sheet_name") or ""),
            "segment_range": str(profile.get("segment_range") or ""),
        },
        "segment_order": list(profile.get("segment_order") or []),
        "assignments": assignments,
    }
    identity = json.dumps(payload, sort_keys=True, ensure_ascii=False)
    digest = hashlib.sha256(identity.encode("utf-8")).hexdigest()[:12]
    manifest_dir = options.workbook.parent.parent / "input_maps"
    manifest_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = manifest_dir / f"company_sort_{digest}.json"
    temporary_path = manifest_path.with_suffix(".json.writing")
    temporary_path.write_text(
        json.dumps(payload, indent=2, sort_keys=True, ensure_ascii=False),
        encoding="utf-8",
    )
    temporary_path.replace(manifest_path)
    return manifest_path.resolve()


def build_backend_command(options: BuildOptions, paths: AppPaths) -> list[str]:
    """Build one backend command for any selected combination of slide types.

    Single-section builds continue to use their dedicated workflow. Any
    multi-section selection uses the integrated Aura runner with explicit skip
    flags, including ``--skip-l1`` when Aura L1 is not selected.
    """
    selected = _selected_section_set(options.selected_sections, options.mode)
    if not selected:
        raise ValueError("Select at least one slide type before generating.")

    if len(selected) == 1:
        only_section = next(iter(selected))
        workflow = {
            "summary": "aura-summary",
            "l0": "aura-l0",
            "l1": "aura-l1",
            "geocost": "geocost",
        }[only_section]
    else:
        # The integrated Aura runner is the section orchestrator. It now
        # supports independently skipping Summary, L0, L1, and Geo Cost.
        workflow = "aura-l1"

    command = [
        str(paths.python_exe),
        "-u",
        str(paths.backend_cli),
        workflow,
        "--excel",
        str(options.workbook),
        "--output-dir",
        str(options.output_dir),
        "--output-name",
        options.output_name,
        "--no-input-prompt",
    ]

    def append_geocost_options(*, integrated: bool) -> None:
        if integrated:
            command.extend(["--geocost-scope", "auto"])
            command.extend(["--geocost-label-threshold", "3"])
            command.extend(["--geocost-max-companies", "0"])
        else:
            command.extend(["--scope", "auto"])
            command.extend(["--label-threshold", "3"])
            command.extend(["--max-companies", "0"])
        if options.geocost_global_sheet.strip():
            command.extend(
                ["--global-geocost-sheet", options.geocost_global_sheet.strip()]
            )
        if options.geocost_us_sheet.strip():
            command.extend(
                ["--us-geocost-sheet", options.geocost_us_sheet.strip()]
            )

    if workflow == "aura-l1":
        if "summary" not in selected:
            command.append("--skip-summary")
        if "l0" not in selected:
            command.append("--skip-l0")
        if "l1" not in selected:
            command.append("--skip-l1")
        if "geocost" not in selected:
            command.append("--skip-geocost")
        else:
            append_geocost_options(integrated=True)
    elif workflow == "aura-l0" and options.only_l0.strip():
        command.extend(["--only-l0", options.only_l0.strip()])
    elif workflow == "geocost":
        append_geocost_options(integrated=False)

    # Summary, L0, and L1 all draw from the Aura source worksheet. Geo Cost is
    # the only standalone workflow that must not receive --sheet-name.
    if selected.intersection({"summary", "l0", "l1"}) and options.aura_sheet.strip():
        command.extend(["--sheet-name", options.aura_sheet.strip()])

    aura_selected = selected.intersection({"l0", "l1"})
    if aura_selected and options.taxonomy_sheet.strip():
        command.extend(["--taxonomy-sheet", options.taxonomy_sheet.strip()])

    if aura_selected:
        aura_mappings = dict(
            options.aura_input_mappings
            or options.l1_input_mappings
            or {}
        )
        if not aura_mappings:
            labels = ", ".join(
                DECK_SECTION_LABELS.get(key, key)
                for key in ("l0", "l1")
                if key in aura_selected
            )
            raise ValueError(
                f"Map the standardized Aura inputs before generating {labels}."
            )
        command.extend(
            ["--aura-input-map", str(_write_aura_input_manifest(options))]
        )

    if "geocost" in selected:
        if not options.geocost_input_mappings:
            raise ValueError(
                "Map Geo Cost table and Geo Cost data row before generating Geo Cost slides."
            )
        command.extend(
            ["--geocost-input-map", str(_write_geocost_input_manifest(options))]
        )

    if selected.intersection(SORTABLE_DECK_SECTIONS):
        sort_order = str(
            options.l1_sort_order or L1_SORT_ORDER_DEFAULT
        ).strip()
        if sort_order not in L1_SORT_ORDER_LABELS:
            raise ValueError(
                "Unsupported company sort order: "
                f"{sort_order!r}. Expected one of {list(L1_SORT_ORDER_LABELS)}."
            )
        direction = str(
            options.l1_segment_sort_direction
            or L1_SEGMENT_SORT_DIRECTION_DEFAULT
        ).strip()
        if direction not in L1_SEGMENT_SORT_DIRECTION_LABELS:
            raise ValueError(
                "Unsupported within-segment order: "
                f"{direction!r}. Expected one of "
                f"{list(L1_SEGMENT_SORT_DIRECTION_LABELS)}."
            )
        sort_manifest = _write_company_sort_manifest(options)
        command.extend(["--company-sort-order", sort_order])
        command.extend(["--company-segment-sort-direction", direction])
        command.extend(["--company-sort-manifest", str(sort_manifest)])

    if options.keep_temp and selected != {"summary"}:
        command.append("--keep-temp")
    if options.debug:
        command.append("--debug")

    return command

def command_for_display(command: Iterable[str]) -> str:
    values = list(command)
    if os.name == "nt":
        return subprocess.list2cmdline(values)
    import shlex

    return shlex.join(values)


def _slide_number(line: str, label: str) -> int | None:
    match = re.search(rf"\b{re.escape(label)}\s+(\d+)\s*:", line, flags=re.I)
    return int(match.group(1)) if match else None


def progress_from_log(
    line: str,
    current: int,
    mode: str,
) -> tuple[int, str | None]:
    """Translate known backend log markers into a smooth UI progress estimate."""
    lower = line.lower()
    stage: str | None = None

    start_match = re.search(r"\bSTART:\s*(.+?)\s*$", line)
    if start_match:
        stage = start_match.group(1).strip()

    value = current

    if mode == "summary":
        if "read source excel" in lower:
            value = max(value, 25)
        if "aura summary table rebuilt" in lower:
            value = max(value, 70)
        if "created aura summary slide deck" in lower:
            value = max(value, 96)

    elif mode == "l0":
        if "read source excel" in lower:
            value = max(value, 10)
        l0_number = _slide_number(line, "L0 slide")
        if l0_number is not None:
            value = max(value, min(82, 15 + l0_number * 16))
        if "save and close aura l0" in lower:
            value = max(value, 88)
        if "final aura l0" in lower or "created aura l0" in lower:
            value = max(value, 97)

    elif mode == "geocost":
        if "detected global geo cost sheet" in lower or "detected us geo cost sheet" in lower:
            value = max(value, 8)
        if "geocost scope=" in lower or "geo cost scope=" in lower:
            value = max(value, 14)
        geo_number = _slide_number(line, "Geo Cost slide")
        if geo_number is not None:
            value = max(value, min(84, 18 + geo_number * 8))
        if "applying final centered labels" in lower:
            value = max(value, 88)
        if "created global geo cost deck" in lower or "created us geo cost deck" in lower:
            value = max(value, 95)
        if "created combined deck" in lower:
            value = max(value, 98)

    else:
        # Complete deck and Aura L1-only both route through aura-l1.
        if "build aura summary" in lower:
            value = max(value, 4)
        if "created aura summary" in lower:
            value = max(value, 11)
        if "build aura l0" in lower:
            value = max(value, 14)

        l0_number = _slide_number(line, "L0 slide")
        if l0_number is not None:
            value = max(value, min(31, 15 + l0_number * 4))

        if "build aura l1" in lower:
            value = max(value, 33)

        l1_number = _slide_number(line, "Slide")
        if l1_number is not None and "l0 slide" not in lower and "geo cost slide" not in lower:
            # The stable deck currently has about 24 L1 slides. This remains a
            # graceful estimate when a workbook contains fewer or more.
            value = max(value, min(66, 34 + int((l1_number / 24) * 31)))

        if "build global geo cost" in lower or "build us geo cost" in lower:
            value = max(value, 68)

        geo_number = _slide_number(line, "Geo Cost slide")
        if geo_number is not None:
            value = max(value, min(88, 68 + int((geo_number / 12) * 18)))

        if "final live-deck label pass complete" in lower:
            value = max(value, 89)
        if "combine" in lower:
            value = max(value, 92)
        if "created combined deck" in lower:
            value = max(value, 97)
        if "finalize aura l0" in lower:
            value = max(value, 98)
        if "final aura l0" in lower:
            value = max(value, 99)

    if "failed:" in lower or "traceback" in lower:
        stage = "Build failed — review the latest log entries"

    return min(max(value, 1), 99), stage


def readable_presentation(path: Path, timeout_seconds: float = 60.0) -> bool:
    deadline = time.monotonic() + max(1.0, timeout_seconds)
    while time.monotonic() < deadline:
        try:
            if path.stat().st_size <= 0:
                raise OSError("empty file")
            with zipfile.ZipFile(path, "r") as archive:
                archive.read("[Content_Types].xml")
                if archive.testzip() is not None:
                    raise zipfile.BadZipFile("CRC failure")
            return True
        except (OSError, zipfile.BadZipFile, KeyError):
            time.sleep(0.5)
    return False


def output_paths_from_logs(logs: Iterable[str]) -> list[Path]:
    result: list[Path] = []
    for line in logs:
        for raw_path in WINDOWS_PPTX_PATTERN.findall(line):
            candidate = Path(raw_path.strip().rstrip(" ).,;"))
            if candidate.suffix.lower() in PRESENTATION_SUFFIXES:
                result.append(candidate)
    return result


def locate_generated_deck(
    output_dir: Path,
    output_name: str,
    started_at: float,
    logs: list[str],
) -> Path | None:
    candidates: list[Path] = []

    for candidate in reversed(output_paths_from_logs(logs)):
        if candidate.is_file():
            candidates.append(candidate)

    requested = output_dir / output_name
    if requested.is_file():
        candidates.append(requested)

    stem = requested.stem
    for suffix in PRESENTATION_SUFFIXES:
        try:
            for candidate in output_dir.glob(f"{stem}*{suffix}"):
                try:
                    if candidate.is_file() and candidate.stat().st_mtime >= started_at - 5:
                        candidates.append(candidate)
                except OSError:
                    pass
        except OSError:
            pass

    unique: dict[str, Path] = {}
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
        except OSError:
            continue
        unique[str(resolved).lower()] = resolved

    ordered = sorted(
        unique.values(),
        key=lambda path: path.stat().st_mtime if path.exists() else 0,
        reverse=True,
    )
    for candidate in ordered:
        if readable_presentation(candidate, timeout_seconds=60.0):
            return candidate
    return None


# -----------------------------------------------------------------------------
# Visual rendering helpers
# -----------------------------------------------------------------------------


def _elapsed_text(seconds: float) -> str:
    total_seconds = max(0, int(seconds))
    minutes, seconds = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)
    if hours:
        return f"{hours:d}h {minutes:02d}m {seconds:02d}s"
    if minutes:
        return f"{minutes:d}m {seconds:02d}s"
    return f"{seconds:d}s"


def render_progress_html(
    percent: int,
    state: str,
    stage: str,
    elapsed_seconds: float = 0.0,
) -> str:
    state_key = state.lower()
    state_label = {
        "ready": "READY",
        "running": "RUNNING",
        "success": "COMPLETE",
        "error": "FAILED",
    }.get(state_key, state.upper())

    safe_stage = html.escape(stage or "Waiting for a workbook")
    safe_elapsed = html.escape(_elapsed_text(elapsed_seconds))
    pct = min(max(int(percent), 0), 100)

    return f"""
    <div class="tc-progress-card">
      <div class="tc-progress-topline">
        <div>
          <div class="tc-console-label">CURRENT STAGE</div>
          <div class="tc-stage-text">{safe_stage}</div>
        </div>
        <div class="tc-progress-meta">
          <span class="tc-state-pill tc-state-{state_key}">{state_label}</span>
          <span class="tc-elapsed">{safe_elapsed}</span>
        </div>
      </div>
      <div class="tc-progress-track" role="progressbar" aria-valuenow="{pct}" aria-valuemin="0" aria-valuemax="100">
        <div class="tc-progress-fill tc-progress-{state_key}" style="width:{pct}%"></div>
      </div>
      <div class="tc-progress-caption"><span>{pct}%</span><span>Slide generation</span></div>
    </div>
    """


BUILD_TRACKER_OPERATIONS: dict[str, tuple[str, ...]] = {
    "summary": (
        "Read source",
        "Update summary table",
        "Save slide",
    ),
    "l0": (
        "Chart data",
        "Text + tagline",
        "Think-cell chart",
        "Settle axes",
        "Save + assemble",
    ),
    "l1": (
        "Chart data",
        "Text + tagline",
        "Think-cell chart",
        "Settle axes",
        "Save + assemble",
    ),
    "geocost": (
        "Read chart data",
        "Text + tagline",
        "Think-cell chart",
        "Assemble slide",
    ),
}


def _tracker_family(label: str) -> dict[str, Any]:
    return {
        "label": label,
        "state": "pending",
        "slides": {},
        "active_slide": "",
        "detail": "Waiting to start",
    }


def _tracker_slide(family_key: str, name: str) -> dict[str, Any]:
    return {
        "name": name,
        "state": "pending",
        "operations": {
            operation: "pending"
            for operation in BUILD_TRACKER_OPERATIONS.get(family_key, ())
        },
    }


def _tracker_add_slide(
    tracker: dict[str, Any],
    family_key: str,
    name: str,
) -> dict[str, Any] | None:
    clean_name = re.sub(r"\s+", " ", str(name or "").strip())
    family = tracker.get("families", {}).get(family_key)
    if family is None or not clean_name:
        return None
    slides = family["slides"]
    if clean_name not in slides:
        slides[clean_name] = _tracker_slide(family_key, clean_name)
    return slides[clean_name]


def _tracker_set_family_active(
    tracker: dict[str, Any], family_key: str, detail: str = ""
) -> None:
    family = tracker.get("families", {}).get(family_key)
    if family is None:
        return
    if family["state"] != "done":
        family["state"] = "active"
    if detail:
        family["detail"] = detail
    tracker["active_family"] = family_key


def _tracker_set_slide_active(
    tracker: dict[str, Any], family_key: str, name: str
) -> dict[str, Any] | None:
    slide = _tracker_add_slide(tracker, family_key, name)
    family = tracker.get("families", {}).get(family_key)
    if slide is None or family is None:
        return None
    _tracker_set_family_active(tracker, family_key, f"Building {name}")
    if slide["state"] != "done":
        slide["state"] = "active"
    family["active_slide"] = slide["name"]
    tracker["active_slide"] = slide["name"]
    tracker["message"] = f"{family.get('label', family_key)} · {slide['name']}"
    return slide


def _tracker_set_operation(
    tracker: dict[str, Any],
    family_key: str,
    operation: str,
    state: str,
) -> None:
    family = tracker.get("families", {}).get(family_key)
    if family is None:
        return
    slide_name = str(family.get("active_slide") or "")
    if not slide_name:
        return
    slide = family.get("slides", {}).get(slide_name)
    if slide is None:
        return
    operations = slide.get("operations", {})
    if operation not in operations:
        return
    if state == "active":
        for key, value in list(operations.items()):
            if value == "active":
                operations[key] = "pending"
    operations[operation] = state


def _tracker_complete_slide(tracker: dict[str, Any], family_key: str) -> None:
    family = tracker.get("families", {}).get(family_key)
    if family is None:
        return
    slide_name = str(family.get("active_slide") or "")
    if not slide_name:
        return
    slide = family.get("slides", {}).get(slide_name)
    if slide is None:
        return
    for operation in slide.get("operations", {}):
        slide["operations"][operation] = "done"
    slide["state"] = "done"
    family["detail"] = f"Completed {slide_name}"


def _tracker_complete_family(
    tracker: dict[str, Any],
    family_key: str,
    *,
    complete_pending: bool = False,
) -> None:
    family = tracker.get("families", {}).get(family_key)
    if family is None:
        return
    for slide in family.get("slides", {}).values():
        if complete_pending and slide.get("state") in {"pending", "active"}:
            for operation in slide.get("operations", {}):
                slide["operations"][operation] = "done"
            slide["state"] = "done"
            continue
        if slide.get("state") == "active":
            for operation in slide.get("operations", {}):
                if slide["operations"][operation] == "active":
                    slide["operations"][operation] = "done"
            if all(value == "done" for value in slide.get("operations", {}).values()):
                slide["state"] = "done"
    family["state"] = "done"
    family["detail"] = "Section complete"
    family["active_slide"] = ""


def _tracker_fail(tracker: dict[str, Any], message: str) -> None:
    tracker["state"] = "error"
    tracker["message"] = message
    active_family = str(tracker.get("active_family") or "")
    family = tracker.get("families", {}).get(active_family)
    if family is not None:
        family["state"] = "error"
        active_slide = str(family.get("active_slide") or "")
        if active_slide and active_slide in family.get("slides", {}):
            family["slides"][active_slide]["state"] = "error"


def _command_flag_value(command: list[str], flag: str) -> str:
    try:
        index = command.index(flag)
    except ValueError:
        return ""
    if index + 1 >= len(command):
        return ""
    return str(command[index + 1])


def _read_tracker_manifest(command: list[str], flag: str) -> dict[str, Any]:
    raw = _command_flag_value(command, flag)
    if not raw:
        return {}
    try:
        path = Path(raw)
        if path.is_file():
            payload = json.loads(path.read_text(encoding="utf-8"))
            return payload if isinstance(payload, dict) else {}
    except Exception:
        pass
    return {}


def initialize_build_tracker(
    options: BuildOptions,
    command: list[str],
) -> dict[str, Any]:
    selected = _selected_section_set(options.selected_sections, options.mode)
    tracker: dict[str, Any] = {
        "state": "running",
        "message": "Starting slide generation",
        "active_family": "",
        "active_slide": "",
        "families": {},
    }
    for key in ("summary", "l0", "l1", "geocost"):
        if key in selected:
            tracker["families"][key] = _tracker_family(
                DECK_SECTION_LABELS.get(key, key)
            )

    if "summary" in selected:
        _tracker_add_slide(tracker, "summary", "Aura Summary")

    aura_manifest = _read_tracker_manifest(command, "--aura-input-map")
    function_lookup = aura_manifest.get("function_lookup")
    matches = function_lookup.get("matches", []) if isinstance(function_lookup, dict) else []
    if isinstance(matches, list):
        if "l1" in selected:
            for item in matches:
                if not isinstance(item, dict):
                    continue
                name = str(
                    item.get("taxonomy_function")
                    or item.get("source_function")
                    or ""
                ).strip()
                if name:
                    _tracker_add_slide(tracker, "l1", name)
        if "l0" in selected:
            requested_l0 = re.sub(r"\s+", " ", str(options.only_l0 or "").strip()).casefold()
            seen_l0: set[str] = set()
            for item in matches:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("l0") or "").strip()
                key = re.sub(r"\s+", " ", name).casefold()
                if requested_l0 and key != requested_l0:
                    continue
                if name and key not in seen_l0:
                    seen_l0.add(key)
                    _tracker_add_slide(tracker, "l0", name)

    geocost_manifest = _read_tracker_manifest(command, "--geocost-input-map")
    if "geocost" in selected and geocost_manifest:
        layout = geocost_manifest.get("layout")
        scope = str(geocost_manifest.get("scope_hint") or "").strip().upper()
        if isinstance(layout, dict):
            names = layout.get("function_names") or []
            if isinstance(names, list):
                for raw_name in names:
                    name = str(raw_name or "").strip()
                    if not name:
                        continue
                    display = f"{scope} · {name}" if scope else name
                    _tracker_add_slide(tracker, "geocost", display)

    return tracker


def _tracker_find_slide_name(
    tracker: dict[str, Any], family_key: str, raw_name: str
) -> str:
    family = tracker.get("families", {}).get(family_key) or {}
    slides = family.get("slides", {}) or {}
    clean = re.sub(r"\s+", " ", str(raw_name or "").strip())
    if clean in slides:
        return clean
    folded = clean.casefold()
    for existing in slides:
        if existing.casefold() == folded:
            return existing
        if " · " in existing and existing.split(" · ", 1)[1].casefold() == folded:
            return existing
    return clean


def _tracker_stage_operation(family_key: str, stage_name: str) -> str:
    lower = stage_name.casefold()
    if family_key == "l1":
        if "write chart data workbook" in lower:
            return "Chart data"
        if "update title" in lower and "tagline" in lower:
            return "Text + tagline"
        if "update think-cell chart" in lower:
            return "Think-cell chart"
        if "settle chart axes" in lower:
            return "Settle axes"
        if (
            "save temporary ppt" in lower
            or "create combined deck" in lower
            or "append slide to combined deck" in lower
        ):
            return "Save + assemble"
    elif family_key == "l0":
        if "write l0 chart data workbook" in lower:
            return "Chart data"
        if "update l0 title" in lower:
            return "Text + tagline"
        if "update l0 think-cell chart" in lower:
            return "Think-cell chart"
        if "settle l0 chart axes" in lower:
            return "Settle axes"
        if (
            "save temporary aura l0 slide" in lower
            or "create l0 combined deck" in lower
            or "append slide to l0 combined deck" in lower
        ):
            return "Save + assemble"
    return ""


def update_build_tracker_from_log(
    tracker: dict[str, Any], line: str
) -> None:
    text = str(line or "")
    lower = text.casefold()

    if "traceback" in lower or " | error" in lower or "failed:" in lower:
        _tracker_fail(tracker, "Build failed — open Detailed build log")
        return

    if (
        "start: build aura summary" in lower
        or "workflow: aura summary" in lower
        or "building aura summary slide from template" in lower
    ):
        _tracker_set_family_active(tracker, "summary", "Building Aura Summary")
        _tracker_set_slide_active(tracker, "summary", "Aura Summary")
        _tracker_set_operation(tracker, "summary", "Read source", "active")
    if "start: build aura l0" in lower:
        _tracker_set_family_active(tracker, "l0", "Preparing Aura L0")
    if "start: build aura l1" in lower:
        _tracker_set_family_active(tracker, "l1", "Preparing Aura L1")
    if "start: build global geo cost" in lower or "start: build us geo cost" in lower:
        _tracker_set_family_active(tracker, "geocost", "Building Geo Cost")

    match = re.search(
        r"={4,}\s*L0 slide\s+\d+\s*:\s*(.*?)\s*=+\s*$",
        text,
        flags=re.I,
    )
    if match:
        name = _tracker_find_slide_name(tracker, "l0", match.group(1))
        _tracker_set_slide_active(tracker, "l0", name)
        return

    match = re.search(
        r"={4,}\s*Slide\s+\d+\s*:\s*(.*?)\s*/\s*(.*?)\s*=+\s*$",
        text,
        flags=re.I,
    )
    if match and "geo cost slide" not in lower and "l0 slide" not in lower:
        name = _tracker_find_slide_name(tracker, "l1", match.group(2))
        _tracker_set_slide_active(tracker, "l1", name)
        return

    match = re.search(
        r"={4,}\s*(GLOBAL|US)\s+Geo Cost slide\s+\d+\s*:\s*(.*?)\s*=+\s*$",
        text,
        flags=re.I,
    )
    if match:
        scope = match.group(1).upper()
        raw_name = match.group(2).strip()
        preferred = f"{scope} · {raw_name}"
        geocost_slides = (
            tracker.get("families", {})
            .get("geocost", {})
            .get("slides", {})
        )
        if preferred in geocost_slides:
            name = preferred
        elif raw_name in geocost_slides:
            name = raw_name
        else:
            name = _tracker_find_slide_name(tracker, "geocost", preferred)
        _tracker_set_slide_active(tracker, "geocost", name)
        _tracker_set_operation(tracker, "geocost", "Read chart data", "active")
        return

    if "aura summary table rebuilt" in lower:
        summary_family = tracker.get("families", {}).get("summary")
        if summary_family is not None and not summary_family.get("active_slide"):
            _tracker_set_slide_active(tracker, "summary", "Aura Summary")
        _tracker_set_operation(tracker, "summary", "Read source", "done")
        _tracker_set_operation(tracker, "summary", "Update summary table", "done")
        _tracker_set_operation(tracker, "summary", "Save slide", "active")
    if "created aura summary slide deck" in lower:
        _tracker_set_operation(tracker, "summary", "Save slide", "done")
        _tracker_complete_slide(tracker, "summary")
        _tracker_complete_family(tracker, "summary")
    if "done: build aura summary" in lower:
        _tracker_complete_family(tracker, "summary")

    stage_match = re.search(r"\b(START|DONE):\s*(.+?)(?:\s+\([0-9.]+s\))?\s*$", text, flags=re.I)
    if stage_match:
        phase = stage_match.group(1).casefold()
        stage_name = stage_match.group(2).strip()
        for family_key in ("l0", "l1"):
            family = tracker.get("families", {}).get(family_key)
            if (
                not family
                or family.get("state") == "done"
                or not family.get("active_slide")
            ):
                continue
            operation = _tracker_stage_operation(family_key, stage_name)
            if operation:
                _tracker_set_operation(
                    tracker,
                    family_key,
                    operation,
                    "active" if phase == "start" else "done",
                )
                if (
                    phase == "done"
                    and operation == "Save + assemble"
                    and (
                        "create combined deck" in stage_name.casefold()
                        or "append slide to combined deck" in stage_name.casefold()
                        or "create l0 combined deck" in stage_name.casefold()
                        or "append slide to l0 combined deck" in stage_name.casefold()
                    )
                ):
                    _tracker_complete_slide(tracker, family_key)

    if "created aura l0 section deck" in lower or "done: build aura l0 section shell" in lower:
        _tracker_complete_slide(tracker, "l0")
        _tracker_complete_family(tracker, "l0")
    if (
        ("created deck:" in lower and tracker.get("active_family") == "l1")
        or "done: build aura l1 slides" in lower
    ):
        _tracker_complete_slide(tracker, "l1")
        _tracker_complete_family(tracker, "l1")

    if "mapped source rows into the powerpoint template" in lower or (
        "scope:" in lower and "function:" in lower and "companies:" in lower
    ):
        _tracker_set_operation(tracker, "geocost", "Read chart data", "done")
        _tracker_set_operation(tracker, "geocost", "Text + tagline", "active")
    if "updated slide title/function text" in lower:
        _tracker_set_operation(tracker, "geocost", "Text + tagline", "done")
        _tracker_set_operation(tracker, "geocost", "Think-cell chart", "active")
    if "think-cell data update succeeded" in lower:
        _tracker_set_operation(tracker, "geocost", "Think-cell chart", "done")
        _tracker_set_operation(tracker, "geocost", "Assemble slide", "active")
    if "initialized live" in lower and "geo cost combined deck" in lower:
        _tracker_set_operation(tracker, "geocost", "Assemble slide", "done")
        _tracker_complete_slide(tracker, "geocost")
    if "appended slide from:" in lower:
        _tracker_set_operation(tracker, "geocost", "Assemble slide", "done")
        _tracker_complete_slide(tracker, "geocost")
    if "created global geo cost deck" in lower or "created us geo cost deck" in lower:
        _tracker_complete_slide(tracker, "geocost")
        family = tracker.get("families", {}).get("geocost")
        if family is not None:
            family["detail"] = "Geo Cost section built"

    if "all selected section decks are complete" in lower:
        for key in tracker.get("families", {}):
            _tracker_complete_family(tracker, key)
        tracker["message"] = "Finalizing and verifying the combined deck"


def _tracker_state_class(state: str) -> str:
    return state if state in {"pending", "active", "done", "error"} else "pending"


def render_build_tracker_html(
    tracker: dict[str, Any],
    elapsed_seconds: float = 0.0,
    *,
    final_state: str = "running",
    final_message: str = "",
) -> str:
    safe_elapsed = html.escape(_elapsed_text(elapsed_seconds))
    state = str(final_state or tracker.get("state") or "running").casefold()
    message = final_message or str(tracker.get("message") or "Building selected slides")
    if state == "success":
        message = final_message or "Deck generated successfully"
    elif state == "error":
        message = final_message or str(tracker.get("message") or "Build stopped")

    families = tracker.get("families", {}) or {}
    active_family_key = str(tracker.get("active_family") or "")
    visible_family_key = active_family_key if active_family_key in families else ""
    if not visible_family_key:
        for candidate_key, candidate_family in families.items():
            if str(candidate_family.get("state") or "pending") == "active":
                visible_family_key = candidate_key
                break
    if not visible_family_key:
        for candidate_key, candidate_family in families.items():
            if str(candidate_family.get("state") or "pending") == "pending":
                visible_family_key = candidate_key
                break
    if not visible_family_key and families:
        visible_family_key = next(reversed(families))

    family_chips: list[str] = []
    family_blocks: list[str] = []
    for family_key, family in families.items():
        family_state = _tracker_state_class(str(family.get("state") or "pending"))
        family_label = html.escape(str(family.get("label") or family_key))
        family_chips.append(
            f'<span class="tc-build-pill tc-build-{family_state}">{family_label}</span>'
        )

        if family_key != visible_family_key:
            continue

        slides = family.get("slides", {}) or {}
        slide_chips: list[str] = []
        completed_count = 0
        for slide_name, slide in slides.items():
            slide_state = _tracker_state_class(str(slide.get("state") or "pending"))
            if slide_state == "done":
                completed_count += 1
            slide_chips.append(
                '<span class="tc-build-pill tc-build-slide '
                f'tc-build-{slide_state}">{html.escape(str(slide_name))}</span>'
            )

        active_slide_name = str(family.get("active_slide") or "")
        active_slide = slides.get(active_slide_name) if active_slide_name else None
        operations_html = ""
        if active_slide is not None:
            operation_chips = []
            for operation, operation_state in active_slide.get("operations", {}).items():
                op_state = _tracker_state_class(str(operation_state or "pending"))
                operation_chips.append(
                    '<span class="tc-build-pill tc-build-operation '
                    f'tc-build-{op_state}">{html.escape(str(operation))}</span>'
                )
            operations_html = (
                '<div class="tc-build-operation-wrap">'
                f'<div class="tc-build-row-label">CURRENT SLIDE · {html.escape(active_slide_name)}</div>'
                f'<div class="tc-build-pill-row">{"".join(operation_chips)}</div>'
                '</div>'
            )

        total_count = len(slides)
        count_suffix = (
            f' · {completed_count} / {total_count} COMPLETE'
            if total_count
            else ""
        )
        slides_html = "".join(slide_chips) or '<span class="tc-build-empty">Waiting for slide list…</span>'
        family_blocks.append(
            '<div class="tc-build-family-panel">'
            f'<div class="tc-build-row-label">SLIDES / FUNCTIONS{count_suffix}</div>'
            f'<div class="tc-build-pill-row tc-build-function-row">{slides_html}</div>'
            f'{operations_html}'
            '</div>'
        )

    state_class = "success" if state == "success" else "error" if state == "error" else "running"
    status_label = "COMPLETE" if state == "success" else "FAILED" if state == "error" else "BUILDING"
    return (
        f'<div class="tc-build-tracker tc-build-tracker-{state_class}">'
        '<div class="tc-build-tracker-head">'
        '<div>'
        '<div class="tc-console-label">BUILD PROGRESS</div>'
        f'<div class="tc-build-message">{html.escape(message)}</div>'
        '</div>'
        '<div class="tc-build-head-meta">'
        f'<span class="tc-build-state tc-build-state-{state_class}">{status_label}</span>'
        f'<span class="tc-elapsed">{safe_elapsed}</span>'
        '</div></div>'
        '<div class="tc-build-row-label">SLIDE TYPES</div>'
        f'<div class="tc-build-pill-row tc-build-family-row">{"".join(family_chips)}</div>'
        f'{"".join(family_blocks)}'
        '</div>'
    )

def _log_line_class(line: str) -> str:
    lower = line.lower()
    if "traceback" in lower or " | error" in lower or "failed:" in lower:
        return "error"
    if " | warning" in lower or "warning" in lower:
        return "warning"
    if "done:" in lower or "succeeded" in lower or "successfully" in lower:
        return "success"
    if "start:" in lower or "==========" in line:
        return "stage"
    if not line.strip():
        return "blank"
    return "normal"


def render_log_html(logs: list[str]) -> str:
    if not logs:
        return """
        <div class="tc-log-shell tc-log-empty-shell">
          <div class="tc-empty-terminal">&gt;_</div>
          <div class="tc-empty-title">Ready to build</div>
          <div class="tc-empty-copy">Upload a source workbook and start the slide maker. Live backend activity will appear here.</div>
        </div>
        """

    visible_logs = logs[-MAX_VISIBLE_LOG_LINES:]
    rendered_lines: list[str] = []
    for line in visible_logs:
        css_class = _log_line_class(line)
        escaped = html.escape(line) if line else "&nbsp;"
        rendered_lines.append(
            f'<div class="tc-log-line tc-log-{css_class}">{escaped}</div>'
        )

    return (
        '<div class="tc-log-shell"><div class="tc-log-scroll">'
        + "".join(rendered_lines)
        + '</div></div>'
    )


def render_result_panel(
    placeholder,
    output_path: Path | None,
    logs: list[str],
    error_message: str | None = None,
) -> None:
    with placeholder.container():
        if output_path is not None and output_path.is_file():
            st.markdown(
                f"""
                <div class="tc-result-card tc-result-success">
                  <div class="tc-result-eyebrow">OUTPUT READY</div>
                  <div class="tc-result-title">{html.escape(output_path.name)}</div>
                  <div class="tc-result-path">{html.escape(str(output_path))}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            try:
                deck_bytes = output_path.read_bytes()
            except OSError as exc:
                st.error(f"The output file could not be read: {exc}")
            else:
                st.download_button(
                    "Download PowerPoint",
                    data=deck_bytes,
                    file_name=output_path.name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    ),
                    type="primary",
                    use_container_width=True,
                    key=f"download_deck_{output_path.name}_{output_path.stat().st_mtime_ns}",
                )

            if os.name == "nt":
                open_col, folder_col = st.columns(2)
                with open_col:
                    if st.button(
                        "Open deck",
                        use_container_width=True,
                        key=f"open_deck_{output_path.name}",
                    ):
                        os.startfile(str(output_path))  # type: ignore[attr-defined]
                with folder_col:
                    if st.button(
                        "Open Downloads",
                        use_container_width=True,
                        key=f"open_folder_{output_path.name}",
                    ):
                        os.startfile(str(output_path.parent))  # type: ignore[attr-defined]

        elif error_message:
            st.markdown(
                f"""
                <div class="tc-result-card tc-result-error">
                  <div class="tc-result-eyebrow">BUILD STOPPED</div>
                  <div class="tc-result-title">No PowerPoint was produced</div>
                  <div class="tc-result-path">{html.escape(error_message)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        if logs:
            st.download_button(
                "Download build log",
                data="\n".join(logs).encode("utf-8"),
                file_name=f"thinkcell_build_{datetime.now():%Y%m%d_%H%M%S}.log",
                mime="text/plain",
                use_container_width=True,
                key=f"download_log_{len(logs)}_{hash(tuple(logs[-5:]))}",
            )


# -----------------------------------------------------------------------------
# Backend execution with live logs
# -----------------------------------------------------------------------------


def _stream_process_output(process: subprocess.Popen[str], output_queue: queue.Queue) -> None:
    try:
        if process.stdout is None:
            return
        for raw_line in process.stdout:
            output_queue.put(raw_line.rstrip("\r\n"))
    finally:
        output_queue.put(None)


def run_backend(
    options: BuildOptions,
    console: ConsoleTargets,
    paths: AppPaths,
) -> tuple[int, list[str], Path | None]:
    command = build_backend_command(options, paths)
    tracker = initialize_build_tracker(options, command)

    st.session_state["last_command"] = command_for_display(command)
    st.session_state["last_output"] = None
    st.session_state["last_logs"] = []
    st.session_state["last_state"] = "running"
    st.session_state["last_build_tracker"] = tracker

    environment = os.environ.copy()
    environment["PYTHONUNBUFFERED"] = "1"
    environment["PYTHONIOENCODING"] = "utf-8"
    environment["PYTHONUTF8"] = "1"

    creation_flags = 0
    if os.name == "nt":
        creation_flags = int(getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200))

    started_at = time.time()
    logs: list[str] = []
    progress_value = 1
    current_stage = "Starting the slide-generation backend"
    process: subprocess.Popen[str] | None = None
    output_queue: queue.Queue = queue.Queue()

    console.progress.markdown(
        render_build_tracker_html(tracker, 0, final_state="running"),
        unsafe_allow_html=True,
    )
    console.log.markdown(render_log_html(logs), unsafe_allow_html=True)
    console.result.empty()

    try:
        process = subprocess.Popen(
            command,
            cwd=str(paths.backend_working_dir),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=environment,
            creationflags=creation_flags,
        )

        reader = threading.Thread(
            target=_stream_process_output,
            args=(process, output_queue),
            daemon=True,
        )
        reader.start()

        stream_closed = False
        last_render = 0.0
        while True:
            received_line = False
            try:
                item = output_queue.get(timeout=0.20)
                if item is None:
                    stream_closed = True
                else:
                    line = str(item)
                    logs.append(line)
                    received_line = True
                    progress_value, stage = progress_from_log(
                        line,
                        progress_value,
                        options.mode,
                    )
                    update_build_tracker_from_log(tracker, line)
                    if stage:
                        current_stage = stage
                        tracker["message"] = current_stage
            except queue.Empty:
                pass

            now = time.monotonic()
            if received_line or now - last_render >= 0.55:
                elapsed = time.time() - started_at
                console.progress.markdown(
                    render_build_tracker_html(
                        tracker,
                        elapsed,
                        final_state=(
                            "error"
                            if tracker.get("state") == "error"
                            else "running"
                        ),
                    ),
                    unsafe_allow_html=True,
                )
                console.log.markdown(render_log_html(logs), unsafe_allow_html=True)
                last_render = now

            if stream_closed and process.poll() is not None and output_queue.empty():
                break

        return_code = process.wait()
    except BaseException:
        if process is not None and process.poll() is None:
            try:
                process.terminate()
            except Exception:
                pass
        raise

    st.session_state["last_logs"] = logs
    st.session_state["last_build_tracker"] = tracker

    output_path: Path | None = None
    elapsed = time.time() - started_at
    if return_code == 0:
        current_stage = "Verifying the generated PowerPoint"
        tracker["message"] = current_stage
        console.progress.markdown(
            render_build_tracker_html(
                tracker, elapsed, final_state="running", final_message=current_stage
            ),
            unsafe_allow_html=True,
        )
        output_path = locate_generated_deck(
            options.output_dir,
            options.output_name,
            started_at,
            logs,
        )
        if output_path is not None:
            for family_key in tracker.get("families", {}):
                _tracker_complete_family(
                    tracker, family_key, complete_pending=True
                )
            tracker["state"] = "success"
            console.progress.markdown(
                render_build_tracker_html(
                    tracker,
                    elapsed,
                    final_state="success",
                    final_message="Deck generated successfully",
                ),
                unsafe_allow_html=True,
            )
            st.session_state["last_output"] = str(output_path)
            st.session_state["last_state"] = "success"
            st.session_state["last_build_tracker"] = tracker
            render_result_panel(console.result, output_path, logs)
        else:
            return_code = 1
            message = "The backend completed, but no readable PowerPoint output was found."
            _tracker_fail(tracker, message)
            console.progress.markdown(
                render_build_tracker_html(
                    tracker, elapsed, final_state="error", final_message=message
                ),
                unsafe_allow_html=True,
            )
            st.session_state["last_state"] = "error"
            st.session_state["last_build_tracker"] = tracker
            render_result_panel(console.result, None, logs, message)
    else:
        message = f"The backend exited with code {return_code}. Review the highlighted log lines."
        _tracker_fail(tracker, message)
        console.progress.markdown(
            render_build_tracker_html(
                tracker, elapsed, final_state="error", final_message=message
            ),
            unsafe_allow_html=True,
        )
        st.session_state["last_state"] = "error"
        st.session_state["last_build_tracker"] = tracker
        render_result_panel(console.result, None, logs, message)

    return return_code, logs, output_path




# -----------------------------------------------------------------------------
# Visual system and canvas orchestration
# -----------------------------------------------------------------------------


def apply_theme() -> None:
    """Apply a light, restrained desktop-style interface without boxy nesting."""
    st.markdown(
        """
        <style data-tc-ui-style="true">
          :root {
            --tc-accent: #c90000;
            --tc-accent-dark: #a80000;
            --tc-accent-soft: #fff3f3;
            --tc-ink: #17191c;
            --tc-text: #30343a;
            --tc-muted: #747b86;
            --tc-faint: #9aa1ab;
            --tc-page: #f5f6f8;
            --tc-surface: #ffffff;
            --tc-surface-soft: #f9fafb;
            --tc-line: #e5e8ec;
            --tc-line-strong: #d8dde3;
            --tc-green: #15825b;
            --tc-green-soft: #eaf8f1;
            --tc-amber: #b97300;
            --tc-shadow: 0 12px 34px rgba(20, 24, 31, 0.065);
          }

          html, body, [class*="css"] {
            font-family: "Segoe UI", Inter, Arial, sans-serif;
          }

          body { color: var(--tc-text); }
          .stApp { background: var(--tc-page); }

          [data-testid="stHeader"] {
            height: 0;
            background: transparent;
          }
          [data-testid="stToolbar"],
          [data-testid="stDecoration"],
          #MainMenu,
          footer {
            display: none !important;
          }

          .block-container {
            max-width: none;
            padding: 1rem 1.15rem 1.35rem;
          }

          [data-testid="stHorizontalBlock"] {
            align-items: flex-start;
            gap: 1rem;
          }

          .st-key-control_panel {
            position: sticky;
            top: 1rem;
            max-height: calc(100vh - 2rem);
            overflow-y: auto;
            background: var(--tc-surface);
            border: 1px solid var(--tc-line);
            border-radius: 19px;
            box-shadow: var(--tc-shadow);
            padding: 1.15rem 1.15rem 1rem;
            scrollbar-width: thin;
            scrollbar-color: #c8cdd4 transparent;
          }

          .st-key-control_panel::-webkit-scrollbar { width: 8px; }
          .st-key-control_panel::-webkit-scrollbar-thumb {
            background: #c8cdd4;
            border-radius: 999px;
          }

          .st-key-canvas {
            min-height: calc(100vh - 2rem);
            padding: 0.15rem 0.1rem 1.25rem 0.35rem;
          }

          .st-key-control_panel [data-testid="stVerticalBlock"] { gap: 0.58rem; }
          .st-key-canvas > [data-testid="stVerticalBlock"] { gap: 0.75rem; }

          .tc-panel-brand {
            display: flex;
            align-items: center;
            gap: 0.72rem;
            margin-bottom: 1.35rem;
          }

          .tc-panel-brand-mark {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 2.35rem;
            height: 2.35rem;
            border-radius: 11px;
            background: var(--tc-ink);
            color: #fff;
            font-size: 0.72rem;
            font-weight: 850;
            letter-spacing: 0.08em;
          }

          .tc-panel-brand-title {
            color: var(--tc-ink);
            font-size: 1.08rem;
            line-height: 1.15;
            font-weight: 800;
            letter-spacing: -0.015em;
          }

          .tc-panel-brand-subtitle {
            color: var(--tc-muted);
            font-size: 0.71rem;
            margin-top: 0.18rem;
          }

          .tc-panel-section-heading {
            display: flex;
            align-items: flex-start;
            gap: 0.66rem;
            margin-bottom: 0.15rem;
          }

          .tc-panel-section-number {
            color: var(--tc-accent);
            font-size: 0.62rem;
            line-height: 1.25rem;
            font-weight: 850;
            letter-spacing: 0.1em;
          }

          .tc-panel-section-title {
            color: var(--tc-ink);
            font-size: 0.88rem;
            font-weight: 780;
            line-height: 1.2;
          }

          .tc-panel-section-caption {
            color: var(--tc-muted);
            font-size: 0.68rem;
            line-height: 1.35;
            margin-top: 0.18rem;
          }

          .tc-panel-divider {
            height: 1px;
            background: var(--tc-line);
            margin: 0.7rem 0 0.65rem;
          }

          [data-testid="stFileUploaderDropzone"] {
            min-height: 5.7rem;
            padding: 0.7rem !important;
            background: var(--tc-surface-soft);
            border: 1px dashed #cbd1d8;
            border-radius: 12px;
          }

          [data-testid="stFileUploaderDropzone"]:hover {
            border-color: #aeb5be;
            background: #f7f8fa;
          }

          [data-testid="stFileUploaderDropzone"] button {
            border: 1px solid var(--tc-line-strong);
            border-radius: 9px;
            background: #fff;
            color: var(--tc-ink);
            font-weight: 720;
            box-shadow: none;
          }

          [data-testid="stFileUploaderDropzone"] small {
            color: var(--tc-muted) !important;
          }

          .tc-file-summary {
            display: grid;
            grid-template-columns: 2rem minmax(0, 1fr) 1.25rem;
            align-items: center;
            gap: 0.58rem;
            margin-top: 0.15rem;
            padding: 0.62rem 0.68rem;
            background: #fbfcfc;
            border: 1px solid var(--tc-line);
            border-radius: 11px;
          }

          .tc-file-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 2rem;
            height: 2rem;
            border-radius: 8px;
            background: #eaf7f0;
            color: #13734f;
            font-size: 0.62rem;
            font-weight: 850;
          }

          .tc-file-copy { min-width: 0; }
          .tc-file-name {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: var(--tc-ink);
            font-size: 0.73rem;
            font-weight: 720;
          }
          .tc-file-meta {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: var(--tc-muted);
            font-size: 0.64rem;
            margin-top: 0.12rem;
          }
          .tc-file-check {
            color: var(--tc-green);
            font-size: 0.86rem;
            font-weight: 850;
          }

          [data-testid="stRadio"] > div { gap: 0.1rem; }
          [data-testid="stRadio"] label {
            min-height: 2.05rem;
            padding: 0.38rem 0.48rem;
            border: 1px solid transparent;
            border-radius: 9px;
            transition: background 120ms ease, border-color 120ms ease;
          }
          [data-testid="stRadio"] label:hover {
            background: #f7f8fa;
            border-color: var(--tc-line);
          }
          [data-testid="stRadio"] label:has(input:checked) {
            background: var(--tc-accent-soft);
            border-color: #f1cccc;
          }
          [data-testid="stRadio"] label p {
            color: var(--tc-text);
            font-size: 0.78rem;
            font-weight: 650;
          }

          .tc-selection-note {
            margin-top: 0.16rem;
            padding: 0.58rem 0.65rem;
            border-left: 2px solid var(--tc-accent);
            background: #fafafa;
            color: var(--tc-muted);
            font-size: 0.68rem;
            line-height: 1.4;
          }

          [data-baseweb="input"] > div,
          [data-baseweb="base-input"],
          [data-baseweb="select"] > div {
            min-height: 2.55rem;
            background: #fff !important;
            border-color: var(--tc-line-strong) !important;
            border-radius: 10px !important;
            box-shadow: none !important;
          }

          [data-baseweb="input"] > div:focus-within,
          [data-baseweb="base-input"]:focus-within,
          [data-baseweb="select"] > div:focus-within {
            border-color: #b8bec6 !important;
            box-shadow: 0 0 0 3px rgba(201, 0, 0, 0.07) !important;
          }

          input, textarea { color: var(--tc-ink) !important; }

          .tc-output-location {
            display: flex;
            flex-direction: column;
            gap: 0.15rem;
            margin-top: 0.1rem;
            padding: 0.48rem 0.58rem;
            background: var(--tc-surface-soft);
            border-radius: 9px;
          }
          .tc-output-location-label {
            color: var(--tc-faint);
            font-size: 0.55rem;
            font-weight: 850;
            letter-spacing: 0.11em;
          }
          .tc-output-location-path {
            color: var(--tc-muted);
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.61rem;
            line-height: 1.35;
            overflow-wrap: anywhere;
          }

          [data-testid="stExpander"] {
            border: 0;
            border-top: 1px solid var(--tc-line);
            border-radius: 0;
            box-shadow: none;
            background: transparent;
          }
          [data-testid="stExpander"] summary {
            padding-left: 0;
            padding-right: 0;
            color: var(--tc-muted);
            font-size: 0.72rem;
            font-weight: 650;
          }

          div.stButton > button {
            min-height: 2.65rem;
            border-radius: 10px;
            font-weight: 760;
            box-shadow: none;
            transition: transform 100ms ease, background 100ms ease;
          }
          div.stButton > button[kind="primary"] {
            background: var(--tc-accent);
            border: 1px solid var(--tc-accent);
            color: #fff;
          }
          div.stButton > button[kind="primary"]:hover {
            background: var(--tc-accent-dark);
            border-color: var(--tc-accent-dark);
            transform: translateY(-1px);
          }
          div.stButton > button:disabled {
            transform: none;
            box-shadow: none;
          }

          .tc-panel-footnote {
            color: var(--tc-muted);
            font-size: 0.64rem;
            line-height: 1.4;
            text-align: center;
            margin-top: -0.15rem;
          }
          .tc-panel-footnote-ready {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.35rem;
            color: #47705f;
          }
          .tc-panel-footnote-ready span {
            width: 0.42rem;
            height: 0.42rem;
            border-radius: 999px;
            background: var(--tc-green);
          }
          .tc-panel-footnote-error { color: #a32626; }

          [data-testid="stSegmentedControl"] {
            margin: 0 0 0.7rem;
          }
          [data-testid="stSegmentedControl"] > div {
            width: fit-content;
            padding: 0.22rem;
            border: 1px solid var(--tc-line);
            border-radius: 11px;
            background: #eceff2;
          }
          [data-testid="stSegmentedControl"] button {
            min-height: 2rem;
            padding: 0.25rem 0.82rem;
            border: 0 !important;
            border-radius: 8px !important;
            color: var(--tc-muted);
            font-size: 0.72rem;
            font-weight: 720;
          }
          [data-testid="stSegmentedControl"] button[aria-pressed="true"] {
            background: #fff !important;
            color: var(--tc-ink) !important;
            box-shadow: 0 1px 4px rgba(20, 24, 31, 0.09);
          }

          .tc-main-header,
          .tc-workspace-header {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            margin-bottom: 0.45rem;
          }
          .tc-workbook-mode { display: none; }
          .tc-main-kicker {
            color: var(--tc-accent);
            font-size: 0.61rem;
            font-weight: 850;
            letter-spacing: 0.14em;
            text-transform: uppercase;
            margin-bottom: 0.28rem;
          }
          .tc-main-title {
            color: var(--tc-ink);
            font-size: clamp(1.55rem, 2.2vw, 2rem);
            line-height: 1.08;
            letter-spacing: -0.035em;
            font-weight: 820;
          }
          .tc-main-subtitle {
            max-width: 54rem;
            color: var(--tc-muted);
            font-size: 0.78rem;
            line-height: 1.45;
            margin-top: 0.28rem;
          }

          .tc-progress-card {
            background: var(--tc-surface);
            border: 1px solid var(--tc-line);
            border-radius: 14px;
            padding: 0.92rem 1rem 0.8rem;
            box-shadow: 0 5px 18px rgba(20, 24, 31, 0.035);
            margin-bottom: 0.72rem;
          }
          .tc-progress-topline {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 1rem;
          }
          .tc-console-label {
            color: var(--tc-faint);
            font-size: 0.56rem;
            font-weight: 850;
            letter-spacing: 0.12em;
            text-transform: uppercase;
            margin-bottom: 0.22rem;
          }
          .tc-stage-text {
            color: var(--tc-ink);
            font-size: 0.88rem;
            font-weight: 720;
            line-height: 1.35;
          }
          .tc-progress-meta {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            white-space: nowrap;
          }
          .tc-state-pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 4.65rem;
            padding: 0.28rem 0.5rem;
            border-radius: 999px;
            font-size: 0.56rem;
            font-weight: 850;
            letter-spacing: 0.08em;
          }
          .tc-state-ready { color: #56606c; background: #edf0f3; }
          .tc-state-running { color: #8c2100; background: #fff0e8; }
          .tc-state-success { color: #0d6847; background: var(--tc-green-soft); }
          .tc-state-error { color: #9f1d1d; background: #fdecec; }
          .tc-elapsed {
            color: var(--tc-muted);
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.68rem;
          }
          .tc-progress-track {
            position: relative;
            height: 0.48rem;
            margin-top: 0.82rem;
            overflow: hidden;
            background: #eceff2;
            border-radius: 999px;
          }
          .tc-progress-fill {
            height: 100%;
            border-radius: 999px;
            transition: width 260ms ease;
          }
          .tc-progress-running {
            background-color: var(--tc-accent);
            background-image: linear-gradient(
              45deg,
              rgba(255,255,255,.17) 25%,
              transparent 25%,
              transparent 50%,
              rgba(255,255,255,.17) 50%,
              rgba(255,255,255,.17) 75%,
              transparent 75%,
              transparent
            );
            background-size: 1.35rem 1.35rem;
            animation: tc-progress-stripes 1s linear infinite;
          }
          .tc-progress-success { background: var(--tc-green); }
          .tc-progress-error { background: var(--tc-accent); }
          .tc-progress-ready { background: #a0a7b0; }
          @keyframes tc-progress-stripes {
            from { background-position: 0 0; }
            to { background-position: 1.35rem 0; }
          }
          .tc-progress-caption {
            display: flex;
            justify-content: space-between;
            color: var(--tc-faint);
            font-size: 0.61rem;
            margin-top: 0.3rem;
          }

          .tc-build-tracker {
            background: #ffffff;
            border: 1px solid var(--tc-line);
            border-radius: 16px;
            padding: 0.9rem 1rem 0.95rem;
            box-shadow: 0 8px 24px rgba(20, 24, 31, 0.055);
            margin-bottom: 0.7rem;
          }
          .tc-build-tracker-head {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 1rem;
            padding-bottom: 0.7rem;
            margin-bottom: 0.68rem;
            border-bottom: 1px solid var(--tc-line);
          }
          .tc-build-message {
            color: var(--tc-ink);
            font-size: 0.92rem;
            font-weight: 750;
            line-height: 1.25;
            margin-top: 0.12rem;
          }
          .tc-build-head-meta {
            display: flex;
            align-items: center;
            gap: 0.55rem;
            flex: 0 0 auto;
          }
          .tc-build-state {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-height: 1.6rem;
            padding: 0.28rem 0.62rem;
            border-radius: 999px;
            font-size: 0.6rem;
            font-weight: 850;
            letter-spacing: 0.06em;
          }
          .tc-build-state-running { background: #17191c; color: #ffffff; }
          .tc-build-state-success { background: var(--tc-green); color: #ffffff; }
          .tc-build-state-error { background: var(--tc-accent); color: #ffffff; }
          .tc-build-row-label {
            color: #8c939d;
            font-size: 0.58rem;
            font-weight: 850;
            letter-spacing: 0.075em;
            margin: 0 0 0.38rem;
          }
          .tc-build-pill-row {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.42rem;
          }
          .tc-build-pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-height: 2rem;
            padding: 0.4rem 0.72rem;
            border-radius: 8px;
            border: 1px solid #292d33;
            background: #17191c;
            color: #ffffff;
            font-size: 0.69rem;
            font-weight: 760;
            line-height: 1.1;
            transition: background 180ms ease, border-color 180ms ease, box-shadow 180ms ease;
          }
          .tc-build-slide {
            min-height: 1.82rem;
            padding: 0.34rem 0.62rem;
            font-size: 0.65rem;
          }
          .tc-build-operation {
            min-height: 1.72rem;
            padding: 0.31rem 0.58rem;
            font-size: 0.62rem;
          }
          .tc-build-active {
            background: #17191c;
            border-color: var(--tc-accent);
            box-shadow: 0 0 0 2px rgba(201, 0, 0, 0.09);
            animation: tc-build-active-pulse 1.45s ease-in-out infinite;
          }
          @keyframes tc-build-active-pulse {
            0%, 100% { box-shadow: 0 0 0 2px rgba(201, 0, 0, 0.08); }
            50% { box-shadow: 0 0 0 4px rgba(201, 0, 0, 0.14); }
          }
          .tc-build-done {
            background: var(--tc-green);
            border-color: var(--tc-green);
            color: #ffffff;
          }
          .tc-build-error {
            background: var(--tc-accent);
            border-color: var(--tc-accent);
            color: #ffffff;
          }
          .tc-build-pending {
            background: #17191c;
            border-color: #292d33;
            color: #ffffff;
          }
          .tc-build-family-panel {
            margin-top: 0.78rem;
            padding-top: 0.72rem;
            border-top: 1px solid var(--tc-line);
          }
          .tc-build-operation-wrap {
            margin-top: 0.72rem;
            padding: 0.68rem 0.72rem 0.72rem;
            border-radius: 10px;
            background: #f7f8fa;
            border: 1px solid #eceff2;
          }
          .tc-build-empty {
            color: var(--tc-faint);
            font-size: 0.66rem;
            font-style: italic;
          }

          .tc-log-shell {
            overflow: hidden;
            background: #15181d;
            border: 1px solid #252a31;
            border-radius: 14px;
            box-shadow: 0 8px 24px rgba(20, 24, 31, 0.09);
          }
          .tc-log-scroll {
            height: 61vh;
            min-height: 410px;
            max-height: 700px;
            overflow-y: auto;
            padding: 0.88rem 0.95rem 1.05rem;
            color: #c7d0da;
            font-family: Consolas, "SFMono-Regular", "Liberation Mono", monospace;
            font-size: 0.69rem;
            line-height: 1.52;
            scrollbar-color: #4b535e #15181d;
            scrollbar-width: thin;
          }
          .tc-log-line {
            padding: 0.055rem 0;
            white-space: pre-wrap;
            word-break: break-word;
          }
          .tc-log-normal { color: #bdc6d0; }
          .tc-log-stage { color: #f6f7f9; font-weight: 700; }
          .tc-log-success { color: #64d3a0; }
          .tc-log-warning { color: #efbc68; }
          .tc-log-error { color: #ff7d7d; font-weight: 700; }
          .tc-log-blank { height: 0.38rem; }
          .tc-log-empty-shell {
            height: 61vh;
            min-height: 410px;
            max-height: 700px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 2rem;
            text-align: center;
          }
          .tc-empty-terminal {
            color: #ff6666;
            font-family: Consolas, monospace;
            font-size: 1.65rem;
            font-weight: 800;
            margin-bottom: 0.65rem;
          }
          .tc-empty-title {
            color: #f1f3f5;
            font-size: 0.9rem;
            font-weight: 750;
            margin-bottom: 0.28rem;
          }
          .tc-empty-copy {
            max-width: 28rem;
            color: #87919c;
            font-size: 0.74rem;
            line-height: 1.5;
          }

          .tc-preview-empty {
            height: 69vh;
            min-height: 500px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 2rem;
            text-align: center;
            background: var(--tc-surface);
            border: 1px dashed #cfd4da;
            border-radius: 16px;
          }
          .tc-preview-empty-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            min-width: 3.75rem;
            height: 2rem;
            padding: 0 0.65rem;
            margin-bottom: 0.75rem;
            border-radius: 8px;
            background: #eaf7f0;
            color: #0d6a47;
            font-family: Consolas, monospace;
            font-size: 0.68rem;
            font-weight: 850;
            letter-spacing: 0.08em;
          }
          .tc-preview-empty-title {
            color: var(--tc-ink);
            font-size: 0.96rem;
            font-weight: 780;
            margin-bottom: 0.22rem;
          }
          .tc-preview-empty-copy {
            color: var(--tc-muted);
            font-size: 0.76rem;
          }

          .tc-preview-metric {
            min-height: 3.15rem;
            padding: 0.52rem 0.68rem;
            background: var(--tc-surface);
            border: 1px solid var(--tc-line);
            border-radius: 11px;
            margin-bottom: 0.2rem;
          }
          .tc-preview-metric-label {
            color: var(--tc-faint);
            font-size: 0.53rem;
            font-weight: 850;
            letter-spacing: 0.1em;
            margin-bottom: 0.2rem;
          }
          .tc-preview-metric-value {
            color: var(--tc-ink);
            font-size: 0.83rem;
            font-weight: 760;
            line-height: 1.25;
          }
          .tc-preview-file-name {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .tc-static-preview-badge {
            min-height: 49px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            margin-top: 1.43rem;
            padding: 0.42rem 0.58rem;
            color: #667085;
            background: #f8fafb;
            border: 1px solid var(--tc-line);
            border-radius: 9px;
            font-size: 0.64rem;
            line-height: 1.35;
          }
          .tc-static-preview-badge strong {
            color: var(--tc-accent-dark);
            font-size: 0.55rem;
            letter-spacing: 0.09em;
          }
          .tc-result-card {
            padding: 0.78rem 0.86rem;
            margin: 0.72rem 0 0.55rem;
            border: 1px solid var(--tc-line);
            border-radius: 12px;
          }
          .tc-result-success {
            background: #f1faf6;
            border-color: #c2e3d2;
          }
          .tc-result-error {
            background: #fff5f5;
            border-color: #edc7c7;
          }
          .tc-result-eyebrow {
            color: var(--tc-muted);
            font-size: 0.54rem;
            font-weight: 850;
            letter-spacing: 0.11em;
            margin-bottom: 0.2rem;
          }
          .tc-result-title {
            color: var(--tc-ink);
            font-size: 0.88rem;
            font-weight: 780;
            margin-bottom: 0.18rem;
          }
          .tc-result-path {
            color: var(--tc-muted);
            font-size: 0.66rem;
            line-height: 1.4;
            word-break: break-word;
          }

          div.stDownloadButton > button {
            min-height: 2.55rem;
            border-radius: 10px;
            font-weight: 730;
          }
          div.stDownloadButton > button[kind="primary"] {
            background: var(--tc-accent);
            border-color: var(--tc-accent);
          }

          div[data-testid="stAlert"] {
            border-radius: 11px;
            border-width: 1px;
            box-shadow: none;
          }

          @media (max-width: 1080px) {
            .block-container { padding: 0.75rem; }
            [data-testid="stHorizontalBlock"] { gap: 0.7rem; }
            .st-key-control_panel { padding: 0.95rem; }
            .tc-log-scroll,
            .tc-log-empty-shell { min-height: 390px; }
          }

          @media (max-width: 820px) {
            .st-key-control_panel {
              position: static;
              max-height: none;
            }
            .st-key-canvas { padding-left: 0; }
            .tc-progress-topline { flex-direction: column; }
            .tc-progress-meta { width: 100%; justify-content: space-between; }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* Compact black control rail: flatter, tighter, and visually quieter. */
          [data-testid="stSidebar"] {
            display: flex !important;
            width: clamp(320px, 25vw, 410px) !important;
            min-width: clamp(320px, 25vw, 410px) !important;
            background: #111214 !important;
            border-right: 1px solid #292c31 !important;
            box-shadow: 8px 0 24px rgba(10, 12, 16, 0.10);
          }
          [data-testid="stSidebar"] > div:first-child {
            width: clamp(320px, 25vw, 410px) !important;
            background: transparent !important;
            scrollbar-width: thin;
            scrollbar-color: #3a3d43 transparent;
          }
          [data-testid="stSidebar"] > div:first-child::-webkit-scrollbar { width: 6px; }
          [data-testid="stSidebar"] > div:first-child::-webkit-scrollbar-thumb {
            background: #3a3d43;
            border-radius: 999px;
          }
          [data-testid="stSidebarCollapsedControl"] { display: none !important; }
          [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
            background: transparent !important;
          }
          [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.38rem; }
          [data-testid="stSidebar"] .st-key-control_panel {
            position: static;
            max-height: none;
            overflow: visible;
            padding: 0.68rem 0.82rem 0.9rem;
            background: transparent;
            border: 0;
            border-radius: 0;
            box-shadow: none;
          }
          [data-testid="stSidebar"] .st-key-control_panel > div { background: transparent; }
          [data-testid="stSidebar"] .st-key-control_panel > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-control_panel > div > [data-testid="stVerticalBlock"] {
            min-height: calc(100vh - 1.58rem);
          }
          [data-testid="stSidebar"] label,
          [data-testid="stSidebar"] p,
          [data-testid="stSidebar"] small { color: #e8eaed; }

          /* Brand */
          .tc-panel-brand {
            display: flex;
            align-items: center;
            gap: 0.58rem;
            margin: 0;
            padding: 0.08rem 0.04rem 0.68rem;
            border-bottom: 1px solid #282a2f;
          }
          .tc-panel-brand-mark {
            flex: 0 0 auto;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.9rem;
            height: 1.9rem;
            border-radius: 7px;
            background: #cf0a0a;
            color: #ffffff;
            font-size: 0.57rem;
            font-weight: 850;
            letter-spacing: 0.06em;
            box-shadow: none;
          }
          .tc-panel-brand-copy { min-width: 0; }
          .tc-panel-brand-title {
            color: #ffffff;
            font-size: 0.98rem;
            line-height: 1.08;
            font-weight: 780;
            letter-spacing: -0.02em;
          }
          .tc-panel-brand-subtitle {
            color: #7f8791;
            font-size: 0.5rem;
            line-height: 1.2;
            font-weight: 750;
            letter-spacing: 0.13em;
            margin-top: 0.18rem;
          }

          /* Section labels */
          .tc-panel-section-heading {
            display: grid;
            grid-template-columns: 1.5rem minmax(0, 1fr);
            align-items: start;
            gap: 0.38rem;
            margin: 0.08rem 0 0.16rem;
          }
          .tc-panel-section-number {
            display: block;
            width: auto;
            height: auto;
            margin: 0.12rem 0 0;
            padding: 0;
            border: 0;
            border-radius: 0;
            background: transparent;
            color: #ef5a5a;
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.52rem;
            line-height: 1.15;
            font-weight: 800;
            letter-spacing: 0.1em;
            box-shadow: none;
          }
          .tc-panel-section-title {
            color: #f2f3f4;
            font-size: 0.78rem;
            line-height: 1.2;
            font-weight: 720;
          }
          .tc-panel-section-caption {
            color: #777f89;
            font-size: 0.57rem;
            line-height: 1.3;
            margin-top: 0.1rem;
          }
          .tc-panel-divider {
            height: 1px;
            background: #282a2f;
            margin: 0.42rem 0 0.38rem;
          }

          /* Compact file uploader */
          [data-testid="stSidebar"] [data-testid="stFileUploader"] { margin-top: 0.02rem; }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            min-height: 3.25rem;
            padding: 0.42rem 0.5rem !important;
            background: #181a1e;
            border: 1px dashed #41454c;
            border-radius: 8px;
            transition: background 100ms ease, border-color 100ms ease;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:hover {
            background: #1d1f24;
            border-color: #626872;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] svg {
            width: 1rem !important;
            height: 1rem !important;
            color: #747b85 !important;
            fill: #747b85 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] {
            padding: 0 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span {
            color: #c7cbd1 !important;
            font-size: 0.63rem !important;
            line-height: 1.2 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
            display: none !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
            min-height: 1.82rem;
            padding: 0.22rem 0.52rem;
            background: #f4f5f6;
            color: #17191c !important;
            border: 0;
            border-radius: 6px;
            font-size: 0.61rem;
            font-weight: 720;
            box-shadow: none;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderFile"] {
            margin-top: 0.24rem;
            padding: 0.3rem 0.38rem;
            background: #181a1e;
            border: 1px solid #292c31;
            border-radius: 7px;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderFileName"] {
            color: #d7dade !important;
            font-size: 0.61rem !important;
          }
          .tc-upload-status {
            display: grid;
            grid-template-columns: 0.42rem minmax(0, 1fr) auto;
            align-items: center;
            gap: 0.38rem;
            min-width: 0;
            padding: 0.08rem 0.08rem 0;
            color: #8d949e;
            font-size: 0.56rem;
            line-height: 1.25;
          }
          .tc-upload-status-dot {
            width: 0.34rem;
            height: 0.34rem;
            border-radius: 999px;
            background: #3fc389;
            box-shadow: 0 0 0 2px rgba(63, 195, 137, 0.10);
          }
          .tc-upload-status-copy {
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .tc-upload-status-size {
            color: #686f79;
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.52rem;
          }

          /* Build scope: a compact 2 x 2 selector rather than four tall cards. */
          [data-testid="stSidebar"] [data-testid="stRadio"] > div[role="radiogroup"] {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.3rem;
          }
          [data-testid="stSidebar"] [data-testid="stRadio"] input[type="radio"],
          [data-testid="stSidebar"] [data-testid="stCheckbox"] input[type="checkbox"] {
            accent-color: #cc0000;
          }
          [data-testid="stSidebar"] [data-testid="stRadio"] label {
            position: relative;
            min-height: 2rem;
            margin: 0;
            padding: 0.32rem 0.36rem;
            background: #181a1e;
            border: 1px solid #292c31;
            border-radius: 7px;
            transition: background 100ms ease, border-color 100ms ease;
          }
          [data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
            background: #1e2025;
            border-color: #464b53;
          }
          [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
            background: #25191b;
            border-color: #7d3034;
            box-shadow: inset 2px 0 0 #d21616;
          }
          [data-testid="stSidebar"] [data-testid="stRadio"] label p {
            color: #d9dce0;
            font-size: 0.63rem;
            line-height: 1.18;
            font-weight: 650;
          }
          [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) p {
            color: #ffffff;
          }
          .tc-selection-note {
            display: flex;
            align-items: center;
            gap: 0.36rem;
            min-width: 0;
            margin: 0.18rem 0 0;
            padding: 0.08rem 0.05rem;
            background: transparent;
            border: 0;
            border-radius: 0;
            color: #7e858f;
            font-size: 0.56rem;
            line-height: 1.28;
          }
          .tc-selection-note span {
            flex: 0 0 auto;
            padding: 0.1rem 0.22rem;
            border: 1px solid #513033;
            border-radius: 4px;
            color: #e36c6c !important;
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.45rem;
            font-weight: 800;
            letter-spacing: 0.06em;
          }

          /* Inputs sit inside the rail instead of appearing as large white blocks. */
          [data-testid="stSidebar"] [data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-baseweb="base-input"],
          [data-testid="stSidebar"] [data-baseweb="select"] > div {
            min-height: 2.18rem;
            background: #181a1e !important;
            border: 1px solid #33373d !important;
            border-radius: 7px !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] [data-baseweb="input"] > div:focus-within,
          [data-testid="stSidebar"] [data-baseweb="base-input"]:focus-within,
          [data-testid="stSidebar"] [data-baseweb="select"] > div:focus-within {
            border-color: #8c383c !important;
            box-shadow: 0 0 0 2px rgba(204, 0, 0, 0.10) !important;
          }
          [data-testid="stSidebar"] input,
          [data-testid="stSidebar"] textarea {
            color: #eef0f2 !important;
            font-size: 0.66rem !important;
          }
          [data-testid="stSidebar"] input::placeholder,
          [data-testid="stSidebar"] textarea::placeholder {
            color: #686f79 !important;
            opacity: 1;
          }
          [data-testid="stSidebar"] [data-testid="stTextInput"] label p {
            color: #8e959f !important;
            font-size: 0.58rem;
          }
          .tc-output-location {
            display: grid;
            grid-template-columns: auto minmax(0, 1fr);
            align-items: center;
            gap: 0.42rem;
            margin-top: 0.15rem;
            padding: 0.05rem;
            background: transparent;
            border: 0;
            border-radius: 0;
          }
          .tc-output-location-label {
            color: #6f7680;
            font-size: 0.46rem;
            font-weight: 800;
            letter-spacing: 0.1em;
          }
          .tc-output-location-path {
            min-width: 0;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: #9097a1;
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.53rem;
            line-height: 1.25;
          }

          /* Advanced controls */
          [data-testid="stSidebar"] [data-testid="stExpander"] {
            margin-top: 0.02rem;
            border: 0;
            border-top: 1px solid #282a2f;
            border-radius: 0;
            background: transparent;
            box-shadow: none;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            min-height: 1.85rem;
            padding: 0.34rem 0 0.2rem;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary p {
            color: #9299a3 !important;
            font-size: 0.6rem;
            font-weight: 620;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary svg {
            color: #737a84 !important;
            fill: #737a84 !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpanderDetails"] {
            padding: 0.05rem 0 0.12rem;
          }
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label {
            min-height: 1.6rem;
            padding: 0.08rem 0;
          }
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
            color: #b8bdc4 !important;
            font-size: 0.6rem;
          }

          /* Persistent but restrained primary action. */
          [data-testid="stSidebar"] .st-key-panel_action {
            position: sticky;
            bottom: 0;
            z-index: 8;
            margin: auto -0.06rem 0;
            padding: 0.55rem 0.06rem 0.02rem;
            background: linear-gradient(
              180deg,
              rgba(17, 18, 20, 0) 0%,
              rgba(17, 18, 20, 0.96) 28%,
              #111214 48%
            );
          }
          [data-testid="stSidebar"] .st-key-panel_action [data-testid="stVerticalBlock"] {
            gap: 0.28rem;
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            min-height: 2.52rem;
            border-radius: 8px;
            background: #cc0000;
            border: 1px solid #cc0000;
            color: #ffffff;
            font-size: 0.77rem;
            font-weight: 760;
            letter-spacing: 0.005em;
            box-shadow: 0 5px 14px rgba(204, 0, 0, 0.16);
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"]:hover {
            background: #aa0000;
            border-color: #aa0000;
            transform: none;
          }
          [data-testid="stSidebar"] div.stButton > button:disabled {
            color: #6f757e !important;
            background: #202227 !important;
            border-color: #2e3136 !important;
            box-shadow: none !important;
            opacity: 1;
          }
          [data-testid="stSidebar"] .tc-panel-footnote {
            margin: 0;
            color: #707782;
            font-size: 0.54rem;
            line-height: 1.3;
            text-align: center;
          }
          [data-testid="stSidebar"] .tc-panel-footnote-ready {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.3rem;
            color: #82cbae;
          }
          [data-testid="stSidebar"] .tc-panel-footnote-ready span {
            width: 0.32rem;
            height: 0.32rem;
            border-radius: 999px;
            background: #3fc389;
          }
          [data-testid="stSidebar"] .tc-panel-footnote-error { color: #f18b8b; }
          [data-testid="stSidebar"] div[data-testid="stAlert"] {
            padding: 0.42rem 0.5rem;
            background: #281c1e;
            border-color: #603036;
            border-radius: 7px;
            font-size: 0.6rem;
          }

          /* Use nearly the full viewport for the workbook canvas. */
          [data-testid="stAppViewBlockContainer"],
          [data-testid="stMainBlockContainer"],
          .block-container {
            max-width: none;
            padding: 0.02rem 0.85rem 0.35rem !important;
          }
          .st-key-canvas {
            min-height: calc(100vh - 0.2rem);
            padding: 0;
          }
          .st-key-canvas > [data-testid="stVerticalBlock"] { gap: 0.38rem; }
          .st-key-canvas:has(.tc-workbook-mode) > [data-testid="stVerticalBlock"] { gap: 0.12rem; }
          .st-key-canvas:has(.tc-workbook-mode) [data-testid="stSegmentedControl"] {
            margin: 0 0 0.2rem;
          }
          .st-key-canvas:has(.tc-workbook-mode) [data-testid="stCustomComponentV1"],
          .st-key-canvas:has(.tc-workbook-mode) [data-testid="stCustomComponent"] {
            margin-bottom: 0 !important;
          }

          @media (max-width: 980px) {
            [data-testid="stSidebar"],
            [data-testid="stSidebar"] > div:first-child {
              width: 320px !important;
              min-width: 320px !important;
            }
            [data-testid="stAppViewBlockContainer"],
            [data-testid="stMainBlockContainer"],
            .block-container { padding: 0.02rem 0.6rem 0.3rem !important; }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* v5 sidebar scale: larger typography and controls without a rail scrollbar. */
          [data-testid="stSidebar"],
          [data-testid="stSidebar"] > div:first-child {
            width: clamp(340px, 25vw, 440px) !important;
            min-width: clamp(340px, 25vw, 440px) !important;
          }
          [data-testid="stSidebar"] > div:first-child,
          [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
            overflow-y: hidden !important;
            scrollbar-width: none !important;
          }
          [data-testid="stSidebar"] > div:first-child::-webkit-scrollbar,
          [data-testid="stSidebar"] [data-testid="stSidebarContent"]::-webkit-scrollbar {
            display: none !important;
            width: 0 !important;
          }
          [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.58rem;
          }
          [data-testid="stSidebar"] .st-key-control_panel {
            padding: 1rem 1.08rem 1.05rem;
            overflow: hidden !important;
          }
          [data-testid="stSidebar"] .st-key-control_panel > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-control_panel > div > [data-testid="stVerticalBlock"] {
            min-height: auto;
          }

          .tc-panel-brand {
            gap: 0.76rem;
            padding: 0.08rem 0.02rem 0.9rem;
            margin-bottom: 0.12rem;
          }
          .tc-panel-brand-mark {
            width: 2.35rem;
            height: 2.35rem;
            border-radius: 9px;
            font-size: 0.67rem;
          }
          .tc-panel-brand-title {
            font-size: 1.24rem;
            line-height: 1.08;
            font-weight: 800;
          }
          .tc-panel-brand-subtitle {
            margin-top: 0.24rem;
            font-size: 0.61rem;
            line-height: 1.25;
            letter-spacing: 0.14em;
          }

          .tc-panel-section-heading {
            grid-template-columns: 1.72rem minmax(0, 1fr);
            gap: 0.46rem;
            margin: 0.18rem 0 0.3rem;
          }
          .tc-panel-section-number {
            margin-top: 0.16rem;
            font-size: 0.6rem;
            line-height: 1.25;
          }
          .tc-panel-section-title {
            font-size: 0.9rem;
            line-height: 1.22;
            font-weight: 750;
          }
          .tc-panel-section-caption {
            margin-top: 0.16rem;
            font-size: 0.64rem;
            line-height: 1.38;
          }
          .tc-panel-divider {
            margin: 0.68rem 0 0.62rem;
          }

          [data-testid="stSidebar"] [data-testid="stFileUploader"] {
            margin-top: 0.08rem;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            min-height: 4.75rem;
            padding: 0.68rem 0.72rem !important;
            border-radius: 10px;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] svg {
            width: 1.25rem !important;
            height: 1.25rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span {
            font-size: 0.72rem !important;
            line-height: 1.3 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
            display: block !important;
            margin-top: 0.12rem;
            color: #777f89 !important;
            font-size: 0.57rem !important;
            line-height: 1.25 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
            min-height: 2.15rem;
            padding: 0.3rem 0.72rem;
            border-radius: 7px;
            font-size: 0.69rem;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderFile"] {
            margin-top: 0.34rem;
            padding: 0.42rem 0.5rem;
            border-radius: 8px;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderFileName"] {
            font-size: 0.68rem !important;
          }
          .tc-upload-status {
            grid-template-columns: 0.48rem minmax(0, 1fr) auto;
            gap: 0.44rem;
            padding: 0.18rem 0.06rem 0;
            font-size: 0.63rem;
          }
          .tc-upload-status-size { font-size: 0.57rem; }

          [data-testid="stSidebar"] [data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-baseweb="base-input"],
          [data-testid="stSidebar"] [data-baseweb="select"] > div {
            min-height: 2.72rem;
            border-radius: 9px !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] span,
          [data-testid="stSidebar"] input,
          [data-testid="stSidebar"] textarea {
            font-size: 0.76rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stTextInput"] label p,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] label p {
            font-size: 0.65rem !important;
          }
          .tc-selection-note {
            gap: 0.45rem;
            margin: 0.28rem 0 0.04rem;
            padding: 0.08rem 0.04rem;
            font-size: 0.64rem;
            line-height: 1.35;
          }
          .tc-selection-note span {
            padding: 0.12rem 0.28rem;
            font-size: 0.49rem;
          }

          .tc-output-location {
            gap: 0.5rem;
            margin-top: 0.24rem;
            padding: 0.08rem 0.02rem;
          }
          .tc-output-location-label { font-size: 0.51rem; }
          .tc-output-location-path {
            font-size: 0.59rem;
            line-height: 1.35;
          }

          [data-testid="stSidebar"] [data-testid="stExpander"] {
            margin-top: 0.18rem;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            min-height: 2.15rem;
            padding: 0.46rem 0 0.28rem;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary p,
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
            font-size: 0.67rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label {
            min-height: 1.85rem;
          }

          [data-testid="stSidebar"] .st-key-panel_action {
            position: static !important;
            margin: 0.5rem 0 0;
            padding: 0;
            background: none;
          }
          [data-testid="stSidebar"] .st-key-panel_action [data-testid="stVerticalBlock"] {
            gap: 0.4rem;
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            min-height: 3rem;
            border-radius: 9px;
            font-size: 0.86rem;
          }
          [data-testid="stSidebar"] .tc-panel-footnote {
            font-size: 0.61rem;
            line-height: 1.35;
          }

          /* Guided review action replaces the old Preview / Build Console toggle. */
          .st-key-workbook_review_action {
            padding-top: 0.18rem;
          }
          .st-key-workbook_review_action div.stButton > button {
            min-height: 2.55rem;
            border-radius: 9px;
            font-size: 0.75rem;
            white-space: nowrap;
          }

          /* Source-sheet overrides sit immediately below the console heading. */
          .st-key-source_override_row {
            margin: 0.04rem 0 0.72rem;
            padding: 0.72rem 0.82rem 0.78rem;
            background: #ffffff;
            border: 1px solid var(--tc-line);
            border-radius: 12px;
            box-shadow: 0 3px 12px rgba(20, 24, 31, 0.025);
          }
          .st-key-source_override_row > [data-testid="stVerticalBlock"],
          .st-key-source_override_row > div > [data-testid="stVerticalBlock"] {
            gap: 0.2rem;
          }
          .st-key-source_override_row [data-testid="stHorizontalBlock"] {
            gap: 0.85rem;
          }
          .st-key-source_override_row [data-testid="stSelectbox"] label p {
            color: #535b66;
            font-size: 0.68rem;
            font-weight: 750;
          }
          .st-key-source_override_row [data-baseweb="select"] > div {
            min-height: 2.7rem;
            border-radius: 9px !important;
          }
          .st-key-source_override_row [data-baseweb="select"] span {
            font-size: 0.75rem;
          }

          @media (max-width: 980px) {
            [data-testid="stSidebar"],
            [data-testid="stSidebar"] > div:first-child {
              width: 340px !important;
              min-width: 340px !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r7: remove Streamlit chrome and every false top-gap source. */
          /*
             apply_theme() is intentionally split across multiple CSS-only
             st.markdown() calls. Streamlit still gives each call a layout row,
             so their normal vertical gap looks like a stubborn top margin.
             Removing those wrappers preserves the CSS but removes the blank rows.
          */
          [data-testid="stElementContainer"]:has(style[data-tc-ui-style="true"]) {
            display: none !important;
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          [data-testid="stHeader"] {
            display: none !important;
            height: 0 !important;
          }
          [data-testid="stAppViewContainer"],
          [data-testid="stAppViewContainer"] > .main,
          [data-testid="stAppViewContainer"] section.main {
            padding-top: 0 !important;
            margin-top: 0 !important;
          }
          [data-testid="stSidebarContent"],
          [data-testid="stSidebarUserContent"] {
            padding-top: 0 !important;
            margin-top: 0 !important;
          }
          [data-testid="stAppViewBlockContainer"],
          [data-testid="stMainBlockContainer"],
          .block-container {
            max-width: none !important;
            padding: 0.12rem 0.62rem 0 !important;
            margin-top: 0 !important;
          }
          .st-key-canvas {
            min-height: 100vh !important;
            padding: 0 !important;
          }
          .st-key-canvas > [data-testid="stVerticalBlock"],
          .st-key-canvas > div > [data-testid="stVerticalBlock"] {
            gap: 0.22rem !important;
          }
          .tc-workspace-header,
          .tc-main-header {
            min-height: 2.45rem !important;
            display: flex !important;
            align-items: center !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .tc-main-title {
            margin: 0 !important;
            font-size: clamp(1.45rem, 1.75vw, 1.9rem) !important;
            line-height: 1.05 !important;
            letter-spacing: -0.035em !important;
          }
          .tc-main-kicker,
          .tc-main-subtitle { display: none !important; }

          .st-key-workbook_review_action {
            padding-top: 0 !important;
            margin: 0 !important;
          }
          .st-key-workbook_review_action div.stButton > button {
            min-height: 2.4rem !important;
            border-radius: 8px !important;
            font-size: 0.76rem !important;
          }
          .tc-input-help-source,
          .tc-finalize-help-source,
          .tc-finalize-blocker {
            display: none !important;
            width: 0 !important;
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: hidden !important;
            visibility: hidden !important;
          }
          .tc-finalize-blocker {
            min-height: 1.55rem;
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 0.42rem;
            margin: -0.04rem 0 0.08rem;
            padding: 0.18rem 0.18rem 0.18rem 0.6rem;
            color: #8f1015;
            font-size: 0.66rem;
            font-weight: 720;
            line-height: 1.25;
            text-align: right;
          }
          .tc-finalize-blocker-dot {
            flex: 0 0 auto;
            width: 0.42rem;
            height: 0.42rem;
            border-radius: 999px;
            background: #cc0000;
            box-shadow: 0 0 0 3px rgba(204,0,0,0.08);
          }

          /* One compact split control per source: preview area + integrated menu. */
          .st-key-source_selector_row {
            margin: 0 0 0.08rem !important;
            padding: 0 !important;
            background: transparent !important;
            border: 0 !important;
            border-radius: 0 !important;
            box-shadow: none !important;
          }
          .st-key-source_selector_row > [data-testid="stVerticalBlock"],
          .st-key-source_selector_row > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
          }
          .st-key-source_selector_row [data-testid="stCustomComponentV1"],
          .st-key-source_selector_row [data-testid="stCustomComponent"],
          .st-key-source_selector_row iframe {
            display: block !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            border: 0 !important;
            background: transparent !important;
          }
          .st-key-source_selector_row iframe {
            min-height: 48px !important;
          }

          /* Aura L1 named inputs now live in the left control rail. */
          .tc-sidebar-l1-heading {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 0.55rem;
            margin: 0.62rem 0 0.32rem;
            padding-top: 0.56rem;
            border-top: 1px solid #292d33;
          }
          .tc-sidebar-l1-title {
            color: #f4f5f6;
            font-size: 0.72rem;
            font-weight: 820;
            line-height: 1.2;
          }
          .tc-sidebar-l1-copy {
            margin-top: 0.12rem;
            color: #858d98;
            font-size: 0.55rem;
            line-height: 1.3;
          }
          .tc-sidebar-l1-count {
            flex: 0 0 auto;
            min-width: 2.42rem;
            padding: 0.18rem 0.36rem;
            border: 1px solid #3b4149;
            border-radius: 999px;
            background: #171a1e;
            color: #aab1bb;
            font-size: 0.55rem;
            font-weight: 800;
            text-align: center;
          }
          .tc-sidebar-l1-count.is-ready {
            border-color: #2e765d;
            background: #14271f;
            color: #79d2ad;
          }
          [data-testid="stSidebar"] .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
            gap: 0.2rem !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] {
            min-width: 0;
            margin: 0 !important;
            padding: 0 !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] > div > [data-testid="stVerticalBlock"] {
            gap: 0.02rem !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] div.stButton > button {
            min-height: 2.04rem !important;
            padding: 0.28rem 0.56rem !important;
            border-radius: 7px !important;
            font-size: 0.65rem !important;
            font-weight: 730 !important;
            line-height: 1.15 !important;
            text-align: left !important;
            justify-content: flex-start !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"] {
            background: #ffffff !important;
            border-color: #d5d9df !important;
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"] * {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"]:hover {
            background: #f1f3f5 !important;
            border-color: #aeb5bf !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] {
            background: #c90000 !important;
            border-color: #e32a2a !important;
            color: #ffffff !important;
            box-shadow: 0 0 0 1px rgba(255,255,255,0.04) inset !important;
          }
          [data-testid="stSidebar"] [class*="st-key-sidebar_l1_input_"] div.stButton > button:disabled {
            opacity: 0.48 !important;
          }
          .tc-sidebar-l1-range {
            width: 100%;
            padding: 0.08rem 0.5rem 0.11rem;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: #666e79;
            font-family: Consolas, "SFMono-Regular", monospace;
            font-size: 0.5rem;
            line-height: 1.2;
          }
          .tc-sidebar-l1-range.is-mapped { color: #7fc9ab; }
          .tc-sidebar-l1-status {
            margin: 0.2rem 0 0.08rem;
            padding: 0.32rem 0.48rem;
            border-left: 2px solid #c90000;
            background: #15181c;
            color: #929aa5;
            font-size: 0.56rem;
            line-height: 1.3;
          }
          .tc-sidebar-l1-status.is-ready {
            border-left-color: #2b9a70;
            background: #14221d;
            color: #79cbaa;
          }
          .tc-sidebar-l1-status.is-error {
            border-left-color: #e34d4d;
            background: #251719;
            color: #f0a0a0;
          }


          /* Preview-window controls remain available but no longer consume a full row. */
          .st-key-canvas [data-testid="stExpander"] {
            margin: 0 0 0.1rem !important;
            border: 1px solid #e1e5e9 !important;
            border-radius: 8px !important;
            background: #ffffff !important;
            box-shadow: none !important;
          }
          .st-key-canvas [data-testid="stExpander"] summary {
            min-height: 2rem !important;
            padding: 0.32rem 0.62rem !important;
          }
          .st-key-canvas [data-testid="stExpander"] summary p {
            font-size: 0.68rem !important;
            font-weight: 700 !important;
          }
          .st-key-canvas [data-testid="stExpanderDetails"] {
            padding: 0.25rem 0.55rem 0.5rem !important;
          }
          .st-key-canvas [data-testid="stNumberInput"] label p {
            font-size: 0.63rem !important;
          }
          .st-key-canvas [data-testid="stNumberInput"] [data-baseweb="input"] > div,
          .st-key-canvas [data-testid="stNumberInput"] [data-baseweb="base-input"] {
            min-height: 2.15rem !important;
          }
          .st-key-canvas [data-testid="stCustomComponentV1"],
          .st-key-canvas [data-testid="stCustomComponent"] {
            margin-top: 0 !important;
            margin-bottom: 0 !important;
          }
          .st-key-canvas div[data-testid="stAlert"] {
            margin: 0.08rem 0 !important;
            padding: 0.42rem 0.6rem !important;
          }

          /* r6 sidebar: fewer labels, cleaner grouping, no scroll rail. */
          [data-testid="stSidebar"],
          [data-testid="stSidebar"] > div:first-child {
            width: clamp(350px, 25vw, 430px) !important;
            min-width: clamp(350px, 25vw, 430px) !important;
            background: #101214 !important;
          }
          [data-testid="stSidebar"] > div:first-child,
          [data-testid="stSidebar"] [data-testid="stSidebarContent"] {
            overflow: hidden !important;
            scrollbar-width: none !important;
          }
          [data-testid="stSidebar"] > div:first-child::-webkit-scrollbar,
          [data-testid="stSidebar"] [data-testid="stSidebarContent"]::-webkit-scrollbar {
            display: none !important;
            width: 0 !important;
          }
          [data-testid="stSidebar"] .st-key-control_panel {
            height: 100vh !important;
            min-height: 100vh !important;
            max-height: 100vh !important;
            overflow: hidden !important;
            margin: 0 !important;
            padding: 1.15rem 1.28rem 1rem !important;
            background: #101214 !important;
            border: 0 !important;
            border-radius: 0 !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] .st-key-control_panel > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-control_panel > div > [data-testid="stVerticalBlock"] {
            min-height: 100% !important;
            gap: 0 !important;
          }
          .tc-panel-section-heading,
          .tc-panel-divider,
          .tc-selection-note { display: none !important; }

          .tc-panel-brand {
            gap: 0.78rem !important;
            margin: 0 !important;
            padding: 0 0 0.15rem !important;
          }
          .tc-panel-brand-mark {
            width: 2.5rem !important;
            height: 2.5rem !important;
            border-radius: 9px !important;
            font-size: 0.72rem !important;
          }
          .tc-panel-brand-title {
            font-size: 1.3rem !important;
            line-height: 1.05 !important;
            font-weight: 800 !important;
          }
          .tc-panel-brand-subtitle {
            margin-top: 0.24rem !important;
            font-size: 0.61rem !important;
            line-height: 1.2 !important;
            letter-spacing: 0.15em !important;
          }

          [data-testid="stSidebar"] .st-key-panel_source {
            margin-top: 1.35rem !important;
          }
          [data-testid="stSidebar"] .st-key-panel_build,
          [data-testid="stSidebar"] .st-key-panel_output {
            margin-top: 1.3rem !important;
          }
          [data-testid="stSidebar"] .st-key-panel_source > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_source > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > div > [data-testid="stVerticalBlock"] {
            gap: 0.38rem !important;
          }
          .tc-panel-field-label {
            margin: 0 0 0.12rem;
            color: #f2f4f6;
            font-size: 0.76rem;
            line-height: 1.2;
            font-weight: 760;
          }

          [data-testid="stSidebar"] [data-testid="stFileUploader"] {
            margin: 0 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            min-height: 4.7rem !important;
            padding: 0.65rem 0.72rem !important;
            background: #171a1e !important;
            border: 1px dashed #41464f !important;
            border-radius: 10px !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"]:hover {
            border-color: #686f79 !important;
            background: #1a1e22 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span {
            color: #f3f4f6 !important;
            font-size: 0.73rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
            color: #8b929b !important;
            font-size: 0.58rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
            min-height: 2.15rem !important;
            padding: 0.28rem 0.72rem !important;
            color: #101214 !important;
            background: #f5f6f7 !important;
            border-color: #f5f6f7 !important;
            border-radius: 7px !important;
            font-size: 0.69rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderFile"] {
            margin-top: 0.38rem !important;
            padding: 0.48rem 0.55rem !important;
            background: #1a1d22 !important;
            border-color: #343943 !important;
            border-radius: 8px !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderFileName"] {
            color: #f1f3f5 !important;
            font-size: 0.7rem !important;
          }
          .tc-upload-status {
            display: flex !important;
            align-items: center !important;
            gap: 0.42rem !important;
            margin: 0.08rem 0 0 !important;
            padding: 0 !important;
            color: #8d949d !important;
            font-size: 0.62rem !important;
          }
          .tc-upload-status-size { display: none !important; }

          [data-testid="stSidebar"] [data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-baseweb="base-input"],
          [data-testid="stSidebar"] [data-baseweb="select"] > div {
            min-height: 2.72rem !important;
            background: #191c20 !important;
            border-color: #353a42 !important;
            border-radius: 9px !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] [data-baseweb="input"] > div:focus-within,
          [data-testid="stSidebar"] [data-baseweb="base-input"]:focus-within,
          [data-testid="stSidebar"] [data-baseweb="select"] > div:focus-within {
            border-color: #737a84 !important;
            box-shadow: 0 0 0 1px #737a84 !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] span,
          [data-testid="stSidebar"] input,
          [data-testid="stSidebar"] textarea {
            color: #f3f4f6 !important;
            -webkit-text-fill-color: #f3f4f6 !important;
            font-size: 0.76rem !important;
          }
          [data-testid="stSidebar"] input::placeholder,
          [data-testid="stSidebar"] textarea::placeholder {
            color: #777e87 !important;
            -webkit-text-fill-color: #777e87 !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] svg {
            fill: #c9cdd2 !important;
          }

          .tc-output-location {
            display: flex !important;
            align-items: center !important;
            gap: 0.42rem !important;
            margin: 0.08rem 0 0 !important;
            padding: 0 !important;
          }
          .tc-output-location-label {
            color: #cc4a4a !important;
            font-size: 0.5rem !important;
            font-weight: 800 !important;
            letter-spacing: 0.12em !important;
          }
          .tc-output-location-path {
            min-width: 0;
            overflow: hidden;
            color: #858c95 !important;
            font-size: 0.59rem !important;
            text-overflow: ellipsis;
            white-space: nowrap;
          }

          [data-testid="stSidebar"] [data-testid="stExpander"] {
            margin-top: 1.05rem !important;
            background: transparent !important;
            border: 1px solid #2b2f35 !important;
            border-radius: 9px !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            min-height: 2.35rem !important;
            padding: 0.48rem 0.65rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary p,
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
            color: #b7bcc3 !important;
            font-size: 0.67rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpanderDetails"] {
            padding: 0.2rem 0.62rem 0.58rem !important;
          }

          [data-testid="stSidebar"] .st-key-panel_action {
            position: static !important;
            margin: 1.25rem 0 0 !important;
            padding: 0 !important;
            background: transparent !important;
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            min-height: 2.95rem !important;
            background: #c90000 !important;
            border-color: #c90000 !important;
            border-radius: 9px !important;
            color: #ffffff !important;
            font-size: 0.84rem !important;
            font-weight: 760 !important;
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"]:hover {
            background: #aa0000 !important;
            border-color: #aa0000 !important;
          }
          [data-testid="stSidebar"] div.stButton > button:disabled {
            color: #9ca2aa !important;
            background: #202329 !important;
            border-color: #30343b !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] .tc-panel-footnote {
            margin-top: 0.42rem !important;
            color: #777e87 !important;
            font-size: 0.61rem !important;
            line-height: 1.35 !important;
            text-align: center !important;
          }
          [data-testid="stSidebar"] .tc-panel-footnote-error {
            color: #f08b8b !important;
          }

          /* Every control keeps its own visible label. */
          [data-testid="stSidebar"] [data-testid="stFileUploader"] > label,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] > label,
          [data-testid="stSidebar"] [data-testid="stTextInput"] > label {
            display: flex !important;
            align-items: center !important;
            min-height: 1rem !important;
            margin: 0 0 0.28rem !important;
            padding: 0 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploader"] > label p,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] > label p,
          [data-testid="stSidebar"] [data-testid="stTextInput"] > label p {
            color: #f0f2f4 !important;
            font-size: 0.72rem !important;
            line-height: 1.25 !important;
            font-weight: 720 !important;
          }
          [data-testid="stSidebar"] .st-key-panel_build > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > div > [data-testid="stVerticalBlock"] {
            gap: 0.68rem !important;
          }

          /* r8: explicit field labels and guaranteed readable dark controls. */
          [data-testid="stSidebar"] .tc-panel-field-label {
            display: flex !important;
            align-items: baseline !important;
            justify-content: space-between !important;
            gap: 0.65rem !important;
            min-height: 1.05rem !important;
            margin: 0 0 0.32rem !important;
            padding: 0 0.02rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-spaced {
            margin-top: 0.76rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-text {
            color: #f5f6f7 !important;
            font-size: 0.76rem !important;
            line-height: 1.25 !important;
            font-weight: 730 !important;
            letter-spacing: 0.005em !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-hint {
            flex: 0 0 auto !important;
            color: #7f8791 !important;
            font-size: 0.54rem !important;
            line-height: 1.2 !important;
            font-weight: 760 !important;
            letter-spacing: 0.09em !important;
            text-transform: uppercase !important;
          }

          /* Streamlit/BaseWeb control markup varies by release. Target both the
             wrapper and every known inner control surface. */
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"],
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"] > div {
            background-color: #1a1d22 !important;
            background-image: none !important;
            border-color: #3a3f48 !important;
            color: #f5f6f7 !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"] {
            min-height: 2.75rem !important;
            border: 1px solid #3a3f48 !important;
            border-radius: 9px !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within > div,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"]:focus,
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"]:focus-within,
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"]:focus-within {
            border-color: #6f7782 !important;
            box-shadow: 0 0 0 1px #6f7782 !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] *,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"] *,
          [data-testid="stSidebar"] [data-testid="stTextInput"] input {
            color: #f5f6f7 !important;
            -webkit-text-fill-color: #f5f6f7 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [data-testid="stTextInput"] input {
            background: transparent !important;
            caret-color: #ffffff !important;
          }
          [data-testid="stSidebar"] [data-testid="stTextInput"] input::placeholder {
            color: #7f8791 !important;
            -webkit-text-fill-color: #7f8791 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] svg {
            color: #d5d9de !important;
            fill: #d5d9de !important;
          }

          /* r9: sidebar controls use a white surface with unequivocally dark text.
             This deliberately follows the control surface Streamlit is actually
             rendering in the current desktop build instead of fighting it with
             a dark BaseWeb skin. */
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"],
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"],
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"] > div {
            background: #ffffff !important;
            background-color: #ffffff !important;
            background-image: none !important;
            border-color: #d5d9df !important;
            color: #111418 !important;
            color-scheme: light !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"] *,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"] *,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] span,
          [data-testid="stSidebar"] [data-testid="stTextInput"] input,
          [data-testid="stSidebar"] [data-testid="stTextInput"] input:disabled {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [data-testid="stTextInput"] input {
            background: transparent !important;
            caret-color: #111418 !important;
          }
          [data-testid="stSidebar"] [data-testid="stTextInput"] input::placeholder {
            color: #8a919b !important;
            -webkit-text-fill-color: #8a919b !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] svg {
            color: #343a42 !important;
            fill: #343a42 !important;
          }
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within > div,
          [data-testid="stSidebar"] [data-testid="stSelectbox"] div[role="combobox"]:focus,
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="input"]:focus-within,
          [data-testid="stSidebar"] [data-testid="stTextInput"] div[data-baseweb="base-input"]:focus-within {
            border-color: #8f969f !important;
            box-shadow: 0 0 0 1px #8f969f !important;
          }

          /* BaseWeb mounts dropdown menus in a page-level portal. */
          div[data-baseweb="popover"] [role="listbox"],
          div[data-baseweb="popover"] [role="option"],
          div[data-baseweb="popover"] [role="option"] * {
            background-color: #ffffff !important;
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
          }
          div[data-baseweb="popover"] [role="option"]:hover,
          div[data-baseweb="popover"] [role="option"][aria-selected="true"] {
            background-color: #f0f2f5 !important;
          }

          /* r10: decisive light-control text fix.
             Do not scope this to stSelectbox: some Streamlit releases render
             the BaseWeb select without that wrapper.  These selectors target
             the actual BaseWeb control, including disabled/placeholder states. */
          [data-testid="stSidebar"] div[data-baseweb="select"],
          [data-testid="stSidebar"] div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] div[role="combobox"],
          [data-testid="stSidebar"] div[data-baseweb="input"],
          [data-testid="stSidebar"] div[data-baseweb="base-input"],
          [data-testid="stSidebar"] div[data-baseweb="input"] > div,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] > div {
            background: #ffffff !important;
            background-color: #ffffff !important;
            background-image: none !important;
            border-color: #d5d9df !important;
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            color-scheme: light !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="select"] *,
          [data-testid="stSidebar"] div[role="combobox"] *,
          [data-testid="stSidebar"] div[data-baseweb="select"] *::before,
          [data-testid="stSidebar"] div[data-baseweb="select"] *::after,
          [data-testid="stSidebar"] div[role="combobox"] *::before,
          [data-testid="stSidebar"] div[role="combobox"] *::after,
          [data-testid="stSidebar"] div[data-baseweb="select"] input,
          [data-testid="stSidebar"] div[role="combobox"] input,
          [data-testid="stSidebar"] div[data-baseweb="input"] input,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] input {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            text-shadow: none !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="select"][aria-disabled="true"],
          [data-testid="stSidebar"] div[role="combobox"][aria-disabled="true"],
          [data-testid="stSidebar"] div[data-baseweb="select"] [aria-disabled="true"],
          [data-testid="stSidebar"] div[role="combobox"] [aria-disabled="true"] {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="select"] svg,
          [data-testid="stSidebar"] div[role="combobox"] svg {
            color: #343a42 !important;
            fill: #343a42 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] div[data-baseweb="input"] input::placeholder,
          [data-testid="stSidebar"] div[data-baseweb="base-input"] input::placeholder {
            color: #7a828d !important;
            -webkit-text-fill-color: #7a828d !important;
            opacity: 1 !important;
          }

          /* The uploader button is intentionally light, so its text must be dark. */
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
            opacity: 1 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span {
            color: #eef1f4 !important;
            -webkit-text-fill-color: #eef1f4 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
            color: #8c939d !important;
            -webkit-text-fill-color: #8c939d !important;
          }

          [data-testid="stSidebar"] .tc-output-location {
            min-height: 2.3rem !important;
            margin: 0 !important;
            padding: 0.52rem 0.68rem !important;
            background: #16191d !important;
            border: 1px solid #2d3239 !important;
            border-radius: 8px !important;
          }
          [data-testid="stSidebar"] .tc-output-location-path {
            color: #aeb4bc !important;
            font-size: 0.63rem !important;
          }

          @media (max-width: 980px) {
            [data-testid="stSidebar"],
            [data-testid="stSidebar"] > div:first-child {
              width: 330px !important;
              min-width: 330px !important;
            }
            [data-testid="stAppViewBlockContainer"],
            [data-testid="stMainBlockContainer"],
            .block-container {
              padding: 0.12rem 0.42rem 0 !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r19: custom 25% -> 5% control-rail toggle. The state is applied to
             the root document by the zero-height controller in control_panel.py. */
          :root {
            --tc-sidebar-expanded-width: clamp(320px, 25vw, 410px);
            --tc-sidebar-compact-width: max(64px, 5vw);
            --tc-sidebar-motion: 220ms cubic-bezier(.2, .8, .2, 1);
          }

          [data-testid="stSidebarCollapsedControl"],
          [data-testid="stSidebarCollapseButton"] {
            display: none !important;
          }

          [data-testid="stSidebar"],
          [data-testid="stSidebar"] > div:first-child {
            transition:
              width var(--tc-sidebar-motion),
              min-width var(--tc-sidebar-motion),
              max-width var(--tc-sidebar-motion),
              flex-basis var(--tc-sidebar-motion),
              box-shadow var(--tc-sidebar-motion) !important;
          }

          [data-testid="stSidebar"] {
            position: relative !important;
            overflow: visible !important;
            width: var(--tc-sidebar-expanded-width) !important;
            min-width: var(--tc-sidebar-expanded-width) !important;
            max-width: var(--tc-sidebar-expanded-width) !important;
            flex: 0 0 var(--tc-sidebar-expanded-width) !important;
          }

          [data-testid="stSidebar"] > div:first-child {
            width: var(--tc-sidebar-expanded-width) !important;
            min-width: var(--tc-sidebar-expanded-width) !important;
            max-width: var(--tc-sidebar-expanded-width) !important;
          }

          #tc-sidebar-toggle {
            position: absolute;
            top: 0.78rem;
            right: -1.03rem;
            z-index: 100000;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 2.08rem;
            height: 2.08rem;
            margin: 0;
            padding: 0;
            border: 1px solid #383c43;
            border-radius: 999px;
            background: #1b1d21;
            color: #f5f6f7;
            box-shadow: 0 7px 18px rgba(0, 0, 0, 0.24);
            cursor: pointer;
            outline: none;
            transition:
              transform 160ms ease,
              background-color 160ms ease,
              border-color 160ms ease,
              box-shadow 160ms ease;
          }

          #tc-sidebar-toggle:hover {
            transform: translateY(-1px) scale(1.04);
            background: #c90000;
            border-color: #ef5656;
            box-shadow: 0 9px 22px rgba(201, 0, 0, 0.26);
          }

          #tc-sidebar-toggle:active {
            transform: translateY(0) scale(0.97);
          }

          #tc-sidebar-toggle:focus-visible {
            box-shadow:
              0 0 0 3px rgba(255, 255, 255, 0.72),
              0 0 0 5px rgba(201, 0, 0, 0.76),
              0 8px 20px rgba(0, 0, 0, 0.24);
          }

          #tc-sidebar-toggle .tc-sidebar-toggle-icon {
            display: block;
            width: 0.92rem;
            height: 1rem;
            overflow: visible;
            pointer-events: none;
          }

          #tc-sidebar-mini-brand {
            position: absolute;
            top: 3.72rem;
            left: 50%;
            z-index: 99999;
            display: none;
            align-items: center;
            justify-content: center;
            width: 2.42rem;
            height: 2.42rem;
            transform: translateX(-50%);
            border: 1px solid rgba(255, 255, 255, 0.13);
            border-radius: 0.72rem;
            background: linear-gradient(145deg, #df1018 0%, #bd0008 100%);
            color: #ffffff;
            font-size: 0.72rem;
            font-weight: 850;
            letter-spacing: 0.04em;
            box-shadow: 0 8px 22px rgba(0, 0, 0, 0.28);
            user-select: none;
          }

          #tc-sidebar-mini-brand span {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }

          #tc-sidebar-mini-actions {
            position: absolute;
            top: 7.05rem;
            left: 50%;
            z-index: 99998;
            display: none;
            width: calc(var(--tc-sidebar-compact-width) - 12px);
            transform: translateX(-50%);
            flex-direction: column;
            align-items: center;
            gap: 0.72rem;
            pointer-events: auto;
          }

          .tc-sidebar-mini-action {
            position: relative;
            display: inline-flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 0.18rem;
            width: min(48px, calc(var(--tc-sidebar-compact-width) - 14px));
            min-width: 0;
            height: 48px;
            margin: 0;
            padding: 6px 3px 5px;
            border: 1px solid rgba(255, 255, 255, 0.13);
            border-radius: 13px;
            background: #17191d;
            color: #f4f5f6;
            box-shadow: 0 7px 18px rgba(0, 0, 0, 0.24);
            cursor: pointer;
            outline: none;
            transition:
              transform 150ms ease,
              background-color 150ms ease,
              border-color 150ms ease,
              box-shadow 150ms ease,
              opacity 150ms ease;
          }

          .tc-sidebar-mini-action svg {
            display: block;
            width: 19px;
            height: 19px;
            color: currentColor;
            flex: 0 0 auto;
          }

          .tc-sidebar-mini-action-label {
            display: block;
            max-width: 100%;
            overflow: hidden;
            color: currentColor !important;
            -webkit-text-fill-color: currentColor !important;
            font-size: 0.48rem;
            line-height: 1;
            font-weight: 800;
            letter-spacing: 0.025em;
            text-overflow: ellipsis;
            white-space: nowrap;
          }

          .tc-sidebar-mini-action:hover:not(:disabled) {
            transform: translateY(-1px);
            background: #24272c;
            border-color: rgba(255, 255, 255, 0.28);
            box-shadow: 0 9px 21px rgba(0, 0, 0, 0.30);
          }

          .tc-sidebar-mini-action:active:not(:disabled) {
            transform: scale(0.96);
          }

          .tc-sidebar-mini-action.is-generate {
            margin-top: 0.24rem;
            background: linear-gradient(145deg, #d70b13 0%, #bd0008 100%);
            border-color: rgba(255, 255, 255, 0.22);
          }

          .tc-sidebar-mini-action.is-generate:hover:not(:disabled) {
            background: linear-gradient(145deg, #ed111b 0%, #c90000 100%);
          }

          .tc-sidebar-mini-action:disabled {
            cursor: default;
            opacity: 0.34;
            filter: grayscale(0.25);
            box-shadow: none;
          }

          html.tc-sidebar-is-collapsed #tc-sidebar-mini-actions,
          body.tc-sidebar-is-collapsed #tc-sidebar-mini-actions {
            display: flex;
          }

          html.tc-sidebar-is-collapsed [data-testid="stSidebar"],
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"] {
            width: var(--tc-sidebar-compact-width) !important;
            min-width: var(--tc-sidebar-compact-width) !important;
            max-width: var(--tc-sidebar-compact-width) !important;
            flex: 0 0 var(--tc-sidebar-compact-width) !important;
            box-shadow: 5px 0 18px rgba(10, 12, 16, 0.12) !important;
          }

          html.tc-sidebar-is-collapsed [data-testid="stSidebar"] > div:first-child,
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"] > div:first-child {
            width: var(--tc-sidebar-compact-width) !important;
            min-width: var(--tc-sidebar-compact-width) !important;
            max-width: var(--tc-sidebar-compact-width) !important;
            overflow: hidden !important;
          }

          html.tc-sidebar-is-collapsed [data-testid="stSidebar"] [data-testid="stSidebarContent"],
          html.tc-sidebar-is-collapsed [data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"] [data-testid="stSidebarContent"],
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
            width: var(--tc-sidebar-compact-width) !important;
            min-width: var(--tc-sidebar-compact-width) !important;
            max-width: var(--tc-sidebar-compact-width) !important;
            padding: 0 !important;
            overflow: hidden !important;
          }

          html.tc-sidebar-is-collapsed [data-testid="stSidebar"] .st-key-control_panel,
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"] .st-key-control_panel {
            display: none !important;
          }

          html.tc-sidebar-is-collapsed #tc-sidebar-mini-brand,
          body.tc-sidebar-is-collapsed #tc-sidebar-mini-brand {
            display: inline-flex;
          }

          html.tc-sidebar-is-collapsed #tc-sidebar-toggle,
          body.tc-sidebar-is-collapsed #tc-sidebar-toggle {
            background: #c90000;
            border-color: #ed4545;
            color: #ffffff;
          }

          html.tc-sidebar-is-collapsed #tc-sidebar-toggle:hover,
          body.tc-sidebar-is-collapsed #tc-sidebar-toggle:hover {
            background: #a90000;
            border-color: #ff6969;
          }

          @media (prefers-reduced-motion: reduce) {
            [data-testid="stSidebar"],
            [data-testid="stSidebar"] > div:first-child,
            #tc-sidebar-toggle {
              transition: none !important;
            }
          }

          @media (max-width: 820px) {
            :root { --tc-sidebar-compact-width: 64px; }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r26: dual collapsible rails. The primary control rail collapses to
             2.5% and the slide-input rail opens beside it only when needed. */
          :root {
            --tc-sidebar-compact-width: max(40px, 2.5vw);
            --tc-sidebar-current-width: var(--tc-sidebar-expanded-width);
            --tc-input-sidebar-expanded-width: clamp(270px, 19vw, 340px);
            --tc-input-sidebar-compact-width: max(40px, 2.5vw);
            --tc-input-sidebar-current-width: var(--tc-input-sidebar-expanded-width);
            --tc-input-sidebar-motion: 220ms cubic-bezier(.2, .8, .2, 1);
          }

          html.tc-sidebar-is-collapsed {
            --tc-sidebar-current-width: var(--tc-sidebar-compact-width);
          }
          html.tc-input-sidebar-is-collapsed {
            --tc-input-sidebar-current-width: var(--tc-input-sidebar-compact-width);
          }

          /* Make the native rail genuinely compact rather than retaining the
             earlier 5%/64px minimum. */
          html.tc-sidebar-is-collapsed [data-testid="stSidebar"],
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"],
          html.tc-sidebar-is-collapsed [data-testid="stSidebar"] > div:first-child,
          body.tc-sidebar-is-collapsed [data-testid="stSidebar"] > div:first-child {
            width: var(--tc-sidebar-compact-width) !important;
            min-width: var(--tc-sidebar-compact-width) !important;
            max-width: var(--tc-sidebar-compact-width) !important;
            flex-basis: var(--tc-sidebar-compact-width) !important;
          }
          #tc-sidebar-mini-brand {
            width: 1.78rem !important;
            height: 1.78rem !important;
            top: 3.55rem !important;
            border-radius: 0.58rem !important;
            font-size: 0.62rem !important;
          }
          #tc-sidebar-toggle {
            width: 1.8rem !important;
            height: 1.8rem !important;
            right: -0.95rem !important;
          }

          /* The dedicated slide-input rail is rendered outside Streamlit's
             native sidebar and positioned directly beside it. */
          .st-key-l1_input_sidebar {
            position: fixed !important;
            top: 0 !important;
            bottom: 0 !important;
            left: var(--tc-sidebar-current-width) !important;
            z-index: 99980 !important;
            display: none !important;
            width: var(--tc-input-sidebar-current-width) !important;
            min-width: var(--tc-input-sidebar-current-width) !important;
            max-width: var(--tc-input-sidebar-current-width) !important;
            height: 100vh !important;
            min-height: 100vh !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: visible !important;
            border: 0 !important;
            border-right: 1px solid rgba(255,255,255,0.22) !important;
            border-radius: 0 !important;
            background:
              linear-gradient(180deg, rgba(201, 0, 18, 0.91) 0%, rgba(150, 0, 20, 0.90) 48%, rgba(92, 0, 17, 0.90) 100%) !important;
            box-shadow: 10px 0 28px rgba(38, 0, 8, 0.18) !important;
            backdrop-filter: blur(15px) saturate(125%);
            -webkit-backdrop-filter: blur(15px) saturate(125%);
            transition:
              left var(--tc-sidebar-motion),
              width var(--tc-input-sidebar-motion),
              min-width var(--tc-input-sidebar-motion),
              max-width var(--tc-input-sidebar-motion),
              box-shadow var(--tc-input-sidebar-motion) !important;
          }
          html.tc-input-sidebar-is-open .st-key-l1_input_sidebar,
          body.tc-input-sidebar-is-open .st-key-l1_input_sidebar {
            display: block !important;
          }
          .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
            position: relative !important;
            z-index: 1 !important;
            height: 100vh !important;
            min-height: 100vh !important;
            max-height: 100vh !important;
            gap: 0 !important;
            margin: 0 !important;
            padding: 1.05rem 0.78rem 1rem !important;
            overflow-x: hidden !important;
            overflow-y: auto !important;
            scrollbar-width: none !important;
            box-sizing: border-box !important;
            transition: opacity 140ms ease, visibility 140ms ease !important;
          }
          .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"]::-webkit-scrollbar,
          .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"]::-webkit-scrollbar {
            display: none !important;
            width: 0 !important;
          }
          html.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          html.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"],
          body.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          body.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
          }

          .st-key-l1_input_sidebar .tc-input-sidebar-accent {
            position: absolute;
            inset: 0;
            pointer-events: none;
            background:
              radial-gradient(circle at 20% 5%, rgba(255,255,255,0.17), transparent 24%),
              linear-gradient(135deg, rgba(255,255,255,0.05), transparent 35%);
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-heading {
            display: flex !important;
            align-items: flex-start !important;
            justify-content: space-between !important;
            gap: 0.55rem !important;
            margin: 0 0 0.52rem !important;
            padding: 0 0 0.65rem !important;
            border-top: 0 !important;
            border-bottom: 1px solid rgba(255,255,255,0.22) !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-title {
            color: #ffffff !important;
            font-size: 0.82rem !important;
            font-weight: 850 !important;
            letter-spacing: -0.01em !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-copy {
            max-width: 15rem;
            margin-top: 0.18rem !important;
            color: rgba(255,255,255,0.72) !important;
            font-size: 0.56rem !important;
            line-height: 1.35 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-count {
            min-width: 2.6rem !important;
            padding: 0.2rem 0.4rem !important;
            border-color: rgba(255,255,255,0.34) !important;
            background: rgba(42,0,9,0.34) !important;
            color: #ffffff !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-count.is-ready {
            border-color: rgba(188,255,222,0.58) !important;
            background: rgba(8,73,46,0.58) !important;
            color: #d6ffed !important;
          }

          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
            gap: 0.24rem !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] {
            min-width: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] > div > [data-testid="stVerticalBlock"] {
            gap: 0.02rem !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button {
            min-height: 2.2rem !important;
            padding: 0.34rem 0.58rem !important;
            border-radius: 8px !important;
            font-size: 0.68rem !important;
            font-weight: 760 !important;
            line-height: 1.15 !important;
            text-align: left !important;
            justify-content: flex-start !important;
            box-shadow: none !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"] {
            background: rgba(255,255,255,0.96) !important;
            border-color: rgba(255,255,255,0.70) !important;
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"] * {
            color: #111418 !important;
            -webkit-text-fill-color: #111418 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"]:hover {
            background: #ffffff !important;
            border-color: rgba(70,0,10,0.42) !important;
            transform: translateY(-1px);
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] {
            background: rgba(20,22,26,0.94) !important;
            border-color: rgba(255,255,255,0.66) !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            box-shadow: 0 7px 18px rgba(42,0,8,0.20) !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] * {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button:disabled {
            opacity: 0.58 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-range {
            padding: 0.08rem 0.5rem 0.12rem !important;
            color: rgba(255,255,255,0.67) !important;
            font-size: 0.51rem !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-range.is-mapped {
            color: #c9ffe6 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status {
            margin: 0.3rem 0 0 !important;
            padding: 0.42rem 0.52rem !important;
            border-left-color: rgba(255,255,255,0.74) !important;
            background: rgba(40,0,8,0.33) !important;
            color: rgba(255,255,255,0.79) !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-ready {
            border-left-color: #a4f4cf !important;
            background: rgba(10,74,48,0.56) !important;
            color: #d8ffed !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-error {
            border-left-color: #ffd0d5 !important;
            background: rgba(68,0,11,0.48) !important;
            color: #ffe3e6 !important;
          }

          #tc-input-sidebar-toggle {
            position: absolute;
            top: 0.78rem;
            right: -0.96rem;
            z-index: 100005;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.92rem;
            height: 1.92rem;
            margin: 0;
            padding: 0;
            border: 1px solid rgba(255,255,255,0.50);
            border-radius: 999px;
            background: rgba(74,0,12,0.92);
            color: #ffffff;
            box-shadow: 0 8px 20px rgba(42,0,9,0.28);
            cursor: pointer;
            outline: none;
            transition: transform 160ms ease, background-color 160ms ease, box-shadow 160ms ease;
          }
          #tc-input-sidebar-toggle:hover {
            transform: translateY(-1px) scale(1.04);
            background: #15171b;
            box-shadow: 0 10px 24px rgba(23,0,5,0.34);
          }
          #tc-input-sidebar-toggle:active { transform: scale(0.97); }
          #tc-input-sidebar-toggle:focus-visible {
            box-shadow: 0 0 0 3px rgba(255,255,255,0.70), 0 0 0 5px rgba(55,0,10,0.66);
          }
          #tc-input-sidebar-toggle .tc-input-sidebar-toggle-icon {
            width: 0.9rem;
            height: 1rem;
            pointer-events: none;
          }

          #tc-input-sidebar-mini {
            position: absolute;
            top: 3.55rem;
            left: 50%;
            z-index: 100004;
            display: none;
            align-items: center;
            justify-content: center;
            flex-direction: column;
            gap: 0.18rem;
            width: 1.78rem;
            transform: translateX(-50%);
            color: #ffffff;
            user-select: none;
          }
          #tc-input-sidebar-mini span {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }
          .tc-input-sidebar-mini-title {
            font-size: 0.76rem;
            font-weight: 900;
            line-height: 1;
          }
          .tc-input-sidebar-mini-subtitle {
            writing-mode: vertical-rl;
            transform: rotate(180deg);
            font-size: 0.42rem;
            font-weight: 800;
            letter-spacing: 0.12em;
            opacity: 0.76;
          }
          html.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini,
          body.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini {
            display: flex;
          }

          /* The workbook canvas yields exactly the width occupied by the second
             rail, and regains it immediately when that rail is collapsed. */
          .st-key-canvas {
            box-sizing: border-box !important;
            transition: margin-left var(--tc-input-sidebar-motion), width var(--tc-input-sidebar-motion) !important;
          }
          html.tc-input-sidebar-is-open .st-key-canvas,
          body.tc-input-sidebar-is-open .st-key-canvas {
            width: calc(100% - var(--tc-input-sidebar-current-width)) !important;
            margin-left: var(--tc-input-sidebar-current-width) !important;
          }

          @media (prefers-reduced-motion: reduce) {
            .st-key-l1_input_sidebar,
            .st-key-canvas,
            #tc-input-sidebar-toggle {
              transition: none !important;
            }
          }
          @media (max-width: 900px) {
            :root {
              --tc-sidebar-compact-width: 40px;
              --tc-input-sidebar-compact-width: 40px;
              --tc-input-sidebar-expanded-width: min(310px, 62vw);
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r27: final pass for the primary black control rail only.
             Keep the L1 input rail unchanged while aligning the Slide Maker
             masthead with the canvas title and using the available height. */

          /* Streamlit reserves a header strip inside some sidebar builds even
             after the standard collapse control is hidden. Remove that strip
             and anchor our control panel directly to the top of the rail. */
          [data-testid="stSidebar"] [data-testid="stSidebarHeader"],
          [data-testid="stSidebar"] [data-testid="stSidebarNav"],
          [data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"],
          [data-testid="stSidebar"] [data-testid="stSidebarCollapsedControl"] {
            display: none !important;
            width: 0 !important;
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          [data-testid="stSidebar"] > div:first-child,
          [data-testid="stSidebar"] [data-testid="stSidebarContent"],
          [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
            padding-top: 0 !important;
            margin-top: 0 !important;
          }
          [data-testid="stSidebar"] [data-testid="stSidebarContent"],
          [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
            position: relative !important;
            height: 100vh !important;
            min-height: 100vh !important;
            overflow: hidden !important;
          }
          [data-testid="stSidebar"] .st-key-control_panel {
            position: absolute !important;
            inset: 0 !important;
            box-sizing: border-box !important;
            width: 100% !important;
            height: 100vh !important;
            min-height: 100vh !important;
            max-height: 100vh !important;
            margin: 0 !important;
            padding: 0.62rem 1.38rem 1.2rem !important;
            overflow: hidden !important;
          }
          [data-testid="stSidebar"] .st-key-control_panel > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-control_panel > div > [data-testid="stVerticalBlock"] {
            min-height: 100% !important;
            gap: 0 !important;
          }

          /* Masthead: visually align Slide Maker with Workbook preview. */
          [data-testid="stSidebar"] .tc-panel-brand {
            display: flex !important;
            align-items: center !important;
            gap: 0.88rem !important;
            margin: 0 !important;
            padding: 0.02rem 0 0.92rem !important;
            border-bottom: 1px solid #2a2e34 !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-mark {
            width: 2.65rem !important;
            height: 2.65rem !important;
            border-radius: 0.68rem !important;
            background: linear-gradient(145deg, #e0121b 0%, #c40008 100%) !important;
            color: #ffffff !important;
            font-size: 0.72rem !important;
            font-weight: 900 !important;
            letter-spacing: 0.055em !important;
            box-shadow: 0 7px 18px rgba(0, 0, 0, 0.22) !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-title {
            color: #ffffff !important;
            font-size: 1.42rem !important;
            line-height: 1.04 !important;
            font-weight: 830 !important;
            letter-spacing: -0.028em !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-subtitle {
            margin-top: 0.3rem !important;
            color: #9299a3 !important;
            font-size: 0.66rem !important;
            line-height: 1.15 !important;
            font-weight: 800 !important;
            letter-spacing: 0.155em !important;
          }

          /* Use the large amount of free vertical space for clearer grouping. */
          [data-testid="stSidebar"] .st-key-panel_source {
            margin-top: 1.48rem !important;
            padding: 0 !important;
          }
          [data-testid="stSidebar"] .st-key-panel_build,
          [data-testid="stSidebar"] .st-key-panel_output {
            margin-top: 1.58rem !important;
            padding-top: 1.42rem !important;
            border-top: 1px solid #25292f !important;
          }
          [data-testid="stSidebar"] .st-key-panel_source > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_source > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > div > [data-testid="stVerticalBlock"] {
            gap: 0.58rem !important;
          }

          /* The stable custom labels are the only labels shown. This removes
             the duplicate SOURCE WORKBOOK line and stray help-dot alignment. */
          [data-testid="stSidebar"] .st-key-panel_source [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-panel_build [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-panel_output [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-panel_source [data-testid="stFileUploader"] > label,
          [data-testid="stSidebar"] .st-key-panel_build [data-testid="stSelectbox"] > label,
          [data-testid="stSidebar"] .st-key-panel_build [data-testid="stTextInput"] > label,
          [data-testid="stSidebar"] .st-key-panel_output [data-testid="stTextInput"] > label,
          [data-testid="stSidebar"] .st-key-panel_source [data-testid="stTooltipIcon"],
          [data-testid="stSidebar"] .st-key-panel_build [data-testid="stTooltipIcon"],
          [data-testid="stSidebar"] .st-key-panel_output [data-testid="stTooltipIcon"] {
            display: none !important;
            width: 0 !important;
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label {
            display: flex !important;
            align-items: baseline !important;
            justify-content: space-between !important;
            gap: 0.75rem !important;
            min-height: 1.12rem !important;
            margin: 0 0 0.38rem !important;
            padding: 0 0.02rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-spaced {
            margin-top: 0.86rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-text {
            color: #f7f8fa !important;
            font-size: 0.86rem !important;
            line-height: 1.2 !important;
            font-weight: 820 !important;
            letter-spacing: -0.006em !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-hint {
            color: #8f97a1 !important;
            font-size: 0.56rem !important;
            line-height: 1.2 !important;
            font-weight: 800 !important;
            letter-spacing: 0.1em !important;
            text-transform: uppercase !important;
          }

          /* Consistent component sizing and spacing across every primary-rail field. */
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            min-height: 5.45rem !important;
            padding: 0.84rem 0.88rem !important;
            border-radius: 0.72rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] span {
            font-size: 0.79rem !important;
            line-height: 1.3 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] small {
            margin-top: 0.18rem !important;
            font-size: 0.62rem !important;
            line-height: 1.25 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
            min-height: 2.32rem !important;
            padding: 0.36rem 0.84rem !important;
            border-radius: 0.52rem !important;
            font-size: 0.75rem !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] > div,
          [data-testid="stSidebar"] [data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-baseweb="base-input"] {
            min-height: 3.02rem !important;
            border-radius: 0.68rem !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] span,
          [data-testid="stSidebar"] [data-baseweb="select"] input,
          [data-testid="stSidebar"] [data-baseweb="input"] input,
          [data-testid="stSidebar"] [data-baseweb="base-input"] input {
            font-size: 0.82rem !important;
            line-height: 1.25 !important;
          }
          [data-testid="stSidebar"] .tc-upload-status {
            margin-top: 0.28rem !important;
            gap: 0.5rem !important;
            color: #9aa1ab !important;
            font-size: 0.67rem !important;
            line-height: 1.35 !important;
          }
          [data-testid="stSidebar"] .tc-upload-status-dot {
            width: 0.42rem !important;
            height: 0.42rem !important;
          }
          [data-testid="stSidebar"] .tc-output-location {
            min-height: 2.7rem !important;
            margin-top: 0.05rem !important;
            padding: 0.62rem 0.76rem !important;
            border-radius: 0.64rem !important;
          }
          [data-testid="stSidebar"] .tc-output-location-path {
            color: #b8bec6 !important;
            font-size: 0.68rem !important;
            line-height: 1.35 !important;
          }

          [data-testid="stSidebar"] [data-testid="stExpander"] {
            margin-top: 1.52rem !important;
            border-radius: 0.68rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            min-height: 2.72rem !important;
            padding: 0.58rem 0.76rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary p {
            color: #d0d4d9 !important;
            font-size: 0.75rem !important;
            font-weight: 720 !important;
          }
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label {
            min-height: 2rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
            font-size: 0.72rem !important;
          }

          [data-testid="stSidebar"] .st-key-panel_action {
            margin: 1.68rem 0 0 !important;
          }
          [data-testid="stSidebar"] .st-key-panel_action [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_action > div > [data-testid="stVerticalBlock"] {
            gap: 0.52rem !important;
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            min-height: 3.18rem !important;
            border-radius: 0.7rem !important;
            font-size: 0.91rem !important;
            font-weight: 780 !important;
          }
          [data-testid="stSidebar"] .tc-panel-footnote {
            margin-top: 0.5rem !important;
            font-size: 0.66rem !important;
            line-height: 1.4 !important;
          }

          /* Preserve a no-scroll rail on short displays without letting labels
             become tiny again. */
          @media (max-height: 770px) {
            [data-testid="stSidebar"] .st-key-control_panel {
              padding-top: 0.48rem !important;
              padding-bottom: 0.72rem !important;
            }
            [data-testid="stSidebar"] .tc-panel-brand {
              padding-bottom: 0.68rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_source {
              margin-top: 1rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_build,
            [data-testid="stSidebar"] .st-key-panel_output {
              margin-top: 1.08rem !important;
              padding-top: 0.96rem !important;
            }
            [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
              min-height: 4.72rem !important;
            }
            [data-testid="stSidebar"] [data-baseweb="select"] > div,
            [data-testid="stSidebar"] [data-baseweb="input"] > div,
            [data-testid="stSidebar"] [data-baseweb="base-input"] {
              min-height: 2.7rem !important;
            }
            [data-testid="stSidebar"] [data-testid="stExpander"] {
              margin-top: 1rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_action {
              margin-top: 1.05rem !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r28: professional primary rail + 50%-transparent Bain-red input rail.
             This is the final visual override and intentionally comes after all
             earlier compatibility styles. */
          :root {
            --tc-bain-red: #cc0000;
            --tc-bain-red-rgb: 204, 0, 0;
            --tc-primary-rail-bg: #101214;
            --tc-primary-rail-line: #292d33;
          }

          /* -----------------------------------------------------------------
             Primary black control rail
             ----------------------------------------------------------------- */
          [data-testid="stSidebar"] .st-key-control_panel {
            padding: 0.96rem 1.36rem 1.08rem !important;
          }
          [data-testid="stSidebar"] .st-key-control_panel > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-control_panel > div > [data-testid="stVerticalBlock"] {
            display: flex !important;
            flex-direction: column !important;
            min-height: 100% !important;
            gap: 0 !important;
          }

          /* Align the Slide Maker masthead with the Workbook preview heading. */
          [data-testid="stSidebar"] .tc-panel-brand {
            min-height: 3rem !important;
            margin: 0 !important;
            padding: 0 0 1rem !important;
            border-bottom: 1px solid var(--tc-primary-rail-line) !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-mark {
            width: 2.72rem !important;
            height: 2.72rem !important;
            border-radius: 0.7rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-title {
            font-size: 1.46rem !important;
            line-height: 1.02 !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-subtitle {
            margin-top: 0.3rem !important;
            font-size: 0.67rem !important;
          }

          /* Use the available vertical space without crowding adjacent fields. */
          [data-testid="stSidebar"] .st-key-panel_source {
            margin-top: 1.52rem !important;
            padding: 0 !important;
          }
          [data-testid="stSidebar"] .st-key-panel_build,
          [data-testid="stSidebar"] .st-key-panel_output {
            margin-top: 1.78rem !important;
            padding-top: 1.52rem !important;
            border-top: 1px solid var(--tc-primary-rail-line) !important;
          }
          [data-testid="stSidebar"] .st-key-panel_source > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_source > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
          }

          /* Padding, rather than a collapsing bottom margin, guarantees that a
             widget can never sit on top of its visible field label. */
          [data-testid="stSidebar"] .tc-panel-field-label {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            justify-content: space-between !important;
            min-height: 1.42rem !important;
            margin: 0 !important;
            padding: 0 0.02rem 0.62rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-spaced {
            margin-top: 1.16rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-text {
            color: #f8f9fa !important;
            font-size: 0.94rem !important;
            line-height: 1.24 !important;
            font-weight: 820 !important;
            letter-spacing: -0.012em !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label-hint {
            color: #929aa5 !important;
            font-size: 0.59rem !important;
            line-height: 1.2 !important;
            font-weight: 800 !important;
            letter-spacing: 0.095em !important;
          }

          /* Give every actual widget its own clear row below the label. */
          [data-testid="stSidebar"] .st-key-panel_source [data-testid="stFileUploader"],
          [data-testid="stSidebar"] .st-key-panel_build [data-testid="stSelectbox"],
          [data-testid="stSidebar"] .st-key-panel_build [data-testid="stTextInput"],
          [data-testid="stSidebar"] .st-key-panel_output [data-testid="stTextInput"] {
            position: relative !important;
            z-index: 1 !important;
            clear: both !important;
            margin: 0 !important;
          }
          [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            min-height: 5.55rem !important;
            padding: 0.88rem 0.9rem !important;
            border-radius: 0.76rem !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] > div,
          [data-testid="stSidebar"] [data-baseweb="input"] > div,
          [data-testid="stSidebar"] [data-baseweb="base-input"] {
            min-height: 3.12rem !important;
            border-radius: 0.72rem !important;
          }
          [data-testid="stSidebar"] [data-baseweb="select"] span,
          [data-testid="stSidebar"] [data-baseweb="select"] input,
          [data-testid="stSidebar"] [data-baseweb="input"] input,
          [data-testid="stSidebar"] [data-baseweb="base-input"] input {
            font-size: 0.84rem !important;
          }
          [data-testid="stSidebar"] .tc-upload-status {
            margin-top: 0.42rem !important;
            font-size: 0.69rem !important;
          }
          [data-testid="stSidebar"] .tc-output-location {
            min-height: 2.84rem !important;
            padding: 0.66rem 0.78rem !important;
            border-radius: 0.68rem !important;
          }
          [data-testid="stSidebar"] .tc-output-location-path {
            font-size: 0.7rem !important;
          }

          [data-testid="stSidebar"] [data-testid="stExpander"] {
            margin-top: 1.65rem !important;
            border-radius: 0.72rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            min-height: 2.86rem !important;
            padding: 0.62rem 0.78rem !important;
          }
          [data-testid="stSidebar"] [data-testid="stExpander"] summary p {
            font-size: 0.78rem !important;
          }

          /* A fixed-bottom CTA uses the otherwise empty lower rail space and is
             easier to find without compressing the configuration sections. */
          [data-testid="stSidebar"] .st-key-panel_action {
            margin-top: auto !important;
            padding-top: 1.6rem !important;
          }
          [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            min-height: 3.3rem !important;
            border-radius: 0.74rem !important;
            font-size: 0.92rem !important;
          }

          /* -----------------------------------------------------------------
             Dedicated L1 input rail: Bain red at exactly 50% opacity
             ----------------------------------------------------------------- */
          :root {
            --tc-input-sidebar-expanded-width: clamp(292px, 20vw, 360px);
          }
          .st-key-l1_input_sidebar {
            background: rgba(var(--tc-bain-red-rgb), 0.50) !important;
            background-color: rgba(var(--tc-bain-red-rgb), 0.50) !important;
            background-image: none !important;
            border-right: 1px solid rgba(93, 0, 0, 0.28) !important;
            box-shadow: 10px 0 30px rgba(48, 0, 0, 0.16) !important;
            backdrop-filter: blur(18px) saturate(132%) !important;
            -webkit-backdrop-filter: blur(18px) saturate(132%) !important;
          }
          .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
            padding: 0.96rem 0.88rem 1rem !important;
          }
          .st-key-l1_input_sidebar .tc-input-sidebar-accent {
            display: none !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-heading {
            margin: 0 0 0.72rem !important;
            padding: 0.72rem 0.76rem !important;
            border: 1px solid rgba(255, 255, 255, 0.38) !important;
            border-radius: 0.78rem !important;
            background: rgba(255, 255, 255, 0.22) !important;
            box-shadow: 0 8px 20px rgba(70, 0, 0, 0.08) !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-title {
            color: #31090c !important;
            font-size: 1rem !important;
            line-height: 1.15 !important;
            font-weight: 880 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-copy {
            max-width: none !important;
            margin-top: 0.28rem !important;
            color: rgba(49, 9, 12, 0.78) !important;
            font-size: 0.64rem !important;
            line-height: 1.42 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-count {
            min-width: 2.8rem !important;
            padding: 0.24rem 0.46rem !important;
            border-color: rgba(49, 9, 12, 0.32) !important;
            background: rgba(49, 9, 12, 0.78) !important;
            color: #ffffff !important;
            font-size: 0.6rem !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-count.is-ready {
            border-color: rgba(13, 90, 59, 0.34) !important;
            background: rgba(13, 90, 59, 0.82) !important;
            color: #ffffff !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
            gap: 0.42rem !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button {
            min-height: 2.58rem !important;
            padding: 0.48rem 0.7rem !important;
            border-radius: 0.7rem !important;
            font-size: 0.74rem !important;
            font-weight: 780 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"] {
            background: rgba(255, 255, 255, 0.88) !important;
            border-color: rgba(255, 255, 255, 0.72) !important;
            color: #17191c !important;
            -webkit-text-fill-color: #17191c !important;
            box-shadow: 0 5px 14px rgba(68, 0, 0, 0.06) !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"] * {
            color: #17191c !important;
            -webkit-text-fill-color: #17191c !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] {
            background: rgba(49, 9, 12, 0.90) !important;
            border-color: rgba(49, 9, 12, 0.58) !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            box-shadow: 0 7px 18px rgba(66, 0, 4, 0.16) !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-range {
            padding: 0.02rem 0.62rem 0.16rem !important;
            color: rgba(49, 9, 12, 0.76) !important;
            font-size: 0.55rem !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-range.is-mapped {
            color: #154c39 !important;
            font-weight: 760 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status {
            margin-top: 0.66rem !important;
            padding: 0.54rem 0.62rem !important;
            border-left-color: rgba(49, 9, 12, 0.72) !important;
            border-radius: 0 0.55rem 0.55rem 0 !important;
            background: rgba(255, 255, 255, 0.28) !important;
            color: #31090c !important;
            font-size: 0.61rem !important;
            line-height: 1.4 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-ready {
            border-left-color: #155c42 !important;
            background: rgba(225, 255, 242, 0.62) !important;
            color: #104633 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-error {
            border-left-color: #7b0000 !important;
            background: rgba(255, 238, 238, 0.62) !important;
            color: #650000 !important;
          }

          #tc-input-sidebar-toggle {
            background: rgba(var(--tc-bain-red-rgb), 0.78) !important;
            border-color: rgba(255, 255, 255, 0.64) !important;
            box-shadow: 0 8px 22px rgba(76, 0, 0, 0.22) !important;
          }
          #tc-input-sidebar-toggle:hover {
            background: rgba(49, 9, 12, 0.92) !important;
          }
          #tc-input-sidebar-mini {
            color: #31090c !important;
          }
          #tc-input-sidebar-mini span {
            color: #31090c !important;
            -webkit-text-fill-color: #31090c !important;
          }

          /* Robust collapsed state: hide every Streamlit content block, not
             only one release-specific direct wrapper. The toggle and compact
             L1 marker are injected outside these blocks and remain visible. */
          html.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar [data-testid="stVerticalBlock"],
          body.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar [data-testid="stVerticalBlock"] {
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
          }
          html.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini,
          body.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini {
            display: flex !important;
          }

          /* Short screens compact gracefully instead of allowing overlap. */
          @media (max-height: 800px) {
            [data-testid="stSidebar"] .st-key-control_panel {
              padding-top: 0.72rem !important;
              padding-bottom: 0.72rem !important;
            }
            [data-testid="stSidebar"] .tc-panel-brand {
              padding-bottom: 0.72rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_source {
              margin-top: 1rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_build,
            [data-testid="stSidebar"] .st-key-panel_output {
              margin-top: 1.12rem !important;
              padding-top: 1rem !important;
            }
            [data-testid="stSidebar"] .tc-panel-field-label-spaced {
              margin-top: 0.82rem !important;
            }
            [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
              min-height: 4.6rem !important;
            }
            [data-testid="stSidebar"] [data-testid="stExpander"] {
              margin-top: 1rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_action {
              margin-top: 1.1rem !important;
              padding-top: 0 !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )



    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r29: reliable native labels in the primary rail and a simplified
             card-based L1 input rail. This final pass deliberately overrides
             every earlier compatibility rule. */

          /* ---------------------------------------------------------------
             Primary rail: use Streamlit's native label rows so controls can
             never overlap their headings on any supported Streamlit build.
             --------------------------------------------------------------- */
          [data-testid="stSidebar"] .st-key-control_panel {
            padding: 0.82rem 1.34rem 1.08rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand {
            padding-bottom: 0.88rem !important;
          }
          [data-testid="stSidebar"] .st-key-panel_source {
            margin-top: 1.18rem !important;
          }
          [data-testid="stSidebar"] .st-key-panel_build,
          [data-testid="stSidebar"] .st-key-panel_output {
            margin-top: 1.48rem !important;
            padding-top: 1.22rem !important;
          }
          [data-testid="stSidebar"] .st-key-panel_source > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_source > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_build > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-panel_output > div > [data-testid="stVerticalBlock"] {
            gap: 1.02rem !important;
          }

          [data-testid="stSidebar"] .st-key-field_source_workbook > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_input_section > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_only_l0 > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_output_name > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_output_folder > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_source_workbook > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_input_section > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_only_l0 > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_output_name > div > [data-testid="stVerticalBlock"],
          [data-testid="stSidebar"] .st-key-field_output_folder > div > [data-testid="stVerticalBlock"] {
            gap: 0.42rem !important;
          }

          /* Undo the older rule that hid native labels. The key-specific
             selectors make this more specific than every compatibility rule. */
          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-field_input_section [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-field_output_name [data-testid="stWidgetLabel"],
          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stFileUploader"] > label,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stSelectbox"] > label,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] > label,
          [data-testid="stSidebar"] .st-key-field_input_section [data-testid="stSelectbox"] > label,
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-testid="stTextInput"] > label,
          [data-testid="stSidebar"] .st-key-field_output_name [data-testid="stTextInput"] > label {
            position: static !important;
            display: flex !important;
            visibility: visible !important;
            opacity: 1 !important;
            align-items: center !important;
            width: auto !important;
            height: auto !important;
            min-height: 1.38rem !important;
            margin: 0 0 0.42rem !important;
            padding: 0 !important;
            overflow: visible !important;
            clip: auto !important;
            clip-path: none !important;
            transform: none !important;
          }
          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stWidgetLabel"] p,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stWidgetLabel"] p,
          [data-testid="stSidebar"] .st-key-field_input_section [data-testid="stWidgetLabel"] p,
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-testid="stWidgetLabel"] p,
          [data-testid="stSidebar"] .st-key-field_output_name [data-testid="stWidgetLabel"] p,
          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stFileUploader"] > label p,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stSelectbox"] > label p,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] > label p,
          [data-testid="stSidebar"] .st-key-field_input_section [data-testid="stSelectbox"] > label p,
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-testid="stTextInput"] > label p,
          [data-testid="stSidebar"] .st-key-field_output_name [data-testid="stTextInput"] > label p {
            margin: 0 !important;
            color: #f7f8fa !important;
            -webkit-text-fill-color: #f7f8fa !important;
            font-size: 0.9rem !important;
            line-height: 1.25 !important;
            font-weight: 820 !important;
            letter-spacing: -0.008em !important;
          }
          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stTooltipIcon"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stTooltipIcon"],
          [data-testid="stSidebar"] .st-key-field_input_section [data-testid="stTooltipIcon"],
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-testid="stTooltipIcon"],
          [data-testid="stSidebar"] .st-key-field_output_name [data-testid="stTooltipIcon"] {
            display: none !important;
          }
          [data-testid="stSidebar"] .tc-panel-field-label {
            display: none !important;
          }
          [data-testid="stSidebar"] .tc-panel-static-label {
            display: block !important;
            min-height: 1.38rem !important;
            margin: 0 0 0.42rem !important;
            padding: 0 !important;
            color: #f7f8fa !important;
            font-size: 0.9rem !important;
            line-height: 1.25 !important;
            font-weight: 820 !important;
            letter-spacing: -0.008em !important;
          }
          [data-testid="stSidebar"] .st-key-field_output_folder [data-testid="stElementContainer"]:has(.tc-panel-static-label) {
            display: block !important;
            visibility: visible !important;
            height: auto !important;
            min-height: 1.38rem !important;
            margin: 0 !important;
            padding: 0 !important;
          }

          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stFileUploader"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stSelectbox"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"],
          [data-testid="stSidebar"] .st-key-field_input_section [data-testid="stSelectbox"],
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-testid="stTextInput"],
          [data-testid="stSidebar"] .st-key-field_output_name [data-testid="stTextInput"] {
            position: static !important;
            display: block !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: visible !important;
          }
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-baseweb="select"] > div,
          [data-testid="stSidebar"] .st-key-field_input_section [data-baseweb="select"] > div,
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-baseweb="input"] > div,
          [data-testid="stSidebar"] .st-key-field_only_l0 [data-baseweb="base-input"],
          [data-testid="stSidebar"] .st-key-field_output_name [data-baseweb="input"] > div,
          [data-testid="stSidebar"] .st-key-field_output_name [data-baseweb="base-input"] {
            min-height: 3rem !important;
          }
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-baseweb="tag"] {
            background: #cc0000 !important;
            background-color: #cc0000 !important;
            border: 1px solid #cc0000 !important;
            border-radius: 0.42rem !important;
            margin: 0.12rem 0.18rem 0.12rem 0 !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-baseweb="tag"] * {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-weight: 800 !important;
          }
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-baseweb="tag"] svg,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-baseweb="tag"] svg * {
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
          }

          /* Deck selector surface: blend with the dark sidebar; only selected
             slide chips carry the Bain-red fill. */
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] div[data-baseweb="select"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] div[data-baseweb="select"] > div,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] div[role="combobox"] {
            background: #111214 !important;
            background-color: #111214 !important;
            background-image: none !important;
            border-color: #34383f !important;
            box-shadow: none !important;
          }

          /* Multiselect right-side controls: remove dark square backgrounds
             and keep the controls integrated with the dark selector surface. */
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] button,
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] [role="button"] {
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            border: 0 !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] .st-key-field_source_workbook [data-testid="stFileUploaderDropzone"] {
            min-height: 5.25rem !important;
          }
          [data-testid="stSidebar"] .st-key-field_output_folder .tc-output-location {
            margin: 0 !important;
          }
          [data-testid="stSidebar"] .st-key-panel_action {
            margin-top: auto !important;
            padding-top: 1.4rem !important;
          }

          /* ---------------------------------------------------------------
             L1 input rail: one clean header and one large two-line card for
             every input. The Streamlit button is the card's click target; the
             assigned sheet/range sits directly beneath the label.
             --------------------------------------------------------------- */
          .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
            padding: 0.94rem 0.84rem 1rem !important;
            gap: 0 !important;
          }
          .st-key-l1_input_sidebar .tc-input-sidebar-header {
            box-sizing: border-box !important;
            display: block !important;
            margin: 0 0 0.9rem !important;
            padding: 0.18rem 0.5rem 0.78rem !important;
            border-bottom: 1px solid rgba(73, 0, 5, 0.26) !important;
            color: #31090c !important;
            -webkit-text-fill-color: #31090c !important;
            font-size: 1.18rem !important;
            line-height: 1.1 !important;
            font-weight: 900 !important;
            letter-spacing: -0.025em !important;
          }
          .st-key-l1_input_sidebar [data-testid="stElementContainer"]:has(.tc-input-sidebar-header) {
            display: block !important;
            visibility: visible !important;
            height: auto !important;
            min-height: 2.15rem !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-heading,
          .st-key-l1_input_sidebar .tc-sidebar-l1-copy,
          .st-key-l1_input_sidebar .tc-sidebar-l1-count {
            display: none !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
            gap: 0.64rem !important;
          }

          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table {
            box-sizing: border-box !important;
            overflow: hidden !important;
            margin: 0 !important;
            padding: 0.24rem 0.26rem 0.42rem !important;
            border: 1px solid rgba(255,255,255,0.72) !important;
            border-radius: 0.82rem !important;
            background: rgba(255,255,255,0.90) !important;
            box-shadow: 0 7px 18px rgba(73,0,5,0.08) !important;
            transition: transform 150ms ease, background-color 150ms ease, box-shadow 150ms ease !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy:hover,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers:hover,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values:hover,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values:hover,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table:hover,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 10px 22px rgba(73,0,5,0.12) !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy:has(button[kind="primary"]),
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers:has(button[kind="primary"]),
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values:has(button[kind="primary"]),
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values:has(button[kind="primary"]),
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table:has(button[kind="primary"]),
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table:has(button[kind="primary"]) {
            border-color: rgba(66,0,5,0.56) !important;
            background: rgba(66,0,5,0.90) !important;
            box-shadow: 0 9px 22px rgba(60,0,4,0.20) !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy > div > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers > div > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values > div > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values > div > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table > div > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton {
            margin: 0 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button {
            display: flex !important;
            justify-content: flex-start !important;
            align-items: center !important;
            width: 100% !important;
            min-height: 2.55rem !important;
            margin: 0 !important;
            padding: 0.54rem 0.62rem 0.24rem !important;
            border: 0 !important;
            border-radius: 0.58rem !important;
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            box-shadow: none !important;
            text-align: left !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"],
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] {
            border-color: transparent !important;
            background: transparent !important;
            background-color: transparent !important;
            background-image: none !important;
            box-shadow: none !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button p {
            width: 100% !important;
            margin: 0 !important;
            color: #17191c !important;
            -webkit-text-fill-color: #17191c !important;
            font-size: 0.82rem !important;
            line-height: 1.24 !important;
            font-weight: 820 !important;
            text-align: left !important;
            white-space: normal !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] p {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-range {
            display: block !important;
            min-height: 1.25rem !important;
            margin: 0 !important;
            padding: 0.02rem 0.62rem 0.48rem !important;
            color: #6b5558 !important;
            -webkit-text-fill-color: #6b5558 !important;
            font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace !important;
            font-size: 0.64rem !important;
            line-height: 1.38 !important;
            font-weight: 650 !important;
            overflow-wrap: anywhere !important;
            white-space: normal !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-range.is-mapped {
            color: #275f4d !important;
            -webkit-text-fill-color: #275f4d !important;
            font-weight: 760 !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy:has(button[kind="primary"]) .tc-sidebar-l1-range,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers:has(button[kind="primary"]) .tc-sidebar-l1-range,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values:has(button[kind="primary"]) .tc-sidebar-l1-range,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values:has(button[kind="primary"]) .tc-sidebar-l1-range,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table:has(button[kind="primary"]) .tc-sidebar-l1-range,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table:has(button[kind="primary"]) .tc-sidebar-l1-range {
            color: rgba(255,255,255,0.80) !important;
            -webkit-text-fill-color: rgba(255,255,255,0.80) !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status {
            margin-top: 0.82rem !important;
            padding: 0.62rem 0.7rem !important;
            border-radius: 0.62rem !important;
            font-size: 0.64rem !important;
          }

          @media (max-height: 800px) {
            [data-testid="stSidebar"] .st-key-panel_source {
              margin-top: 0.9rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_build,
            [data-testid="stSidebar"] .st-key-panel_output {
              margin-top: 1rem !important;
              padding-top: 0.9rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_source > [data-testid="stVerticalBlock"],
            [data-testid="stSidebar"] .st-key-panel_build > [data-testid="stVerticalBlock"],
            [data-testid="stSidebar"] .st-key-panel_output > [data-testid="stVerticalBlock"],
            [data-testid="stSidebar"] .st-key-panel_source > div > [data-testid="stVerticalBlock"],
            [data-testid="stSidebar"] .st-key-panel_build > div > [data-testid="stVerticalBlock"],
            [data-testid="stSidebar"] .st-key-panel_output > div > [data-testid="stVerticalBlock"] {
              gap: 0.72rem !important;
            }
            [data-testid="stSidebar"] .st-key-panel_action {
              padding-top: 0.8rem !important;
            }
            .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
            .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
              gap: 0.46rem !important;
            }
            .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button {
              min-height: 2.32rem !important;
              padding-top: 0.42rem !important;
            }
            .st-key-l1_input_sidebar .tc-sidebar-l1-range {
              padding-bottom: 0.36rem !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r31: clean, opaque Bain-red companion rail.  Each L1 mapping is
             one self-contained two-line button, which removes the fragile
             separate range footer that previously clipped between cards. */
          :root {
            --tc-bain-red: #cc0000;
            --tc-input-sidebar-expanded-width: clamp(300px, 19vw, 354px);
          }

          .st-key-l1_input_sidebar {
            background: var(--tc-bain-red) !important;
            background-color: var(--tc-bain-red) !important;
            background-image: none !important;
            border-right: 1px solid #980000 !important;
            box-shadow: 10px 0 26px rgba(26, 0, 0, 0.18) !important;
            backdrop-filter: none !important;
            -webkit-backdrop-filter: none !important;
          }
          .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
            box-sizing: border-box !important;
            min-height: 100vh !important;
            padding: 1.12rem 0.9rem 1.2rem !important;
            gap: 0 !important;
            background: transparent !important;
            overflow-x: visible !important;
          }
          .st-key-l1_input_sidebar .tc-input-sidebar-accent {
            display: none !important;
          }

          /* Leave clear breathing room around the heading and away from both
             circular rail toggles. */
          .st-key-l1_input_sidebar .tc-input-sidebar-header {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            min-height: 2.7rem !important;
            margin: 0 0 0.9rem !important;
            padding: 0.08rem 2rem 0.8rem 1.3rem !important;
            border-bottom: 1px solid rgba(255,255,255,0.36) !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 1.24rem !important;
            line-height: 1.08 !important;
            font-weight: 850 !important;
            letter-spacing: -0.024em !important;
          }
          .st-key-l1_input_sidebar [data-testid="stElementContainer"]:has(.tc-input-sidebar-header) {
            display: block !important;
            visibility: visible !important;
            height: auto !important;
            min-height: 2.7rem !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-heading,
          .st-key-l1_input_sidebar .tc-sidebar-l1-copy,
          .st-key-l1_input_sidebar .tc-sidebar-l1-count,
          .st-key-l1_input_sidebar .tc-sidebar-l1-range {
            display: none !important;
          }

          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list {
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
            gap: 0.90rem !important;
            margin: 0 !important;
            padding: 0 !important;
          }

          /* The wrapper carries no visual treatment; the button itself is the
             complete card.  This prevents nested borders and clipped footers. */
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_taxonomy,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_company_headers,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_revenue_values,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_fte_values,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l0_quartile_table,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_input_l1_quartile_table {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            border: 0 !important;
            border-radius: 0 !important;
            background: transparent !important;
            box-shadow: none !important;
            overflow: visible !important;
            transform: none !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button,
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="secondary"],
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            justify-content: flex-start !important;
            width: 100% !important;
            min-height: 4.72rem !important;
            margin: 0 !important;
            padding: 0.84rem 0.92rem !important;
            border: 1px solid rgba(255,255,255,0.82) !important;
            border-radius: 0.72rem !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            background-image: none !important;
            box-shadow: 0 5px 14px rgba(68,0,0,0.13) !important;
            color: #17191c !important;
            text-align: left !important;
            transition:
              transform 150ms ease,
              border-color 150ms ease,
              box-shadow 150ms ease,
              background-color 150ms ease !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button:hover {
            transform: translateY(-1px) !important;
            border-color: #ffffff !important;
            box-shadow: 0 8px 18px rgba(68,0,0,0.19) !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] {
            border-color: #101214 !important;
            background: #101214 !important;
            background-color: #101214 !important;
            box-shadow: 0 7px 18px rgba(35,0,0,0.25) !important;
            color: #ffffff !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button:disabled {
            opacity: 0.62 !important;
          }

          /* Streamlit renders the label as Markdown.  The bold first line is
             the input name; the second line is the AUTO/MANUAL range. */
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button p {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            color: #cc0000 !important;
            -webkit-text-fill-color: #cc0000 !important;
            font-size: 0.73rem !important;
            line-height: 1.42 !important;
            font-weight: 800 !important;
            letter-spacing: 0 !important;
            text-align: left !important;
            white-space: normal !important;
            overflow-wrap: anywhere !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button p strong {
            display: inline !important;
            color: #15171a !important;
            -webkit-text-fill-color: #15171a !important;
            font-size: 0.98rem !important;
            line-height: 1.26 !important;
            font-weight: 850 !important;
            letter-spacing: -0.01em !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] p {
            color: #ff8f8f !important;
            -webkit-text-fill-color: #ff8f8f !important;
          }
          .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button[kind="primary"] p strong {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }

          .st-key-l1_input_sidebar .tc-sidebar-l1-status,
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-ready,
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-error {
            box-sizing: border-box !important;
            margin: 0.95rem 0 0 !important;
            padding: 0.56rem 0.68rem !important;
            border: 1px solid rgba(255,255,255,0.24) !important;
            border-left: 1px solid rgba(255,255,255,0.24) !important;
            border-radius: 0.58rem !important;
            background: rgba(16,18,20,0.18) !important;
            color: rgba(255,255,255,0.92) !important;
            -webkit-text-fill-color: rgba(255,255,255,0.92) !important;
            font-size: 0.62rem !important;
            line-height: 1.35 !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-ready {
            background: rgba(16,18,20,0.28) !important;
          }
          .st-key-l1_input_sidebar .tc-sidebar-l1-status.is-error {
            background: rgba(68,0,0,0.58) !important;
          }

          /* L1-only peer ordering. Rendered as three compact native buttons
             inside the same visible block as the mapping cards, so Streamlit
             cannot hide it behind a version-specific BaseWeb select wrapper. */
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section {
            box-sizing: border-box !important;
            display: block !important;
            visibility: visible !important;
            width: 100% !important;
            height: auto !important;
            min-height: 0 !important;
            margin: 0.72rem 0 0 !important;
            padding: 0.32rem 0 0.48rem !important;
            border: 0 !important;
            background: transparent !important;
            opacity: 1 !important;
            overflow: visible !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section > div > [data-testid="stVerticalBlock"] {
            display: flex !important;
            visibility: visible !important;
            gap: 0.68rem !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            opacity: 1 !important;
          }
          .st-key-l1_input_sidebar .tc-input-sort-divider {
            display: block !important;
            width: 100% !important;
            height: 1px !important;
            margin: 0 0 0.82rem !important;
            background: rgba(255,255,255,0.58) !important;
          }
          .st-key-l1_input_sidebar .tc-input-sort-heading {
            display: block !important;
            margin: 0 0 0.16rem !important;
            padding: 0 0.08rem !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 1.02rem !important;
            line-height: 1.16 !important;
            font-weight: 900 !important;
            letter-spacing: 0.035em !important;
            text-transform: uppercase !important;
          }
          .st-key-l1_input_sidebar .tc-input-sort-copy {
            display: block !important;
            margin: 0.10rem 0 0 !important;
            padding: 0 0.08rem !important;
            color: rgba(255,255,255,0.88) !important;
            -webkit-text-fill-color: rgba(255,255,255,0.88) !important;
            font-size: 0.72rem !important;
            line-height: 1.38 !important;
            font-weight: 650 !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton {
            display: block !important;
            visibility: visible !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            opacity: 1 !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button[kind="secondary"],
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button[kind="primary"] {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            justify-content: flex-start !important;
            width: 100% !important;
            min-height: 2.55rem !important;
            margin: 0 !important;
            padding: 0.56rem 0.72rem !important;
            border: 1px solid rgba(255,255,255,0.84) !important;
            border-radius: 0.62rem !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            background-image: none !important;
            box-shadow: 0 4px 11px rgba(68,0,0,0.13) !important;
            color: #17191c !important;
            -webkit-text-fill-color: #17191c !important;
            text-align: left !important;
            opacity: 1 !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button:hover {
            border-color: #ffffff !important;
            box-shadow: 0 6px 14px rgba(68,0,0,0.20) !important;
            transform: translateY(-1px) !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button[kind="primary"] {
            border-color: #101214 !important;
            background: #101214 !important;
            background-color: #101214 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button p,
          .st-key-l1_input_sidebar .st-key-sidebar_l1_sort_section div.stButton > button span {
            margin: 0 !important;
            padding: 0 !important;
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
            font-size: 0.82rem !important;
            line-height: 1.24 !important;
            font-weight: 800 !important;
            text-align: left !important;
            white-space: normal !important;
          }

          #tc-input-sidebar-toggle {
            top: 0.82rem !important;
            background: #101214 !important;
            border-color: rgba(255,255,255,0.62) !important;
            color: #ffffff !important;
            box-shadow: 0 7px 18px rgba(38,0,0,0.28) !important;
          }
          #tc-input-sidebar-toggle:hover {
            background: #2a2d32 !important;
            border-color: #ffffff !important;
          }
          #tc-input-sidebar-mini,
          #tc-input-sidebar-mini span {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }

          @media (max-height: 820px) {
            .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
            .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
              padding-top: 0.88rem !important;
              padding-bottom: 0.9rem !important;
            }
            .st-key-l1_input_sidebar .tc-input-sidebar-header {
              min-height: 2.35rem !important;
              margin-bottom: 0.68rem !important;
              padding-bottom: 0.62rem !important;
              font-size: 1.12rem !important;
            }
            .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > [data-testid="stVerticalBlock"],
            .st-key-l1_input_sidebar .st-key-sidebar_l1_input_list > div > [data-testid="stVerticalBlock"] {
              gap: 0.68rem !important;
            }
            .st-key-l1_input_sidebar [class*="st-key-sidebar_l1_input_"] div.stButton > button {
              min-height: 3.85rem !important;
              padding-top: 0.64rem !important;
              padding-bottom: 0.64rem !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )



    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r56: stable segment rail with per-company assignment editing. */
          :root {
            --tc-segment-sidebar-expanded-width: clamp(320px, 22vw, 410px);
            --tc-segment-sidebar-compact-width: 3.55rem;
            --tc-segment-sidebar-current-width: var(--tc-segment-sidebar-expanded-width);
            --tc-segment-sidebar-motion: 220ms cubic-bezier(.2, .8, .2, 1);
          }
          html.tc-segment-sidebar-is-collapsed,
          body.tc-segment-sidebar-is-collapsed {
            --tc-segment-sidebar-current-width: var(--tc-segment-sidebar-compact-width);
          }

          .st-key-l1_segment_sidebar {
            position: fixed !important;
            top: 0 !important;
            bottom: 0 !important;
            left: calc(var(--tc-sidebar-current-width) + var(--tc-input-sidebar-current-width)) !important;
            z-index: 99970 !important;
            display: none !important;
            width: var(--tc-segment-sidebar-current-width) !important;
            min-width: var(--tc-segment-sidebar-current-width) !important;
            max-width: var(--tc-segment-sidebar-current-width) !important;
            height: 100vh !important;
            margin: 0 !important;
            padding: 0 !important;
            overflow: visible !important;
            border: 0 !important;
            border-right: 1px solid #6B6B6B !important;
            border-radius: 0 !important;
            background: #4A4A4A !important;
            box-shadow: 10px 0 28px rgba(0,0,0,0.20) !important;
            transition:
              left var(--tc-sidebar-motion),
              width var(--tc-segment-sidebar-motion),
              min-width var(--tc-segment-sidebar-motion),
              max-width var(--tc-segment-sidebar-motion) !important;
          }
          html.tc-segment-sidebar-is-open .st-key-l1_segment_sidebar,
          body.tc-segment-sidebar-is-open .st-key-l1_segment_sidebar {
            display: block !important;
          }
          .st-key-l1_segment_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_segment_sidebar > div > [data-testid="stVerticalBlock"] {
            position: relative !important;
            height: 100vh !important;
            min-height: 100vh !important;
            max-height: 100vh !important;
            box-sizing: border-box !important;
            gap: 0 !important;
            margin: 0 !important;
            padding: 1.15rem 1.05rem 1.45rem !important;
            overflow-x: hidden !important;
            overflow-y: auto !important;
            scrollbar-width: thin !important;
            scrollbar-color: #B4B4B4 transparent !important;
            transition: opacity 140ms ease, visibility 140ms ease !important;
          }
          html.tc-segment-sidebar-is-collapsed .st-key-l1_segment_sidebar > [data-testid="stVerticalBlock"],
          html.tc-segment-sidebar-is-collapsed .st-key-l1_segment_sidebar > div > [data-testid="stVerticalBlock"],
          body.tc-segment-sidebar-is-collapsed .st-key-l1_segment_sidebar > [data-testid="stVerticalBlock"],
          body.tc-segment-sidebar-is-collapsed .st-key-l1_segment_sidebar > div > [data-testid="stVerticalBlock"] {
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
          }

          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            min-height: 3.25rem !important;
            margin: 0 0 0.9rem !important;
            padding: 0.22rem 1.55rem 0.95rem 0.15rem !important;
            border-bottom: 1px solid #858585 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 1.38rem !important;
            line-height: 1.12 !important;
            font-weight: 900 !important;
            letter-spacing: -0.025em !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-sidebar-summary {
            display: flex !important;
            flex-wrap: wrap !important;
            align-items: center !important;
            gap: 0.28rem 0.42rem !important;
            margin: 0 0 1rem !important;
            padding: 0 0.18rem !important;
            color: #F0F0F0 !important;
            -webkit-text-fill-color: #F0F0F0 !important;
            font-size: 0.78rem !important;
            line-height: 1.42 !important;
            font-weight: 700 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-sidebar-summary strong {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-weight: 900 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-divider {
            width: 100% !important;
            height: 1px !important;
            margin: 1.15rem 0 0.9rem !important;
            background: rgba(255,255,255,0.28) !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-heading {
            margin: 0 0 0.22rem !important;
            padding: 0 0.12rem !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 0.9rem !important;
            line-height: 1.25 !important;
            font-weight: 880 !important;
            letter-spacing: 0.005em !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-copy {
            margin: 0 0 0.72rem !important;
            padding: 0 0.12rem !important;
            color: #E7E7E7 !important;
            -webkit-text-fill-color: #E7E7E7 !important;
            font-size: 0.7rem !important;
            line-height: 1.4 !important;
            font-weight: 620 !important;
          }

          /* Direct per-company assignment editor. No nested forms or expanders. */
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_toggle {
            margin: 0 0 0.82rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_toggle div.stButton > button {
            min-height: 2.65rem !important;
            border: 1px solid rgba(255,255,255,0.44) !important;
            border-radius: 0.65rem !important;
            background: rgba(20,20,20,0.30) !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 0.76rem !important;
            font-weight: 840 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-editor-note {
            margin: 0 0 0.72rem !important;
            padding: 0.68rem 0.75rem !important;
            border: 1px solid rgba(255,255,255,0.20) !important;
            border-radius: 0.62rem !important;
            background: rgba(20,20,20,0.22) !important;
            color: #F3F3F3 !important;
            -webkit-text-fill-color: #F3F3F3 !important;
            font-size: 0.70rem !important;
            line-height: 1.42 !important;
            font-weight: 660 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-client-row {
            display: grid !important;
            grid-template-columns: minmax(0, 0.42fr) minmax(0, 0.58fr) !important;
            align-items: center !important;
            gap: 0.5rem !important;
            width: 100% !important;
            min-height: 2.55rem !important;
            margin: 0 0 0.55rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-fixed,
          .st-key-l1_segment_sidebar .tc-segment-assignment-company,
          .st-key-l1_segment_sidebar .tc-segment-assignment-company-card {
            display: flex !important;
            align-items: center !important;
            min-width: 0 !important;
            min-height: 2.5rem !important;
            padding: 0.45rem 0.65rem !important;
            border-radius: 0.55rem !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            white-space: nowrap !important;
            font-size: 0.74rem !important;
            line-height: 1.2 !important;
            font-weight: 810 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-fixed {
            justify-content: center !important;
            border: 1px solid #E34A4A !important;
            background: #CC0000 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-company,
          .st-key-l1_segment_sidebar .tc-segment-assignment-company-card {
            border: 1px solid #C8C8C8 !important;
            background: #F7F7F7 !important;
            color: #1E1E1E !important;
            -webkit-text-fill-color: #1E1E1E !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] {
            margin: 0 !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] [data-baseweb="input"],
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] [data-baseweb="base-input"] {
            min-height: 2.5rem !important;
            border: 1px solid #C8C8C8 !important;
            border-radius: 0.55rem !important;
            background: #ffffff !important;
            box-shadow: none !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] input {
            min-height: 2.5rem !important;
            padding: 0.42rem 0.62rem !important;
            color: #15171A !important;
            -webkit-text-fill-color: #15171A !important;
            background: transparent !important;
            font-size: 0.75rem !important;
            font-weight: 820 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-row-gap {
            height: 0.50rem !important;
            min-height: 0.50rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_save,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_reset,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_close {
            margin: 0.22rem 0 0.40rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_save div.stButton > button,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_reset div.stButton > button,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_close div.stButton > button {
            min-height: 2.45rem !important;
            padding: 0.42rem 0.35rem !important;
            border-radius: 0.58rem !important;
            font-size: 0.66rem !important;
            line-height: 1.2 !important;
            font-weight: 840 !important;
            white-space: normal !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-sidebar-error {
            margin-top: 0.5rem !important;
            padding: 0.75rem 0.8rem !important;
            border-left: 3px solid #ff6f76 !important;
            border-radius: 0.45rem !important;
            background: #4a2426 !important;
            color: #ffd9dc !important;
            -webkit-text-fill-color: #ffd9dc !important;
            font-size: 0.74rem !important;
            line-height: 1.45 !important;
          }

          .st-key-l1_segment_sidebar .st-key-l1_segment_direction {
            margin: 0 0 1.1rem !important;
            padding: 0 0.12rem 1rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction [data-testid="stHorizontalBlock"] {
            gap: 0.55rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction div.stButton > button {
            min-height: 2.65rem !important;
            padding: 0.52rem 0.55rem !important;
            border: 1px solid #858585 !important;
            border-radius: 0.62rem !important;
            background: #ffffff !important;
            color: #333333 !important;
            -webkit-text-fill-color: #333333 !important;
            font-size: 0.68rem !important;
            line-height: 1.25 !important;
            font-weight: 790 !important;
            white-space: normal !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction div.stButton > button[kind="primary"] {
            border-color: #cc0000 !important;
            background: #cc0000 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction div.stButton > button * {
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
          }

          .st-key-l1_segment_sidebar .st-key-l1_segment_mapping_list {
            overflow: hidden !important;
            margin: 0 0 0.35rem !important;
            padding: 0 !important;
            border: 1px solid #858585 !important;
            border-radius: 0.72rem !important;
            background: #f4f4f4 !important;
            box-shadow: 0 8px 18px rgba(0,0,0,0.16) !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_mapping_list > [data-testid="stVerticalBlock"],
          .st-key-l1_segment_sidebar .st-key-l1_segment_mapping_list > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_mapping_list [data-testid="stElementContainer"] {
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-table-head,
          .st-key-l1_segment_sidebar .tc-segment-table-row {
            display: grid !important;
            grid-template-columns: minmax(5.7rem, 0.82fr) minmax(7rem, 1.18fr) !important;
            align-items: center !important;
            width: 100% !important;
            box-sizing: border-box !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-table-head {
            background: #5c5c5c !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 0.66rem !important;
            font-weight: 850 !important;
            letter-spacing: 0.08em !important;
            text-transform: uppercase !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-table-head > div,
          .st-key-l1_segment_sidebar .tc-segment-cell {
            min-width: 0 !important;
            box-sizing: border-box !important;
            padding: 0.57rem 0.62rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-table-head > div + div,
          .st-key-l1_segment_sidebar .tc-segment-cell + .tc-segment-cell {
            border-left: 1px solid #b4b4b4 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-table-row {
            border-top: 1px solid #d6d6d6 !important;
            background: #ffffff !important;
            color: #242424 !important;
            -webkit-text-fill-color: #242424 !important;
            font-size: 0.72rem !important;
            line-height: 1.28 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_mapping_list [data-testid="stElementContainer"]:nth-child(odd) .tc-segment-table-row {
            background: #f2f2f2 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-table-row.is-client {
            background: #d6d6d6 !important;
            font-weight: 820 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-name {
            display: flex !important;
            align-items: center !important;
            gap: 0.45rem !important;
            font-weight: 780 !important;
          }
          .st-key-l1_segment_sidebar .tc-company-name {
            overflow: hidden !important;
            text-overflow: ellipsis !important;
            white-space: nowrap !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-swatch {
            display: inline-block !important;
            flex: 0 0 auto !important;
            width: 0.72rem !important;
            height: 0.72rem !important;
            border: 1px solid #858585 !important;
            border-radius: 0.18rem !important;
            background: var(--tc-segment-swatch-color, #858585) !important;
          }

          /* v57: spacing polish for the stable per-company editor.
             Apply spacing to the real Streamlit row wrappers so it survives reruns. */
          .st-key-l1_segment_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_segment_sidebar > div > [data-testid="stVerticalBlock"] {
            padding: 1.30rem 1.15rem 1.80rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            margin-bottom: 1.10rem !important;
            padding-bottom: 1.02rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-sidebar-summary {
            margin-bottom: 1.18rem !important;
            line-height: 1.52 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-divider {
            margin-top: 1.48rem !important;
            margin-bottom: 1.05rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-heading {
            margin-bottom: 0.44rem !important;
            font-size: 0.94rem !important;
            line-height: 1.34 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-copy {
            margin-bottom: 1.02rem !important;
            font-size: 0.73rem !important;
            line-height: 1.52 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_toggle {
            margin-bottom: 1.00rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-editor-note {
            margin-bottom: 1.02rem !important;
            padding-top: 0.78rem !important;
            padding-bottom: 0.78rem !important;
            line-height: 1.52 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-client-row {
            min-height: 2.72rem !important;
            margin-bottom: 0.72rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-fixed,
          .st-key-l1_segment_sidebar .tc-segment-assignment-company,
          .st-key-l1_segment_sidebar .tc-segment-assignment-company-card {
            min-height: 2.68rem !important;
            padding-top: 0.54rem !important;
            padding-bottom: 0.54rem !important;
          }
          .st-key-l1_segment_sidebar [data-testid="stHorizontalBlock"]:has([class*="st-key-l1_segment_assignment_field_"]) {
            align-items: stretch !important;
            gap: 0.68rem !important;
            margin: 0 0 0.70rem !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] [data-baseweb="input"],
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] [data-baseweb="base-input"],
          .st-key-l1_segment_sidebar [class*="st-key-l1_segment_assignment_field_"] input {
            min-height: 2.68rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-row-gap {
            display: none !important;
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar [data-testid="stHorizontalBlock"]:has(.st-key-l1_segment_assignment_save) {
            align-items: stretch !important;
            gap: 0.68rem !important;
            margin: 0.74rem 0 1.28rem !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_save,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_reset,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_close {
            margin: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_save div.stButton > button,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_reset div.stButton > button,
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_close div.stButton > button {
            min-height: 2.58rem !important;
            padding-top: 0.50rem !important;
            padding-bottom: 0.50rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction {
            margin-top: 0.18rem !important;
            margin-bottom: 1.30rem !important;
            padding-bottom: 1.10rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction [data-testid="stHorizontalBlock"] {
            gap: 0.72rem !important;
            margin-top: 0.78rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction div.stButton > button {
            min-height: 2.78rem !important;
            padding-top: 0.60rem !important;
            padding-bottom: 0.60rem !important;
          }

          /* v58: bounded 12-row assignment lists and clearer section spacing. */
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block {
            margin: 0 0 1.02rem !important;
            padding: 0.18rem 0.14rem 0 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block .tc-segment-section-heading {
            margin: 0 0 0.52rem !important;
            padding: 0 !important;
            font-size: 0.98rem !important;
            line-height: 1.28 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block .tc-segment-section-copy {
            margin: 0 !important;
            padding: 0 !important;
            font-size: 0.75rem !important;
            line-height: 1.55 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-editor-note {
            margin: 0 0 1.08rem !important;
            padding: 0.82rem 0.82rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll {
            width: 100% !important;
            max-height: 39.75rem !important;
            margin: 0.10rem 0 1.08rem !important;
            padding: 0.20rem 0.52rem 0.28rem 0 !important;
            overflow-x: hidden !important;
            overflow-y: auto !important;
            border: 0 !important;
            border-radius: 0 !important;
            background: transparent !important;
            scrollbar-width: thin !important;
            scrollbar-color: #B4B4B4 rgba(0,0,0,0.12) !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll::-webkit-scrollbar {
            width: 8px !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll::-webkit-scrollbar-track {
            border-radius: 999px !important;
            background: rgba(0,0,0,0.12) !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll::-webkit-scrollbar-thumb {
            border: 2px solid transparent !important;
            border-radius: 999px !important;
            background: #B4B4B4 !important;
            background-clip: padding-box !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll > [data-testid="stVerticalBlock"],
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll [data-testid="stHorizontalBlock"]:has([class*="st-key-l1_segment_assignment_field_"]) {
            align-items: stretch !important;
            gap: 0.68rem !important;
            margin: 0 0 0.68rem !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_scroll .tc-segment-assignment-client-row {
            margin-bottom: 0.68rem !important;
          }
          .st-key-l1_segment_sidebar [data-testid="stHorizontalBlock"]:has(.st-key-l1_segment_assignment_save) {
            margin: 0.90rem 0 1.42rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_toggle {
            margin-top: 0.12rem !important;
            margin-bottom: 1.18rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction {
            margin-top: 0.30rem !important;
          }

          #tc-segment-sidebar-toggle {
            position: absolute;
            top: 0.82rem;
            right: -0.96rem;
            z-index: 100006;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.92rem;
            height: 1.92rem;
            margin: 0;
            padding: 0;
            border: 1px solid rgba(255,255,255,0.62);
            border-radius: 999px;
            background: #101214;
            color: #ffffff;
            box-shadow: 0 7px 18px rgba(0,0,0,0.32);
            cursor: pointer;
            outline: none;
          }
          #tc-segment-sidebar-toggle:hover {
            background: #cc0000;
            border-color: #ffffff;
          }
          #tc-segment-sidebar-toggle .tc-segment-sidebar-toggle-icon {
            width: 0.9rem;
            height: 1rem;
            pointer-events: none;
          }
          #tc-segment-sidebar-mini {
            position: absolute;
            top: 3.55rem;
            left: 50%;
            z-index: 100004;
            display: none;
            align-items: center;
            justify-content: center;
            width: 2.08rem;
            height: 2.08rem;
            transform: translateX(-50%);
            border: 1px solid rgba(255,255,255,0.78);
            border-radius: 0.72rem;
            background: #ffffff;
            color: #4A4A4A;
            box-shadow: 0 5px 14px rgba(0,0,0,0.20);
            user-select: none;
          }
          .tc-segment-sidebar-mini-title {
            font-size: 0.66rem;
            line-height: 1;
            font-weight: 950;
            letter-spacing: 0.02em;
          }
          #tc-segment-sidebar-mini-actions {
            position: absolute;
            top: 6.20rem;
            left: 50%;
            z-index: 100005;
            display: none;
            width: 2.45rem;
            transform: translateX(-50%);
            flex-direction: column;
            align-items: center;
            gap: 0.54rem;
          }
          html.tc-segment-sidebar-is-collapsed #tc-segment-sidebar-mini,
          body.tc-segment-sidebar-is-collapsed #tc-segment-sidebar-mini,
          html.tc-segment-sidebar-is-collapsed #tc-segment-sidebar-mini-actions,
          body.tc-segment-sidebar-is-collapsed #tc-segment-sidebar-mini-actions {
            display: flex;
          }
          .tc-segment-sidebar-mini-action {
            position: relative;
            box-sizing: border-box;
            display: inline-flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 0.12rem;
            width: 2.34rem;
            height: 2.70rem;
            margin: 0;
            padding: 0.30rem 0.18rem 0.24rem;
            border: 1px solid rgba(255,255,255,0.34);
            border-radius: 0.72rem;
            background: #292C2F;
            color: #ffffff;
            box-shadow: 0 5px 12px rgba(0,0,0,0.22);
            cursor: pointer;
            outline: none;
            transition: transform 120ms ease, background 120ms ease, border-color 120ms ease;
          }
          .tc-segment-sidebar-mini-action svg {
            width: 1.12rem;
            height: 1.12rem;
            flex: 0 0 auto;
          }
          .tc-segment-sidebar-mini-action-label {
            font-size: 0.43rem;
            line-height: 1;
            font-weight: 900;
            letter-spacing: 0.01em;
            white-space: nowrap;
          }
          .tc-segment-sidebar-mini-action:hover:not(:disabled) {
            transform: translateY(-1px);
            border-color: #ffffff;
            background: #1A1D20;
          }
          .tc-segment-sidebar-mini-action.is-selected {
            border-color: #ffffff;
            background: #cc0000;
            box-shadow: 0 6px 15px rgba(66,0,0,0.34);
          }
          .tc-segment-sidebar-mini-status {
            position: absolute;
            top: 0.25rem;
            right: 0.25rem;
            display: none;
            width: 0.34rem;
            height: 0.34rem;
            border: 1px solid #ffffff;
            border-radius: 999px;
            background: #31b568;
          }
          .tc-segment-sidebar-mini-action.is-selected .tc-segment-sidebar-mini-status {
            display: block;
          }
          .tc-segment-sidebar-mini-action:disabled {
            cursor: default;
            opacity: 0.40;
          }

          /* Collapsed sort rail must contain only the injected compact controls.
             Hide the entire normal Streamlit content tree so no headings/copy
             can wrap vertically behind the icons. */
          html.tc-segment-sidebar-is-collapsed .st-key-l1_segment_sidebar > :not(#tc-segment-sidebar-toggle):not(#tc-segment-sidebar-mini):not(#tc-segment-sidebar-mini-actions),
          body.tc-segment-sidebar-is-collapsed .st-key-l1_segment_sidebar > :not(#tc-segment-sidebar-toggle):not(#tc-segment-sidebar-mini):not(#tc-segment-sidebar-mini-actions) {
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
          }

          html.tc-segment-sidebar-is-open .st-key-canvas,
          body.tc-segment-sidebar-is-open .st-key-canvas {
            width: calc(100% - var(--tc-input-sidebar-current-width) - var(--tc-segment-sidebar-current-width)) !important;
            margin-left: calc(var(--tc-input-sidebar-current-width) + var(--tc-segment-sidebar-current-width)) !important;
          }

          @media (prefers-reduced-motion: reduce) {
            .st-key-l1_segment_sidebar,
            #tc-segment-sidebar-toggle,
            .st-key-canvas {
              transition: none !important;
            }
          }
          @media (max-width: 1100px) {
            :root {
              --tc-segment-sidebar-expanded-width: min(360px, 43vw);
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r67: canvas-level workbook upload and preparation indicator. */
          .st-key-workbook_load_status {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            transition: margin-left var(--tc-input-sidebar-motion, 160ms ease),
                        width var(--tc-input-sidebar-motion, 160ms ease) !important;
          }
          .st-key-workbook_load_status > [data-testid="stVerticalBlock"],
          .st-key-workbook_load_status > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          html.tc-input-sidebar-is-open .st-key-workbook_load_status,
          body.tc-input-sidebar-is-open .st-key-workbook_load_status {
            width: calc(100% - var(--tc-input-sidebar-current-width)) !important;
            margin-left: var(--tc-input-sidebar-current-width) !important;
          }
          html.tc-segment-sidebar-is-open .st-key-workbook_load_status,
          body.tc-segment-sidebar-is-open .st-key-workbook_load_status {
            width: calc(100% - var(--tc-input-sidebar-current-width) - var(--tc-segment-sidebar-current-width)) !important;
            margin-left: calc(var(--tc-input-sidebar-current-width) + var(--tc-segment-sidebar-current-width)) !important;
          }

          .tc-upload-loader-shell {
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: calc(100vh - 0.65rem);
            width: 100%;
            padding: 1.1rem 1rem 1.4rem;
            box-sizing: border-box;
            background:
              radial-gradient(circle at 18% 18%, rgba(204,0,0,0.055), transparent 34%),
              radial-gradient(circle at 84% 78%, rgba(44,52,61,0.06), transparent 38%),
              #f6f7f8;
          }
          .tc-upload-loader-card {
            position: relative;
            width: min(760px, 92%);
            overflow: hidden;
            border: 1px solid #e1e4e8;
            border-radius: 20px;
            background: rgba(255,255,255,0.96);
            box-shadow: 0 22px 65px rgba(25,30,35,0.10), 0 2px 8px rgba(25,30,35,0.05);
            padding: 1.55rem 1.65rem 1.35rem;
          }
          .tc-upload-loader-card::before {
            content: "";
            position: absolute;
            inset: 0 0 auto 0;
            height: 4px;
            background: linear-gradient(90deg, #cc0000 0%, #e4494f 48%, #cc0000 100%);
          }
          .tc-upload-loader-main {
            display: grid;
            grid-template-columns: 70px minmax(0, 1fr) auto;
            gap: 1.15rem;
            align-items: center;
          }
          .tc-upload-loader-visual {
            position: relative;
            display: grid;
            place-items: center;
            width: 66px;
            height: 66px;
          }
          .tc-upload-loader-orbit {
            position: absolute;
            inset: 0;
            border-radius: 50%;
            border: 3px solid #eceff2;
            border-top-color: #cc0000;
            border-right-color: #cc0000;
            animation: tc-upload-loader-spin 1.05s linear infinite;
          }
          .tc-upload-loader-icon {
            display: grid;
            place-items: center;
            width: 44px;
            height: 44px;
            border-radius: 12px;
            background: #17191d;
            color: #ffffff;
            font-size: 0.64rem;
            font-weight: 900;
            letter-spacing: 0.045em;
            box-shadow: 0 8px 18px rgba(23,25,29,0.16);
          }
          .tc-upload-loader-card.is-complete .tc-upload-loader-orbit,
          .tc-upload-loader-card.is-error .tc-upload-loader-orbit {
            animation: none;
          }
          .tc-upload-loader-card.is-complete .tc-upload-loader-orbit {
            border-color: #1e8a5a;
          }
          .tc-upload-loader-card.is-complete .tc-upload-loader-icon {
            background: #1e8a5a;
            font-size: 1.1rem;
          }
          .tc-upload-loader-card.is-error .tc-upload-loader-orbit {
            border-color: #cc0000;
          }
          .tc-upload-loader-card.is-error .tc-upload-loader-icon {
            background: #cc0000;
            font-size: 1.1rem;
          }
          .tc-upload-loader-copy { min-width: 0; }
          .tc-upload-loader-kicker {
            margin: 0 0 0.28rem;
            color: #cc0000;
            font-size: 0.68rem;
            font-weight: 900;
            letter-spacing: 0.12em;
          }
          .tc-upload-loader-title {
            margin: 0;
            color: #17191d;
            font-size: clamp(1.3rem, 2vw, 1.72rem);
            font-weight: 820;
            line-height: 1.12;
            letter-spacing: -0.025em;
          }
          .tc-upload-loader-detail {
            margin-top: 0.42rem;
            color: #5f6873;
            font-size: 0.88rem;
            line-height: 1.45;
          }
          .tc-upload-loader-file {
            margin-top: 0.56rem;
            color: #8a929b;
            font-size: 0.72rem;
            font-weight: 650;
            line-height: 1.3;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .tc-upload-loader-metric {
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            gap: 0.2rem;
            min-width: 74px;
          }
          .tc-upload-loader-metric strong {
            color: #17191d;
            font-size: 1.42rem;
            font-weight: 850;
            line-height: 1;
          }
          .tc-upload-loader-metric span {
            color: #8a929b;
            font-size: 0.68rem;
            font-weight: 650;
          }
          .tc-upload-loader-track {
            position: relative;
            height: 8px;
            margin: 1.28rem 0 1.05rem;
            overflow: hidden;
            border-radius: 999px;
            background: #eceff2;
          }
          .tc-upload-loader-track > span {
            display: block;
            height: 100%;
            min-width: 2%;
            border-radius: inherit;
            background: linear-gradient(90deg, #a90000 0%, #cc0000 48%, #e05259 100%);
            box-shadow: 0 0 12px rgba(204,0,0,0.22);
            transition: width 260ms ease;
          }
          .tc-upload-loader-card.is-complete .tc-upload-loader-track > span {
            background: linear-gradient(90deg, #166c46 0%, #1e8a5a 100%);
          }
          .tc-upload-loader-steps {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 0.7rem;
          }
          .tc-upload-loader-step {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.45rem;
            min-height: 2.25rem;
            padding: 0.46rem 0.58rem;
            border: 1px solid #e5e8eb;
            border-radius: 10px;
            background: #f8f9fa;
            color: #8a929b;
            font-size: 0.72rem;
            font-weight: 760;
          }
          .tc-upload-loader-step-marker {
            display: grid;
            place-items: center;
            flex: 0 0 20px;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            background: #e7eaee;
            color: #757e88;
            font-size: 0.66rem;
            font-weight: 900;
          }
          .tc-upload-loader-step.is-active {
            border-color: rgba(204,0,0,0.30);
            background: #fff7f7;
            color: #7e0000;
          }
          .tc-upload-loader-step.is-active .tc-upload-loader-step-marker {
            background: #cc0000;
            color: #ffffff;
            box-shadow: 0 0 0 4px rgba(204,0,0,0.10);
          }
          .tc-upload-loader-step.is-done {
            border-color: #d4e8dd;
            background: #f5fbf8;
            color: #256246;
          }
          .tc-upload-loader-step.is-done .tc-upload-loader-step-marker {
            background: #1e8a5a;
            color: #ffffff;
          }
          @keyframes tc-upload-loader-spin {
            to { transform: rotate(360deg); }
          }
          @media (max-width: 760px) {
            .tc-upload-loader-shell { padding: 0.75rem; }
            .tc-upload-loader-card { width: 100%; padding: 1.25rem 1.05rem 1.05rem; }
            .tc-upload-loader-main { grid-template-columns: 58px minmax(0, 1fr); }
            .tc-upload-loader-visual { width: 56px; height: 56px; }
            .tc-upload-loader-icon { width: 38px; height: 38px; }
            .tc-upload-loader-metric { display: none; }
            .tc-upload-loader-steps { grid-template-columns: repeat(2, minmax(0, 1fr)); }
          }
          @media (prefers-reduced-motion: reduce) {
            .tc-upload-loader-orbit { animation: none !important; }
            .tc-upload-loader-track > span { transition: none !important; }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r75: sorting lives in its own third rail. The red Inputs rail only
             contains one launcher at the bottom. */
          .st-key-l1_input_sidebar .st-key-input_sort_launcher {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 1.15rem 0 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .st-key-input_sort_launcher > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-input_sort_launcher > div > [data-testid="stVerticalBlock"] {
            gap: 0.72rem !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .tc-input-sort-launch-divider {
            display: block !important;
            width: 100% !important;
            height: 1px !important;
            margin: 0 0 0.12rem !important;
            background: rgba(255,255,255,0.62) !important;
          }
          .st-key-l1_input_sidebar .st-key-input_sort_launcher div.stButton,
          .st-key-l1_input_sidebar .st-key-input_sort_launcher div.stButton > button {
            width: 100% !important;
            margin: 0 !important;
          }
          .st-key-l1_input_sidebar .st-key-input_sort_launcher div.stButton > button {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            justify-content: flex-start !important;
            min-height: 4.25rem !important;
            padding: 0.76rem 0.9rem !important;
            border: 1px solid rgba(255,255,255,0.88) !important;
            border-radius: 0.72rem !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            background-image: none !important;
            box-shadow: 0 6px 15px rgba(68,0,0,0.17) !important;
            color: #17191c !important;
            text-align: left !important;
          }
          .st-key-l1_input_sidebar .st-key-input_sort_launcher div.stButton > button:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 9px 20px rgba(68,0,0,0.23) !important;
          }
          .st-key-l1_input_sidebar .st-key-input_sort_launcher div.stButton > button p {
            width: 100% !important;
            margin: 0 !important;
            color: #cc0000 !important;
            -webkit-text-fill-color: #cc0000 !important;
            font-size: 0.72rem !important;
            line-height: 1.38 !important;
            font-weight: 800 !important;
            text-align: left !important;
          }
          .st-key-l1_input_sidebar .st-key-input_sort_launcher div.stButton > button p strong {
            display: inline !important;
            color: #15171a !important;
            -webkit-text-fill-color: #15171a !important;
            font-size: 1rem !important;
            line-height: 1.24 !important;
            font-weight: 900 !important;
          }

          /* Universal mode choices in the third grey rail. */
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes {
            box-sizing: border-box !important;
            width: calc(100% - 1.40rem) !important;
            margin: 0.35rem 0.70rem 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes > [data-testid="stVerticalBlock"],
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes > div > [data-testid="stVerticalBlock"] {
            gap: 0.86rem !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button,
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button[kind="secondary"],
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button[kind="primary"] {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            justify-content: flex-start !important;
            width: 100% !important;
            min-height: 3.25rem !important;
            margin: 0 !important;
            padding: 0.68rem 0.82rem !important;
            border: 1px solid #d4d4d4 !important;
            border-radius: 0.68rem !important;
            background: #f8f8f8 !important;
            background-color: #f8f8f8 !important;
            background-image: none !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.13) !important;
            color: #17191c !important;
            -webkit-text-fill-color: #17191c !important;
            text-align: left !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button:hover {
            border-color: #ffffff !important;
            background: #ffffff !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 7px 16px rgba(0,0,0,0.18) !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button[kind="primary"] {
            border-color: #cc0000 !important;
            background: #cc0000 !important;
            background-color: #cc0000 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            box-shadow: 0 7px 18px rgba(82,0,0,0.28) !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button:disabled {
            opacity: 0.58 !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button p,
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button span {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
            font-size: 0.82rem !important;
            line-height: 1.32 !important;
            font-weight: 820 !important;
            text-align: left !important;
            white-space: normal !important;
          }
          .st-key-l1_segment_sidebar .st-key-sort_sidebar_modes div.stButton > button p strong {
            font-size: 0.9rem !important;
            font-weight: 900 !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r77: v59-proven third rail restored. The Company sort launcher is
             a direct Streamlit button and opens the complete grey rail. */
          .st-key-l1_input_sidebar .tc-input-sort-launch-divider {
            display: block !important;
            width: 100% !important;
            height: 1px !important;
            margin: 1.05rem 0 0.72rem !important;
            background: rgba(255,255,255,0.70) !important;
          }
          .st-key-l1_input_sidebar .st-key-open_sort_sidebar_v77 {
            box-sizing: border-box !important;
            display: block !important;
            visibility: visible !important;
            position: sticky !important;
            bottom: 0.78rem !important;
            z-index: 90 !important;
            /* Match the horizontal inset used by the input cards above. */
            width: calc(100% - 1.56rem) !important;
            min-height: 4.25rem !important;
            margin: 0 0.78rem !important;
            padding: 0 !important;
            opacity: 1 !important;
          }
          .st-key-l1_input_sidebar .st-key-open_sort_sidebar_v77 div.stButton,
          .st-key-l1_input_sidebar .st-key-open_sort_sidebar_v77 div.stButton > button {
            display: block !important;
            visibility: visible !important;
            width: 100% !important;
            margin: 0 !important;
            opacity: 1 !important;
          }
          .st-key-l1_input_sidebar .st-key-open_sort_sidebar_v77 div.stButton > button {
            box-sizing: border-box !important;
            min-height: 4.25rem !important;
            padding: 0.76rem 0.9rem !important;
            border: 1px solid rgba(255,255,255,0.94) !important;
            border-radius: 0.76rem !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            background-image: none !important;
            box-shadow: 0 7px 18px rgba(68,0,0,0.22) !important;
            color: #15171a !important;
            text-align: left !important;
          }
          .st-key-l1_input_sidebar .st-key-open_sort_sidebar_v77 div.stButton > button p {
            width: 100% !important;
            margin: 0 !important;
            color: #cc0000 !important;
            -webkit-text-fill-color: #cc0000 !important;
            font-size: 0.72rem !important;
            line-height: 1.38 !important;
            font-weight: 820 !important;
            text-align: left !important;
          }
          .st-key-l1_input_sidebar .st-key-open_sort_sidebar_v77 div.stButton > button p strong {
            color: #15171a !important;
            -webkit-text-fill-color: #15171a !important;
            font-size: 1rem !important;
            line-height: 1.2 !important;
            font-weight: 900 !important;
          }
          /* Once opened, the third rail owns its 2.5% collapsed state and can
             be expanded again with its own arrow, matching the stable v59 UI. */
          html.tc-segment-sidebar-is-open .st-key-l1_segment_sidebar,
          body.tc-segment-sidebar-is-open .st-key-l1_segment_sidebar {
            display: block !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r78: segment mode is a real display/edit workspace.  The universal
             mode buttons are direct Streamlit children so no nested container
             can consume the remaining rail height. */
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] {
            box-sizing: border-box !important;
            display: block !important;
            width: calc(100% - 1.40rem) !important;
            margin: 0 0.70rem 0.86rem 0.70rem !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton,
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button {
            width: 100% !important;
            margin: 0 !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button,
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button[kind="secondary"],
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button[kind="primary"] {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            justify-content: flex-start !important;
            width: 100% !important;
            min-height: 3.2rem !important;
            padding: 0.68rem 0.86rem !important;
            border: 1px solid #d3d3d3 !important;
            border-radius: 0.72rem !important;
            background: #f8f8f8 !important;
            background-color: #f8f8f8 !important;
            background-image: none !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.14) !important;
            color: #17191c !important;
            -webkit-text-fill-color: #17191c !important;
            text-align: left !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button:hover {
            border-color: #ffffff !important;
            background: #ffffff !important;
            background-color: #ffffff !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 7px 17px rgba(0,0,0,0.19) !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button[kind="primary"] {
            border-color: #cc0000 !important;
            background: #cc0000 !important;
            background-color: #cc0000 !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            box-shadow: 0 7px 18px rgba(72,0,0,0.30) !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button p,
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button span {
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
            font-size: 0.79rem !important;
            line-height: 1.34 !important;
            font-weight: 820 !important;
            text-align: left !important;
            white-space: normal !important;
          }
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] div.stButton > button p strong {
            font-size: 0.91rem !important;
            font-weight: 900 !important;
          }

          .st-key-l1_segment_sidebar .tc-segment-section-card {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0 0 0.82rem !important;
            padding: 0.78rem 0.82rem !important;
            border: 1px solid rgba(255,255,255,0.22) !important;
            border-radius: 0.72rem !important;
            background: rgba(20,20,20,0.18) !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-card .tc-segment-section-heading {
            margin-bottom: 0.32rem !important;
            padding: 0 !important;
            font-size: 0.94rem !important;
            line-height: 1.28 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-section-card .tc-segment-section-copy {
            margin: 0 !important;
            padding: 0 !important;
            font-size: 0.72rem !important;
            line-height: 1.48 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-detected-banner {
            display: flex !important;
            flex-wrap: wrap !important;
            align-items: center !important;
            gap: 0.30rem 0.45rem !important;
            width: 100% !important;
            margin: 0 0 0.92rem !important;
            padding: 0.64rem 0.74rem !important;
            border: 1px solid rgba(255,255,255,0.22) !important;
            border-radius: 0.66rem !important;
            background: rgba(20,20,20,0.24) !important;
            color: #f5f5f5 !important;
            -webkit-text-fill-color: #f5f5f5 !important;
            font-size: 0.73rem !important;
            line-height: 1.35 !important;
            font-weight: 700 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-detected-banner strong {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-weight: 900 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-mode-summary {
            margin: 0.20rem 0 1.05rem !important;
            padding: 0.78rem 0.82rem !important;
            border-left: 3px solid #cc0000 !important;
            border-radius: 0.56rem !important;
            background: rgba(20,20,20,0.24) !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-mode-summary-title {
            margin: 0 0 0.24rem !important;
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 0.86rem !important;
            line-height: 1.25 !important;
            font-weight: 900 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-mode-summary-copy {
            color: #ececec !important;
            -webkit-text-fill-color: #ececec !important;
            font-size: 0.71rem !important;
            line-height: 1.45 !important;
            font-weight: 620 !important;
          }
          .st-key-l1_segment_sidebar iframe {
            display: block !important;
            width: 100% !important;
            margin: 0.12rem 0 0.90rem !important;
            border: 0 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-change-mode-card {
            margin-top: 0.10rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction {
            margin-bottom: 0.45rem !important;
            padding-bottom: 0.35rem !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction [data-testid="stHorizontalBlock"] {
            margin-top: 0.72rem !important;
            gap: 0.68rem !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r79: unify all three left-side rails on one header/content rhythm
             and de-clutter the expanded Segment Order workspace. */
          :root {
            --tc-rail-header-height: 3.72rem;
            --tc-rail-title-size: 1.30rem;
            --tc-rail-title-line-height: 1.08;
            --tc-rail-header-bottom-gap: 0.96rem;
            --tc-segment-section-gap: 1.02rem;
          }

          /* Primary rail: keep the TC badge, but place the masthead on the same
             vertical rhythm and title scale as Inputs and Segment Order. */
          [data-testid="stSidebar"] .st-key-control_panel {
            padding-top: 0.78rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand {
            box-sizing: border-box !important;
            min-height: var(--tc-rail-header-height) !important;
            margin: 0 0 var(--tc-rail-header-bottom-gap) !important;
            padding: 0 0 0.82rem !important;
            align-items: center !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-mark {
            width: 2.55rem !important;
            height: 2.55rem !important;
            flex: 0 0 2.55rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-title {
            font-size: var(--tc-rail-title-size) !important;
            line-height: var(--tc-rail-title-line-height) !important;
            font-weight: 880 !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand-subtitle {
            margin-top: 0.26rem !important;
          }

          /* Inputs rail: same top inset, header height, baseline, and divider. */
          .st-key-l1_input_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar > div > [data-testid="stVerticalBlock"] {
            padding-top: 0.78rem !important;
          }
          .st-key-l1_input_sidebar .tc-input-sidebar-header {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            min-height: var(--tc-rail-header-height) !important;
            margin: 0 0 var(--tc-rail-header-bottom-gap) !important;
            padding: 0 2rem 0.82rem 0.42rem !important;
            border-bottom: 1px solid rgba(255,255,255,0.36) !important;
            font-size: var(--tc-rail-title-size) !important;
            line-height: var(--tc-rail-title-line-height) !important;
            font-weight: 880 !important;
          }
          .st-key-l1_input_sidebar [data-testid="stElementContainer"]:has(.tc-input-sidebar-header) {
            min-height: var(--tc-rail-header-height) !important;
          }

          /* Segment rail: mirror the Inputs header instead of using a separate
             indentation/height system. */
          .st-key-l1_segment_sidebar > [data-testid="stVerticalBlock"],
          .st-key-l1_segment_sidebar > div > [data-testid="stVerticalBlock"] {
            padding: 0.78rem 0.90rem 1.80rem !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            min-height: var(--tc-rail-header-height) !important;
            margin: 0 0 var(--tc-rail-header-bottom-gap) !important;
            padding: 0 2rem 0.82rem 0.42rem !important;
            border-bottom: 1px solid rgba(255,255,255,0.36) !important;
            font-size: var(--tc-rail-title-size) !important;
            line-height: var(--tc-rail-title-line-height) !important;
            font-weight: 880 !important;
            letter-spacing: -0.024em !important;
          }

          /* Compact the rail-level metadata into one deliberately aligned row. */
          .st-key-l1_segment_sidebar .tc-segment-sidebar-summary {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0 0 0.78rem !important;
            padding: 0 0.42rem !important;
            gap: 0.26rem 0.38rem !important;
            font-size: 0.72rem !important;
            line-height: 1.35 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-detected-banner {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0 0 var(--tc-segment-section-gap) !important;
            padding: 0.62rem 0.76rem !important;
          }

          /* The embedded mapping already has its own title. Hiding the outer
             duplicate copy removes the stacked-heading clutter above the list. */
          .st-key-l1_segment_sidebar .tc-segment-mapping-card {
            display: none !important;
          }
          .st-key-l1_segment_sidebar iframe {
            box-sizing: border-box !important;
            display: block !important;
            width: 100% !important;
            margin: 0 0 var(--tc-segment-section-gap) !important;
            padding: 0 !important;
          }

          /* Section headings below the mapping use the same card geometry as
             the mapping shell, rather than floating text separated by full-width rules. */
          .st-key-l1_segment_sidebar .tc-segment-section-divider {
            display: none !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block,
          .st-key-l1_segment_sidebar .tc-segment-section-card.tc-segment-direction-card,
          .st-key-l1_segment_sidebar .tc-segment-section-card.tc-segment-change-mode-card {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0 0 0.78rem !important;
            padding: 0.76rem 0.82rem !important;
            border: 1px solid rgba(255,255,255,0.22) !important;
            border-radius: 0.72rem !important;
            background: rgba(20,20,20,0.18) !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block {
            margin-top: 0 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block .tc-segment-section-heading,
          .st-key-l1_segment_sidebar .tc-segment-section-card.tc-segment-direction-card .tc-segment-section-heading,
          .st-key-l1_segment_sidebar .tc-segment-section-card.tc-segment-change-mode-card .tc-segment-section-heading {
            margin: 0 0 0.30rem !important;
            padding: 0 !important;
            font-size: 0.92rem !important;
            line-height: 1.28 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-assignment-heading-block .tc-segment-section-copy,
          .st-key-l1_segment_sidebar .tc-segment-section-card.tc-segment-direction-card .tc-segment-section-copy,
          .st-key-l1_segment_sidebar .tc-segment-section-card.tc-segment-change-mode-card .tc-segment-section-copy {
            margin: 0 !important;
            padding: 0 !important;
            font-size: 0.72rem !important;
            line-height: 1.46 !important;
          }

          /* Keep controls on the same visual grid as their explanatory cards. */
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_toggle,
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction {
            box-sizing: border-box !important;
            width: 100% !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_assignment_toggle {
            margin: 0 0 var(--tc-segment-section-gap) !important;
          }

          /* Compact workspace navigation: visible whenever Segment Order is open.
             It returns to Company Sort without changing the active sort mode. */
          .st-key-l1_segment_sidebar .st-key-l1_segment_workspace_back {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0.10rem 0 0.72rem !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_workspace_back div.stButton {
            width: auto !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_workspace_back div.stButton > button {
            width: auto !important;
            min-width: 5.35rem !important;
            min-height: 2.00rem !important;
            padding: 0.30rem 0.72rem !important;
            border: 1px solid rgba(255,255,255,0.30) !important;
            border-radius: 999px !important;
            background: rgba(20,20,20,0.18) !important;
            box-shadow: none !important;
            color: #F5F5F5 !important;
            -webkit-text-fill-color: #F5F5F5 !important;
            font-size: 0.70rem !important;
            line-height: 1 !important;
            font-weight: 820 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_workspace_back div.stButton > button:hover {
            border-color: rgba(255,255,255,0.52) !important;
            background: rgba(20,20,20,0.28) !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_workspace_back div.stButton > button p,
          .st-key-l1_segment_sidebar .st-key-l1_segment_workspace_back div.stButton > button span {
            margin: 0 !important;
            color: inherit !important;
            -webkit-text-fill-color: inherit !important;
            font-size: inherit !important;
            line-height: inherit !important;
            font-weight: inherit !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction {
            margin: 0 0 var(--tc-segment-section-gap) !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .st-key-l1_segment_direction [data-testid="stHorizontalBlock"] {
            margin-top: 0 !important;
          }
          .st-key-l1_segment_sidebar .tc-segment-mode-summary {
            box-sizing: border-box !important;
            width: 100% !important;
            margin: 0.08rem 0 var(--tc-segment-section-gap) !important;
          }

          /* Short displays should preserve the same header system instead of
             shrinking only the Inputs title. */
          @media (max-height: 820px) {
            [data-testid="stSidebar"] .tc-panel-brand,
            .st-key-l1_input_sidebar .tc-input-sidebar-header,
            .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
              min-height: 3.35rem !important;
              margin-bottom: 0.72rem !important;
              padding-bottom: 0.68rem !important;
            }
            [data-testid="stSidebar"] .tc-panel-brand-title,
            .st-key-l1_input_sidebar .tc-input-sidebar-header,
            .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
              font-size: 1.18rem !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r81: Company Sort home uses one consistent content grid.
             The metadata, intro card, sort buttons, and active-mode summary
             now share the same horizontal inset and an intentional vertical rhythm. */
          .st-key-l1_segment_sidebar {
            --tc-sort-home-inline: 0.70rem;
            --tc-sort-home-width: calc(100% - 1.40rem);
          }

          /* r82: the Company Sort top area is one real DOM grid. This avoids
             margin collapse between separate Streamlit markdown wrappers and
             makes the vertical rhythm deterministic at every viewport height. */
          .st-key-l1_segment_sidebar .tc-sort-home-topstack {
            box-sizing: border-box !important;
            display: grid !important;
            grid-template-columns: minmax(0, 1fr) !important;
            row-gap: 0.86rem !important;
            width: 100% !important;
            margin: 0 0 0.98rem !important;
            padding: 0 !important;
          }
          .st-key-l1_segment_sidebar .tc-sort-home-topstack .tc-segment-sidebar-header {
            width: 100% !important;
            margin: 0 !important;
            padding-bottom: 0.82rem !important;
          }
          .st-key-l1_segment_sidebar .tc-sort-home-topstack .tc-segment-sidebar-summary {
            width: var(--tc-sort-home-width) !important;
            min-height: 1.18rem !important;
            margin: 0 var(--tc-sort-home-inline) !important;
            padding: 0 0.12rem !important;
          }
          .st-key-l1_segment_sidebar .tc-sort-home-topstack .tc-segment-section-card.tc-segment-mode-card {
            width: var(--tc-sort-home-width) !important;
            margin: 0 var(--tc-sort-home-inline) 0.72rem !important;
            padding: 0.78rem 0.82rem !important;
          }

          /* Keep the four choices on the same grid with calm, even separation. */
          .st-key-l1_segment_sidebar [class*="st-key-sort_sidebar_mode_"] {
            width: var(--tc-sort-home-width) !important;
            margin-left: var(--tc-sort-home-inline) !important;
            margin-right: var(--tc-sort-home-inline) !important;
            margin-bottom: 0.90rem !important;
          }

          /* Active-mode status card: same width and edges as every sort button. */
          .st-key-l1_segment_sidebar .tc-segment-mode-summary {
            box-sizing: border-box !important;
            width: var(--tc-sort-home-width) !important;
            margin: 0.14rem var(--tc-sort-home-inline) 1.00rem !important;
            padding: 0.78rem 0.82rem !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r84: compact Inputs rail mirrors the primary rail with useful,
             clickable controls instead of a vertical wordmark. */
          #tc-input-sidebar-mini {
            top: 3.55rem !important;
            width: 2.05rem !important;
            height: 2.05rem !important;
            border: 1px solid rgba(255,255,255,0.55) !important;
            border-radius: 0.66rem !important;
            background: rgba(255,255,255,0.96) !important;
            color: #c90012 !important;
            box-shadow: 0 7px 18px rgba(58,0,10,0.24) !important;
          }
          #tc-input-sidebar-mini span,
          #tc-input-sidebar-mini .tc-input-sidebar-mini-title {
            color: #c90012 !important;
            -webkit-text-fill-color: #c90012 !important;
            font-size: 0.58rem !important;
            line-height: 1 !important;
            font-weight: 950 !important;
            letter-spacing: 0.04em !important;
          }
          #tc-input-sidebar-mini .tc-input-sidebar-mini-subtitle {
            display: none !important;
          }

          #tc-input-sidebar-mini-actions {
            position: absolute;
            top: 6.25rem;
            left: 50%;
            z-index: 100004;
            display: none;
            width: calc(var(--tc-input-sidebar-compact-width) - 4px);
            max-height: calc(100vh - 7rem);
            transform: translateX(-50%);
            flex-direction: column;
            align-items: center;
            gap: 0.48rem;
            padding: 0 0 0.55rem;
            overflow: hidden;
            pointer-events: auto;
            box-sizing: border-box;
          }
          html.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini-actions,
          body.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini-actions {
            display: flex;
          }

          .tc-input-sidebar-mini-action {
            position: relative;
            display: inline-flex;
            flex: 0 0 auto;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 0.08rem;
            width: min(34px, calc(var(--tc-input-sidebar-compact-width) - 6px));
            height: 39px;
            min-width: 0;
            margin: 0;
            padding: 4px 1px 3px;
            border: 1px solid rgba(255,255,255,0.30);
            border-radius: 10px;
            background: rgba(17,19,22,0.94);
            color: #ffffff;
            box-shadow: 0 6px 14px rgba(55,0,9,0.24);
            cursor: pointer;
            outline: none;
            box-sizing: border-box;
            transition: transform 140ms ease, background-color 140ms ease,
                        border-color 140ms ease, box-shadow 140ms ease,
                        opacity 140ms ease;
          }
          .tc-input-sidebar-mini-action svg {
            display: block;
            width: 15px;
            height: 15px;
            flex: 0 0 auto;
            color: currentColor;
          }
          .tc-input-sidebar-mini-action-label {
            display: block;
            max-width: 30px;
            overflow: hidden;
            color: currentColor !important;
            -webkit-text-fill-color: currentColor !important;
            font-size: 0.40rem !important;
            line-height: 1 !important;
            font-weight: 850 !important;
            letter-spacing: 0 !important;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .tc-input-sidebar-mini-action:hover:not(:disabled) {
            transform: translateY(-1px);
            background: #24272c;
            border-color: rgba(255,255,255,0.62);
            box-shadow: 0 8px 18px rgba(55,0,9,0.30);
          }
          .tc-input-sidebar-mini-action:active:not(:disabled) {
            transform: scale(0.95);
          }
          .tc-input-sidebar-mini-action.is-active {
            background: #ffffff;
            color: #c90012;
            border-color: #ffffff;
            box-shadow: 0 0 0 2px rgba(255,255,255,0.28), 0 7px 18px rgba(55,0,9,0.28);
          }
          .tc-input-sidebar-mini-action.is-mapped .tc-input-sidebar-mini-status {
            position: absolute;
            top: 3px;
            right: 3px;
            display: block;
            width: 5px;
            height: 5px;
            border: 1px solid rgba(255,255,255,0.82);
            border-radius: 999px;
            background: #27a66b;
            box-shadow: 0 1px 4px rgba(0,0,0,0.22);
          }
          .tc-input-sidebar-mini-action.is-active.is-mapped .tc-input-sidebar-mini-status {
            border-color: #ffffff;
          }
          .tc-input-sidebar-mini-action:disabled {
            cursor: default;
            opacity: 0.34;
            box-shadow: none;
          }
          .tc-input-sidebar-mini-divider {
            width: 22px;
            height: 1px;
            margin: 0.08rem 0 0.02rem;
            background: rgba(255,255,255,0.30);
            flex: 0 0 auto;
          }
          .tc-input-sidebar-mini-action.is-sort {
            background: rgba(255,255,255,0.13);
            border-color: rgba(255,255,255,0.42);
          }
          .tc-input-sidebar-mini-action.is-sort:hover:not(:disabled) {
            background: rgba(255,255,255,0.22);
          }

          @media (max-height: 760px) {
            #tc-input-sidebar-mini-actions {
              top: 5.95rem;
              gap: 0.34rem;
            }
            .tc-input-sidebar-mini-action {
              height: 34px;
            }
            .tc-input-sidebar-mini-action svg {
              width: 13px;
              height: 13px;
            }
          }


          /* r85: compact Inputs must be a true alternate rail, not an overlay.
             Hide the complete normal Streamlit content tree while collapsed;
             only the injected compact controls remain visible. */
          html.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar > :not(#tc-input-sidebar-toggle):not(#tc-input-sidebar-mini):not(#tc-input-sidebar-mini-actions),
          body.tc-input-sidebar-is-collapsed .st-key-l1_input_sidebar > :not(#tc-input-sidebar-toggle):not(#tc-input-sidebar-mini):not(#tc-input-sidebar-mini-actions) {
            display: none !important;
            visibility: hidden !important;
            opacity: 0 !important;
            pointer-events: none !important;
          }
          html.tc-input-sidebar-is-collapsed #tc-input-sidebar-toggle,
          html.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini,
          html.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini-actions,
          body.tc-input-sidebar-is-collapsed #tc-input-sidebar-toggle,
          body.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini,
          body.tc-input-sidebar-is-collapsed #tc-input-sidebar-mini-actions {
            visibility: visible !important;
            opacity: 1 !important;
            pointer-events: auto !important;
          }

          /* Compact company-count status: one quiet line, no card/button shell. */
          .st-key-l1_input_sidebar .st-key-sidebar_company_count_info {
            width: 100% !important;
            margin: 0.08rem 0 0.32rem !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .st-key-sidebar_company_count_info > [data-testid="stVerticalBlock"],
          .st-key-l1_input_sidebar .st-key-sidebar_company_count_info > div > [data-testid="stVerticalBlock"] {
            gap: 0 !important;
            padding: 0 !important;
          }
          .st-key-l1_input_sidebar .tc-company-count-inline {
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            gap: 0.38rem !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0.12rem 0 !important;
            color: rgba(255,255,255,0.88) !important;
            -webkit-text-fill-color: rgba(255,255,255,0.88) !important;
            font-size: 0.68rem !important;
            line-height: 1.2 !important;
            font-weight: 700 !important;
            letter-spacing: 0.01em !important;
            text-align: center !important;
            white-space: nowrap !important;
          }
          .st-key-l1_input_sidebar .tc-company-count-inline strong {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            font-size: 0.72rem !important;
            font-weight: 900 !important;
          }
          .st-key-l1_input_sidebar .tc-company-count-dot {
            display: inline-block !important;
            width: 0.38rem !important;
            height: 0.38rem !important;
            flex: 0 0 0.38rem !important;
            border-radius: 999px !important;
            background: #36b879 !important;
            box-shadow: 0 0 0 2px rgba(255,255,255,0.16) !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r87: one masthead system across Slide Maker, Inputs, and Company Sort.
             The companion-rail titles get enough left clearance from the
             preceding rail toggle, and all three headers share the same title
             scale, header height, divider rhythm, and bottom spacing. */
          :root {
            --tc-unified-rail-title-size: 1.24rem;
            --tc-unified-rail-title-weight: 860;
            --tc-unified-rail-title-line: 1.08;
            --tc-unified-rail-header-height: 3.72rem;
            --tc-unified-rail-header-gap: 0.96rem;
            --tc-unified-rail-title-clearance: 1.08rem;
            --tc-unified-rail-divider: rgba(255,255,255,0.26);
          }

          /* Same title typography everywhere. */
          [data-testid="stSidebar"] .tc-panel-brand-title,
          .st-key-l1_input_sidebar .tc-input-sidebar-header,
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            font-size: var(--tc-unified-rail-title-size) !important;
            line-height: var(--tc-unified-rail-title-line) !important;
            font-weight: var(--tc-unified-rail-title-weight) !important;
            letter-spacing: -0.022em !important;
          }

          /* Keep Slide Maker's clean masthead/divider as the reference. */
          [data-testid="stSidebar"] .tc-panel-brand {
            box-sizing: border-box !important;
            min-height: var(--tc-unified-rail-header-height) !important;
            margin: 0 0 var(--tc-unified-rail-header-gap) !important;
            padding-bottom: 0.82rem !important;
            border-bottom: 1px solid var(--tc-primary-rail-line, #292d33) !important;
          }

          /* Inputs + Company Sort: identical header geometry, with left room
             for the collapse button on the rail immediately to their left. */
          .st-key-l1_input_sidebar .tc-input-sidebar-header,
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            position: relative !important;
            box-sizing: border-box !important;
            display: flex !important;
            align-items: center !important;
            min-height: var(--tc-unified-rail-header-height) !important;
            margin: 0 0 var(--tc-unified-rail-header-gap) !important;
            padding: 0 2.05rem 0.82rem var(--tc-unified-rail-title-clearance) !important;
            border-bottom: 0 !important;
          }

          /* Draw the companion-rail divider at the same visual inset as the
             Slide Maker line, rather than letting it run edge-to-edge. */
          .st-key-l1_input_sidebar .tc-input-sidebar-header::after,
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header::after {
            content: "" !important;
            position: absolute !important;
            left: 0.45rem !important;
            right: 0.45rem !important;
            bottom: 0 !important;
            height: 1px !important;
            background: var(--tc-unified-rail-divider) !important;
            pointer-events: none !important;
          }

          /* Company Sort home is a CSS grid. Its row-gap already supplies the
             post-header spacing, so do not double-count the common margin. */
          .st-key-l1_segment_sidebar .tc-sort-home-topstack .tc-segment-sidebar-header {
            margin-bottom: 0 !important;
          }

          /* Keep short-height layouts symmetric too. */
          @media (max-height: 820px) {
            [data-testid="stSidebar"] .tc-panel-brand,
            .st-key-l1_input_sidebar .tc-input-sidebar-header,
            .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
              min-height: 3.42rem !important;
              padding-bottom: 0.70rem !important;
            }
            [data-testid="stSidebar"] .tc-panel-brand-title,
            .st-key-l1_input_sidebar .tc-input-sidebar-header,
            .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
              font-size: 1.17rem !important;
            }
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


    st.markdown(
        """
        <style data-tc-ui-style="true">
          /* r88: final visual normalization. Keep this last so older compatibility
             blocks cannot re-introduce header overlap or stretched workflow chips. */
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] div[data-baseweb="select"],
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] div[role="combobox"] {
            box-sizing: border-box !important;
            width: 100% !important;
            min-height: 3rem !important;
            padding-right: 3.6rem !important;
            background: #111214 !important;
            background-color: #111214 !important;
            background-image: none !important;
            border: 1px solid #34383f !important;
            border-radius: 0.55rem !important;
            box-shadow: none !important;
          }
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] [data-baseweb="tag"] {
            box-sizing: border-box !important;
            display: inline-flex !important;
            flex: 0 0 auto !important;
            width: auto !important;
            min-width: 0 !important;
            max-width: 11rem !important;
            height: 2.18rem !important;
            margin: 0.14rem 0.28rem 0.14rem 0 !important;
            background: #cc0000 !important;
            background-color: #cc0000 !important;
            border: 1px solid #cc0000 !important;
            border-radius: 0.46rem !important;
          }
          [data-testid="stSidebar"] .st-key-field_workflow_mode [data-testid="stMultiSelect"] [data-baseweb="tag"] * {
            color: #fff !important;
            -webkit-text-fill-color: #fff !important;
            font-weight: 800 !important;
          }

          :root {
            --tc-final-rail-title-size: 1.22rem;
            --tc-final-rail-header-height: 3.58rem;
            --tc-final-rail-title-clearance: 2.05rem;
          }
          [data-testid="stSidebar"] .tc-panel-brand-title,
          .st-key-l1_input_sidebar .tc-input-sidebar-header,
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            font-size: var(--tc-final-rail-title-size) !important;
            line-height: 1.08 !important;
            font-weight: 860 !important;
            letter-spacing: -0.022em !important;
          }
          .st-key-l1_input_sidebar .tc-input-sidebar-header,
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header {
            box-sizing: border-box !important;
            min-height: var(--tc-final-rail-header-height) !important;
            margin: 0 0 0.92rem !important;
            padding: 0 var(--tc-final-rail-title-clearance) 0.78rem var(--tc-final-rail-title-clearance) !important;
            display: flex !important;
            align-items: center !important;
            position: relative !important;
            border-bottom: 0 !important;
          }
          /* r89 divider normalization: the Slide Maker divider is already inset by
             the primary panel's 1.38rem horizontal gutter. Mirror that *visible*
             inset in the companion rails and keep every divider on the exact
             bottom edge of the same-height masthead. */
          .st-key-l1_input_sidebar .tc-input-sidebar-header::after,
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header::after {
            content: "" !important;
            position: absolute !important;
            bottom: 0 !important;
            height: 1px !important;
            background: #2a2e34 !important;
            pointer-events: none !important;
          }
          /* Inputs has no horizontal rail padding, so use the full Slide Maker
             gutter directly. */
          .st-key-l1_input_sidebar .tc-input-sidebar-header::after {
            left: 1.38rem !important;
            right: 1.38rem !important;
          }
          /* Company Sort already lives inside a 0.90rem horizontal rail gutter;
             add the remaining 0.48rem so its visible divider starts/ends at the
             same 1.38rem distance from the rail edge. */
          .st-key-l1_segment_sidebar .tc-segment-sidebar-header::after {
            left: 0.48rem !important;
            right: 0.48rem !important;
          }
          [data-testid="stSidebar"] .tc-panel-brand {
            min-height: var(--tc-final-rail-header-height) !important;
            border-bottom-color: #2a2e34 !important;
            border-bottom-width: 1px !important;
          }
        </style>
        """,
        unsafe_allow_html=True,
    )


def initialize_canvas_state() -> None:
    defaults = {
        "last_output": None,
        "last_logs": [],
        "last_command": "",
        "last_state": "ready",
        "canvas_stage": "review",
        "workbook_review_finalized": False,
        "preview_source_role": "aura",
        "preview_input_navigation_token": "",
        "selected_aura_sheet_override": "",
        "resolved_aura_sheet": "",
        "selected_taxonomy_sheet_override": "",
        "resolved_taxonomy_sheet": "",
        "resolved_geocost_sheet": "",
        "taxonomy_sheet_profile": {},
        "slide_input_mappings": {},
        "active_input_by_section": {},
        "l1_mapping_validation": {},
        "l1_auto_detection": {},
        "l1_master_label_profile": {},
        "l1_function_lookup_profile": {},
        "auto_mapping_rerun_signature": "",
        "segment_profile_rerun_signature": "",
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def _render_console_header() -> None:
    st.markdown(
        """
        <div class="tc-main-header">
          <div class="tc-main-title">Build console</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_layout_error(
    console: ConsoleTargets,
    layout_problems: list[str],
) -> None:
    console.progress.markdown(
        render_progress_html(0, "error", "Project layout is not ready", 0),
        unsafe_allow_html=True,
    )
    error_logs = ["PROJECT LAYOUT ERROR", *layout_problems]
    console.log.markdown(render_log_html(error_logs), unsafe_allow_html=True)
    render_result_panel(
        console.result,
        None,
        error_logs,
        "Fix the project layout shown in the build log before starting.",
    )


def _inject_finalize_hover_tooltip(help_text: str) -> None:
    """Update Finalize help in the persistent document-level tooltip controller."""
    payload = json.dumps(str(help_text or "").strip(), ensure_ascii=False)
    markup = r"""
    <!doctype html>
    <html><head><style>html,body{margin:0;padding:0;width:0;height:0;overflow:hidden}</style></head>
    <body><script>
    (() => {
      const win = window.parent;
      const doc = win.document;
      const TEXT = __TC_FINALIZE_HELP__;
      let state = win.__tcPersistentDiagnosticTooltip;

      // The Inputs rail normally creates the controller first.  Keep a small
      // bootstrap here as a safety net so Finalize still gets a tooltip if the
      // render order changes in a future Streamlit release.
      if (!state) {
        state = {
          inputMap: {}, finalizeText: '', timer: null, hoverTarget: null, hoverKind: '',
          bubbleId: 'tc-persistent-diagnostic-tooltip'
        };
        win.__tcPersistentDiagnosticTooltip = state;

        function ensureBubble() {
          let b = doc.getElementById(state.bubbleId);
          if (b) return b;
          b = doc.createElement('div');
          b.id = state.bubbleId;
          b.setAttribute('role', 'tooltip');
          Object.assign(b.style, {
            position:'fixed', zIndex:'2147483647', maxWidth:'440px', minWidth:'240px',
            padding:'13px 15px', borderRadius:'12px', background:'#15171a', color:'#fff',
            border:'1px solid rgba(255,255,255,.16)', boxShadow:'0 18px 42px rgba(0,0,0,.40)',
            fontFamily:'Arial, Helvetica, sans-serif', fontSize:'12px', fontWeight:'650', lineHeight:'1.48',
            whiteSpace:'pre-line', overflowWrap:'anywhere', pointerEvents:'none', opacity:'0', visibility:'hidden',
            transform:'translateY(4px)', transition:'opacity 120ms ease, transform 120ms ease, visibility 120ms ease'
          });
          const a=doc.createElement('div'); Object.assign(a.style,{position:'absolute',left:'12px',right:'12px',top:'0',height:'3px',background:'#cc0000',borderRadius:'0 0 3px 3px'});
          const body=doc.createElement('div'); body.className='tc-persistent-diagnostic-tooltip-body';
          const ar=doc.createElement('div'); ar.className='tc-persistent-diagnostic-tooltip-arrow'; Object.assign(ar.style,{position:'absolute',width:'10px',height:'10px',background:'#15171a',transform:'rotate(45deg)',border:'1px solid rgba(255,255,255,.10)'});
          b.append(a,body,ar); doc.body.appendChild(b); return b;
        }
        state.hide=function(){if(state.timer){win.clearTimeout(state.timer);state.timer=null;}state.hoverTarget=null;state.hoverKind='';const b=doc.getElementById(state.bubbleId);if(b){b.style.opacity='0';b.style.visibility='hidden';b.style.transform='translateY(4px)';}};
        state.resolve=function(node){if(!node||!node.closest)return null;const card=node.closest('[class*="st-key-sidebar_l1_input_"]');if(card&&!card.classList.contains('st-key-sidebar_l1_input_list')){const cls=Array.from(card.classList).find(n=>n.startsWith('st-key-sidebar_l1_input_')&&n!=='st-key-sidebar_l1_input_list');const key=cls?cls.slice('st-key-sidebar_l1_input_'.length):'';const text=String((state.inputMap||{})[key]||'').trim();if(text)return{target:card,text,kind:'input'};}const f=node.closest('.st-key-workbook_review_action');if(f){const text=String(state.finalizeText||'').trim();if(text)return{target:f,text,kind:'finalize'};}return null;};
        state.show=function(item){if(!item||!item.target||!item.target.isConnected)return;const b=ensureBubble(),body=b.querySelector('.tc-persistent-diagnostic-tooltip-body');if(body)body.textContent=item.text;b.style.display='block';b.style.visibility='hidden';b.style.opacity='0';const r=item.target.getBoundingClientRect(),br=b.getBoundingClientRect(),gap=12,pad=10;let left,top,side;if(item.kind==='finalize'){left=r.right-br.width;top=r.bottom+gap;side='top';if(top+br.height>win.innerHeight-pad){top=r.top-br.height-gap;side='bottom';}}else{left=r.right+gap;top=r.top+(r.height-br.height)/2;side='left';if(left+br.width>win.innerWidth-pad){left=r.left-br.width-gap;side='right';}}left=Math.max(pad,Math.min(left,win.innerWidth-br.width-pad));top=Math.max(pad,Math.min(top,win.innerHeight-br.height-pad));b.style.left=Math.round(left)+'px';b.style.top=Math.round(top)+'px';const ar=b.querySelector('.tc-persistent-diagnostic-tooltip-arrow');if(ar){ar.style.left='';ar.style.right='';ar.style.top='';ar.style.bottom='';if(side==='left'){ar.style.left='-6px';ar.style.top=Math.max(14,Math.min(br.height-24,r.top+r.height/2-top-5))+'px';}else if(side==='right'){ar.style.right='-6px';ar.style.top=Math.max(14,Math.min(br.height-24,r.top+r.height/2-top-5))+'px';}else if(side==='top'){ar.style.top='-6px';ar.style.right='24px';}else{ar.style.bottom='-6px';ar.style.right='24px';}}b.style.visibility='visible';b.style.opacity='1';b.style.transform='translateY(0)';};
        if(!state.listenersInstalled){state.listenersInstalled=true;const DELAY=2000;doc.addEventListener('mouseover',(e)=>{const item=state.resolve(e.target);if(!item)return;if(state.hoverTarget===item.target&&state.hoverKind===item.kind)return;state.hide();state.hoverTarget=item.target;state.hoverKind=item.kind;state.timer=win.setTimeout(()=>{state.timer=null;const current=state.resolve(item.target);if(current&&current.target===item.target)state.show(current);},DELAY);},true);doc.addEventListener('mouseout',(e)=>{if(!state.hoverTarget)return;const rel=e.relatedTarget;if(rel&&state.hoverTarget.contains&&state.hoverTarget.contains(rel))return;const leaving=e.target&&e.target.closest?e.target.closest('[class*="st-key-sidebar_l1_input_"], .st-key-workbook_review_action'):null;if(leaving===state.hoverTarget||!rel)state.hide();},true);doc.addEventListener('mousedown',()=>state.hide(),true);win.addEventListener('scroll',()=>state.hide(),true);win.addEventListener('resize',()=>state.hide(),true);}
      }
      state.finalizeText = String(TEXT || '').trim();
    })();
    </script></body></html>
    """.replace('__TC_FINALIZE_HELP__', payload)
    components.html(markup, height=0, scrolling=False)


def render_canvas(
    paths: AppPaths,
    controls: ControlPanelState,
    layout_problems: list[str],
    load_progress: WorkbookLoadProgress | None = None,
) -> None:
    """Render the guided workbook-review then build-console workflow."""
    if controls.workbook_path is None:
        st.session_state["canvas_stage"] = "review"
        st.session_state["workbook_review_finalized"] = False

    stage = str(st.session_state.get("canvas_stage") or "review").lower()
    review_complete = bool(
        st.session_state.get("workbook_review_finalized", False)
    )

    if stage != "console" or not review_complete:
        render_workbook_preview(
            controls.preview_workbook_path,
            controls.workbook_display_name,
            preview_metadata=controls.preview_metadata,
            preview_error=controls.preview_error,
            workflow_mode=controls.workflow_mode,
            selected_sections=controls.selected_sections,
            input_section=controls.input_section,
            load_progress=load_progress,
        )
        return

    if load_progress is not None and load_progress.active:
        load_progress.clear(minimum_visible_seconds=0.0)

    _render_console_header()
    _render_source_selectors(
        controls.preview_workbook_path or controls.workbook_path,
        controls.workflow_mode,
        selected_sections=controls.selected_sections,
        preview_enabled=False,
    )
    progress_placeholder = st.empty()
    with st.expander("Detailed build log", expanded=False):
        log_placeholder = st.empty()
    result_placeholder = st.empty()
    console = ConsoleTargets(
        progress=progress_placeholder,
        log=log_placeholder,
        result=result_placeholder,
    )

    if layout_problems:
        _render_layout_error(console, layout_problems)
        return

    previous_logs = list(st.session_state.get("last_logs") or [])
    previous_output_raw = st.session_state.get("last_output")
    previous_output = Path(previous_output_raw) if previous_output_raw else None
    previous_state = str(st.session_state.get("last_state") or "ready")
    previous_tracker = st.session_state.get("last_build_tracker")

    if isinstance(previous_tracker, dict) and previous_tracker.get("families"):
        tracker_state = (
            "success"
            if previous_state == "success"
            else "error"
            if previous_state == "error"
            else "running"
        )
        tracker_message = (
            "Deck generated successfully"
            if tracker_state == "success"
            else "Previous build did not complete"
            if tracker_state == "error"
            else "Build status"
        )
        progress_placeholder.markdown(
            render_build_tracker_html(
                previous_tracker,
                0,
                final_state=tracker_state,
                final_message=tracker_message,
            ),
            unsafe_allow_html=True,
        )
    elif previous_state == "success" and previous_output and previous_output.is_file():
        progress_placeholder.markdown(
            render_progress_html(100, "success", "Deck generated successfully", 0),
            unsafe_allow_html=True,
        )
    elif previous_state == "error":
        progress_placeholder.markdown(
            render_progress_html(0, "error", "Previous build did not complete", 0),
            unsafe_allow_html=True,
        )
    else:
        ready_stage = (
            "Ready to generate the selected deck"
            if controls.workbook_path is not None
            else "Waiting for a source workbook"
        )
        progress_placeholder.markdown(
            render_progress_html(0, "ready", ready_stage, 0),
            unsafe_allow_html=True,
        )

    log_placeholder.markdown(render_log_html(previous_logs), unsafe_allow_html=True)
    if previous_output and previous_output.is_file():
        render_result_panel(result_placeholder, previous_output, previous_logs)

    if not controls.generate_clicked:
        return

    try:
        options = controls.build_options()
        run_backend(options, console, paths)
    except Exception as exc:
        st.session_state["last_state"] = "error"
        st.session_state["last_build_tracker"] = {}
        message = f"UI error: {exc}"
        progress_placeholder.markdown(
            render_progress_html(0, "error", message, 0),
            unsafe_allow_html=True,
        )
        error_logs = previous_logs + [message]
        st.session_state["last_logs"] = error_logs
        log_placeholder.markdown(render_log_html(error_logs), unsafe_allow_html=True)
        render_result_panel(result_placeholder, None, error_logs, message)


__all__ = [
    "apply_theme",
    "initialize_canvas_state",
    "inspect_uploaded_workbook",
    "render_canvas",
]
