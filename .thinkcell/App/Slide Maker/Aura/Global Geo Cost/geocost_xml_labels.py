from __future__ import annotations

import logging
import posixpath
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from geocost_common import (
    AUTOLABEL_PREFIX,
    GeoCostData,
    LabelTarget,
    canonical_country_key,
    country_key_set as _country_key_set,
    font_hex_for_bucket as _font_hex_for_bucket,
)

PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _qn(namespace: str, local_name: str) -> str:
    return f"{{{namespace}}}{local_name}"


def _xml_parser():
    for prefix, namespace in (
        ("p", PML_NS),
        ("a", DML_NS),
        ("c", CHART_NS),
        ("r", REL_NS),
    ):
        ET.register_namespace(prefix, namespace)
    return ET


def _shape_text_xml(shape, etree) -> str:
    return "".join(
        (item.text or "")
        for item in shape.findall(f".//{_qn(DML_NS, 't')}")
    ).strip()


def _shape_bounds_xml(shape, etree) -> tuple[float, float, float, float] | None:
    xfrm = shape.find(f".//{_qn(PML_NS, 'xfrm')}")
    if xfrm is None:
        xfrm = shape.find(f".//{_qn(DML_NS, 'xfrm')}")
    if xfrm is None:
        return None
    off = xfrm.find(_qn(DML_NS, "off"))
    ext = xfrm.find(_qn(DML_NS, "ext"))
    if off is None or ext is None:
        return None
    try:
        x = float(off.get("x", "0"))
        y = float(off.get("y", "0"))
        cx = abs(float(ext.get("cx", "0")))
        cy = abs(float(ext.get("cy", "0")))
    except Exception:
        return None
    return x, y, cx, cy


def _center_within_bounds(
    item_bounds: tuple[float, float, float, float] | None,
    target_bounds: tuple[float, float, float, float],
    margin: float,
) -> bool:
    if item_bounds is None:
        return False
    x, y, cx, cy = item_bounds
    tx, ty, tcx, tcy = target_bounds
    center_x = x + cx / 2.0
    center_y = y + cy / 2.0
    return (
        tx - margin <= center_x <= tx + tcx + margin
        and ty - margin <= center_y <= ty + tcy + margin
    )


def _select_chart_frame(slide_root, etree):
    frames = [
        frame
        for frame in slide_root.iter(_qn(PML_NS, "graphicFrame"))
        if frame.find(f".//{_qn(CHART_NS, 'chart')}") is not None
    ]
    if not frames:
        raise RuntimeError("No PowerPoint chart frame was found on a GeoCost slide.")
    def area(frame) -> float:
        bounds = _shape_bounds_xml(frame, etree)
        return 0.0 if bounds is None else bounds[2] * bounds[3]
    return max(frames, key=area)


def _ordered_slide_part_names(source, etree) -> list[str]:
    presentation_name = "ppt/presentation.xml"
    rels_name = "ppt/_rels/presentation.xml.rels"
    if presentation_name not in source.namelist() or rels_name not in source.namelist():
        raise RuntimeError("The PowerPoint package is missing presentation relationships.")
    presentation_root = etree.fromstring(source.read(presentation_name))
    rels_root = etree.fromstring(source.read(rels_name))
    target_by_id = {
        item.get("Id"): item.get("Target", "")
        for item in rels_root.findall(_qn(PKG_REL_NS, "Relationship"))
    }
    ordered: list[str] = []
    for slide_id in presentation_root.findall(
        f".//{_qn(PML_NS, 'sldId')}"
    ):
        rel_id = slide_id.get(_qn(REL_NS, "id"))
        target = target_by_id.get(rel_id, "")
        if not target:
            continue
        ordered.append(posixpath.normpath(posixpath.join("ppt", target)))
    if not ordered:
        raise RuntimeError("No slides were found in the PowerPoint presentation order.")
    return ordered


def _chart_relationship_target(
    slide_name: str,
    slide_root,
    rels_root,
    etree,
) -> tuple[str, tuple[float, float, float, float]]:
    chart_frame = _select_chart_frame(slide_root, etree)
    bounds = _shape_bounds_xml(chart_frame, etree)
    if bounds is None:
        raise RuntimeError(f"Could not read chart bounds from {slide_name}.")
    chart_node = chart_frame.find(f".//{_qn(CHART_NS, 'chart')}")
    if chart_node is None:
        raise RuntimeError(f"Could not find chart relationship in {slide_name}.")
    rel_id = chart_node.get(_qn(REL_NS, "id"))
    if not rel_id:
        raise RuntimeError(f"Chart relationship ID is missing in {slide_name}.")
    relationship = None
    for item in rels_root.findall(_qn(PKG_REL_NS, "Relationship")):
        if item.get("Id") == rel_id:
            relationship = item
            break
    if relationship is None:
        raise RuntimeError(f"Chart relationship {rel_id!r} was not found for {slide_name}.")
    target = relationship.get("Target", "")
    if not target:
        raise RuntimeError(f"Chart relationship target is blank for {slide_name}.")
    chart_name = posixpath.normpath(posixpath.join(posixpath.dirname(slide_name), target))
    return chart_name, bounds


def _remove_stale_country_shapes_xml(
    slide_root,
    chart_bounds: tuple[float, float, float, float],
    template_country_names: list[str],
    etree,
) -> int:
    """Remove the old think-cell-rendered country labels and their leader lines.
    This edits only ordinary slide shapes. It does not parse or modify the
    undocumented think-cell OLE data model.
    """
    keys = _country_key_set(template_country_names)
    sp_tree = slide_root.find(f".//{_qn(PML_NS, 'spTree')}")
    if sp_tree is None:
        return 0
    margin = 320040.0
    removed = 0
    for child in list(sp_tree):
        local_name = child.tag.rsplit("}", 1)[-1]
        if local_name == "sp":
            name_node = child.find(f".//{_qn(PML_NS, 'cNvPr')}")
            shape_name = "" if name_node is None else str(name_node.get("name", ""))
            text = _shape_text_xml(child, etree)
            is_generated = shape_name.startswith(AUTOLABEL_PREFIX)
            is_country = canonical_country_key(text) in keys
            if (is_generated or is_country) and _center_within_bounds(
                _shape_bounds_xml(child, etree),
                chart_bounds,
                margin,
            ):
                sp_tree.remove(child)
                removed += 1
        elif local_name == "cxnSp" and _center_within_bounds(
            _shape_bounds_xml(child, etree),
            chart_bounds,
            margin,
        ):
            sp_tree.remove(child)
            removed += 1
    return removed


def _append_label_text(
    parent,
    text: str,
    font_rgb: str,
    etree,
) -> None:
    tx = etree.SubElement(parent, _qn(CHART_NS, "tx"))
    rich = etree.SubElement(tx, _qn(CHART_NS, "rich"))
    etree.SubElement(rich, _qn(DML_NS, "bodyPr"))
    etree.SubElement(rich, _qn(DML_NS, "lstStyle"))
    paragraph = etree.SubElement(rich, _qn(DML_NS, "p"))
    etree.SubElement(paragraph, _qn(DML_NS, "pPr"), algn="ctr")
    run = etree.SubElement(paragraph, _qn(DML_NS, "r"))
    run_props = etree.SubElement(
        run,
        _qn(DML_NS, "rPr"),
        lang="en-US",
        sz="1000",
        kern="1200",
    )
    fill = etree.SubElement(run_props, _qn(DML_NS, "solidFill"))
    etree.SubElement(fill, _qn(DML_NS, "srgbClr"), val=font_rgb)
    etree.SubElement(run_props, _qn(DML_NS, "latin"), typeface="+mn-lt")
    etree.SubElement(run, _qn(DML_NS, "t")).text = text
    end_props = etree.SubElement(
        paragraph,
        _qn(DML_NS, "endParaRPr"),
        lang="en-US",
        sz="1000",
    )
    end_fill = etree.SubElement(end_props, _qn(DML_NS, "solidFill"))
    etree.SubElement(end_fill, _qn(DML_NS, "srgbClr"), val=font_rgb)
    etree.SubElement(end_props, _qn(DML_NS, "latin"), typeface="+mn-lt")


def _append_label_visibility_flags(parent, etree) -> None:
    for tag in (
        "showLegendKey",
        "showVal",
        "showCatName",
        "showSerName",
        "showPercent",
        "showBubbleSize",
    ):
        etree.SubElement(parent, _qn(CHART_NS, tag), val="0")


def _build_series_data_labels(targets: list[LabelTarget], etree):
    d_labels = etree.Element(_qn(CHART_NS, "dLbls"))
    for target in sorted(targets, key=lambda item: item.company_index):
        label = etree.SubElement(d_labels, _qn(CHART_NS, "dLbl"))
        etree.SubElement(label, _qn(CHART_NS, "idx"), val=str(target.company_index))
        _append_label_text(
            label,
            target.display_label,
            _font_hex_for_bucket(target.segment_bucket),
            etree,
        )
        shape_props = etree.SubElement(label, _qn(CHART_NS, "spPr"))
        etree.SubElement(shape_props, _qn(DML_NS, "noFill"))
        line = etree.SubElement(shape_props, _qn(DML_NS, "ln"))
        etree.SubElement(line, _qn(DML_NS, "noFill"))
        etree.SubElement(label, _qn(CHART_NS, "dLblPos"), val="ctr")
        _append_label_visibility_flags(label, etree)
    etree.SubElement(d_labels, _qn(CHART_NS, "dLblPos"), val="ctr")
    _append_label_visibility_flags(d_labels, etree)
    etree.SubElement(d_labels, _qn(CHART_NS, "showLeaderLines"), val="0")
    return d_labels


def _patch_chart_xml_with_targets(
    chart_bytes: bytes,
    targets: list[LabelTarget],
    etree,
    country_count: int,
) -> tuple[bytes, int]:
    chart_root = etree.fromstring(chart_bytes)
    bar_charts = list(chart_root.iter(_qn(CHART_NS, "barChart")))
    if not bar_charts:
        raise RuntimeError("The GeoCost PowerPoint chart is not a bar/column chart.")
    bar_chart = max(
        bar_charts,
        key=lambda item: len(item.findall(_qn(CHART_NS, "ser"))),
    )
    series = bar_chart.findall(_qn(CHART_NS, "ser"))
    if len(series) < country_count:
        raise RuntimeError(
            f"The visible PowerPoint chart exposes {len(series)} series; "
            f"expected at least {country_count}."
        )
    targets_by_row: dict[int, list[LabelTarget]] = {}
    for target in targets:
        if not 0 <= target.row_index < country_count:
            raise RuntimeError(f"Country row index is outside the chart: {target.row_index}.")
        targets_by_row.setdefault(target.row_index, []).append(target)
    for row_index, series_node in enumerate(series[:country_count]):
        for existing in list(series_node.findall(_qn(CHART_NS, "dLbls"))):
            series_node.remove(existing)
        row_targets = targets_by_row.get(row_index, [])
        if not row_targets:
            continue
        d_labels = _build_series_data_labels(row_targets, etree)
        children = list(series_node)
        insertion_index = len(children)
        for candidate_name in ("cat", "val", "shape", "extLst"):
            candidate = series_node.find(_qn(CHART_NS, candidate_name))
            if candidate is not None:
                insertion_index = min(insertion_index, children.index(candidate))
        series_node.insert(insertion_index, d_labels)
    output = etree.tostring(
        chart_root,
        xml_declaration=True,
        encoding="utf-8",
    )
    return output, len(targets)


def _verify_chart_xml_targets(
    chart_bytes: bytes,
    targets: list[LabelTarget],
    etree,
    country_count: int,
) -> None:
    root = etree.fromstring(chart_bytes)
    bar_charts = list(root.iter(_qn(CHART_NS, "barChart")))
    if not bar_charts:
        raise RuntimeError("Could not verify GeoCost chart labels: bar chart missing.")
    bar_chart = max(
        bar_charts,
        key=lambda item: len(item.findall(_qn(CHART_NS, "ser"))),
    )
    series = bar_chart.findall(_qn(CHART_NS, "ser"))
    actual: dict[tuple[int, int], str] = {}
    for row_index, series_node in enumerate(series[:country_count]):
        d_labels = series_node.find(_qn(CHART_NS, "dLbls"))
        if d_labels is None:
            continue
        for label in d_labels.findall(_qn(CHART_NS, "dLbl")):
            idx = label.find(_qn(CHART_NS, "idx"))
            if idx is None:
                continue
            company_index = int(idx.get("val", "-1"))
            text = "".join(
                (item.text or "")
                for item in label.iter(_qn(DML_NS, "t"))
            )
            position = label.find(_qn(CHART_NS, "dLblPos"))
            if position is None or position.get("val") != "ctr":
                raise RuntimeError(
                    f"Label at row {row_index}, company {company_index} is not centered."
                )
            actual[(company_index, row_index)] = text
    expected = {
        (target.company_index, target.row_index): target.display_label
        for target in targets
    }
    if actual != expected:
        raise RuntimeError(
            "Standard chart-label verification failed. "
            f"Expected {sorted(expected.items())}, got {sorted(actual.items())}."
        )


def patch_standard_chart_labels_in_pptx(
    pptx_path: Path,
    data_by_slide: dict[int, GeoCostData],
    logger: logging.Logger,
) -> None:
    """Patch only standard PowerPoint slide/chart XML.
    The think-cell chart data is populated beforehand through AddRangeData. This
    pass removes the stale rendered country-label shapes and adds one centered
    PowerPoint chart-point label for every qualifying country × company segment.
    It never reads or modifies the undocumented think-cell OLE/root structure.
    """
    etree = _xml_parser()
    replacement_parts: dict[str, bytes] = {}
    verification_parts: list[tuple[str, list[LabelTarget], int]] = []
    total_removed = 0
    total_labels = 0
    chart_owner: dict[str, int] = {}
    with zipfile.ZipFile(pptx_path, "r") as source:
        source_names = set(source.namelist())
        ordered_slide_names = _ordered_slide_part_names(source, etree)
        for slide_index, data in sorted(data_by_slide.items()):
            if slide_index < 1 or slide_index > len(ordered_slide_names):
                raise RuntimeError(
                    f"Requested slide {slide_index}, but {pptx_path.name} has "
                    f"{len(ordered_slide_names)} slide(s)."
                )
            slide_name = ordered_slide_names[slide_index - 1]
            rels_name = posixpath.join(
                posixpath.dirname(slide_name),
                "_rels",
                posixpath.basename(slide_name) + ".rels",
            )
            if slide_name not in source_names or rels_name not in source_names:
                raise RuntimeError(f"Slide {slide_index} or its relationships are missing from {pptx_path.name}.")
            slide_root = etree.fromstring(source.read(slide_name))
            rels_root = etree.fromstring(source.read(rels_name))
            chart_name, chart_bounds = _chart_relationship_target(
                slide_name,
                slide_root,
                rels_root,
                etree,
            )
            if chart_name not in source_names:
                raise RuntimeError(f"Chart part {chart_name!r} is missing from {pptx_path.name}.")
            previous_owner = chart_owner.get(chart_name)
            if previous_owner is not None and previous_owner != slide_index:
                raise RuntimeError(
                    f"Chart part {chart_name} is shared by slides {previous_owner} and {slide_index}; "
                    "per-slide label targets would conflict."
                )
            chart_owner[chart_name] = slide_index
            removed = _remove_stale_country_shapes_xml(
                slide_root,
                chart_bounds,
                data.template_country_names,
                etree,
            )
            total_removed += removed
            replacement_parts[slide_name] = etree.tostring(
                slide_root,
                xml_declaration=True,
                encoding="utf-8",
            )
            original_chart = replacement_parts.get(chart_name, source.read(chart_name))
            patched_chart, label_count = _patch_chart_xml_with_targets(
                original_chart,
                data.label_targets,
                etree,
                data.template_layout.country_count,
            )
            replacement_parts[chart_name] = patched_chart
            verification_parts.append((chart_name, data.label_targets, data.template_layout.country_count))
            total_labels += label_count
        temp_path = pptx_path.with_suffix(pptx_path.suffix + ".standardlabels")
        with zipfile.ZipFile(temp_path, "w") as target:
            for info in source.infolist():
                target.writestr(info, replacement_parts.get(info.filename, source.read(info.filename)))
    temp_path.replace(pptx_path)
    with zipfile.ZipFile(pptx_path, "r") as final_zip:
        for chart_name, targets, country_count in verification_parts:
            _verify_chart_xml_targets(final_zip.read(chart_name), targets, etree, country_count)
    logger.info(
        "Applied %s centered per-bar chart label(s) and removed %s stale country-label/leader-line shape(s) "
        "using standard PowerPoint chart XML; think-cell OLE labels were not patched.",
        total_labels,
        total_removed,
    )
