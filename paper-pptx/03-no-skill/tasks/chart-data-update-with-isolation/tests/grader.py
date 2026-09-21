#!/usr/bin/env python3

"""Self-contained Harbor grader for chart-ownership-matrix."""

from __future__ import annotations

import ast

import copy

import hashlib

import importlib.metadata

import importlib.util

import io

import json

import math

import os

import posixpath

import re

import shlex

import shutil

import subprocess

import sys

import tempfile

import time

import types

import unittest.mock

import xml.etree.ElementTree as ET

from collections import Counter, defaultdict

from contextlib import contextmanager

from dataclasses import dataclass, field

from pathlib import Path

from typing import Callable, Iterable

from zipfile import BadZipFile, ZipFile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"

A = "http://schemas.openxmlformats.org/drawingml/2006/main"

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PR = "http://schemas.openxmlformats.org/package/2006/relationships"

P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"

CT = "http://schemas.openxmlformats.org/package/2006/content-types"

X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

NS = {"p": P, "a": A, "c": C, "r": R, "pr": PR, "p14": P14, "ct": CT, "x": X}

REL_SLIDE = R + "/slide"

REL_CHART = R + "/chart"

REL_IMAGE = R + "/image"

REL_PACKAGE = R + "/package"

KNOWN_HASHES = {
    "eval_fixture_01_perception.pptx": "344a0ec5c1c44a59373e7c217e2ca4bec5b8a8b4bad5a04f6243166c1e1b33ff",
    "eval_fixture_02_text.pptx": "ee96ed8e601511e41e3bac73a84741e96964a5ef63b3dae41161d994f5afab6a",
    "eval_fixture_03_objects.pptx": "26bcfb5db1674875744662bf533c521c6764df3c1ffd3115f975f64ab47d78de",
    "eval_fixture_04_structure.pptx": "39b4266cd63cbd9b39751fedf213343a9eed45d2510f4204ef52da822420f6fc",
    "eval_fixture_pair_source.pptx": "1e2298971b3212e9d8554a087b9d07ece39bacad7ca9e7d74ca5716ebbb6f9ec",
    "eval_fixture_pair_destination.pptx": "6a7069efc15fa0ae52b8981b01f2ee860bcab68549c0605eafedfa64ccfba554",
    "eval_fixture_negative_shared_chart_graphs.pptx": "ba8431a81a429349af27e706b3a637276683b8869278b329ee300b8575fc9492",
    "eval_fixture_corrupt_package.zip": "75d7f96daeac1f2b94c5d36ac3181a207c5ae776ddd4d23acb8423f0c210a207",
    "replacement_same_format.png": "b60f709965eb54e1400cba53617ae2c02f8028f5399cafa32d8ef18bfc803915",
    "replacement_cross_format.jpeg": "ad5afb51d11ed1b0bcd80103519ad198129eb9489637b00b9547f208408548f0",
}

class VerificationError(AssertionError):
    pass

@dataclass
class Scorecard:
    """Collect diagnostic checks for binary task grading."""

    checks: list[dict] = field(default_factory=list)

    def record(self, passed: bool, message: str, *, group: str | None = None) -> None:
        self.checks.append(
            {"passed": bool(passed), "message": message, "group": group or _ACTIVE_GROUP}
        )

    def ratio(self, group: str) -> float:
        selected = [item for item in self.checks if item["group"] == group]
        if not selected:
            return 0.0
        return sum(item["passed"] for item in selected) / len(selected)

_ACTIVE_SCORECARD: Scorecard | None = None

_ACTIVE_GROUP = "task"

@contextmanager
def collect_checks(scorecard: Scorecard, group: str):
    global _ACTIVE_SCORECARD, _ACTIVE_GROUP
    previous_card, previous_group = _ACTIVE_SCORECARD, _ACTIVE_GROUP
    _ACTIVE_SCORECARD, _ACTIVE_GROUP = scorecard, group
    try:
        yield
    finally:
        _ACTIVE_SCORECARD, _ACTIVE_GROUP = previous_card, previous_group

def check(condition: bool, message: str) -> None:
    if _ACTIVE_SCORECARD is not None:
        _ACTIVE_SCORECARD.record(bool(condition), message)
        if not condition:
            return
    elif not condition:
        raise VerificationError(message)

def qn(namespace: str, local: str) -> str:
    return "{%s}%s" % (namespace, local)

def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]

def file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

def assert_known_input(path: Path) -> None:
    check(path.is_file(), "missing input %s" % path)
    expected = KNOWN_HASHES.get(path.name)
    check(expected is not None, "no verifier baseline for %s" % path.name)
    check(file_hash(path) == expected, "input fixture was modified: %s" % path)

def package(path: Path) -> dict[str, bytes]:
    check(path.is_file(), "missing PowerPoint output %s" % path)
    try:
        with ZipFile(path) as archive:
            infos = archive.infolist()
            names = [info.filename for info in infos]
            check(len(infos) <= 10000, "%s exceeds the ZIP member-count limit" % path)
            check(sum(info.compress_size for info in infos) <= 512 * 1024 * 1024, "%s exceeds the compressed package-size limit" % path)
            check(sum(info.file_size for info in infos) <= 2 * 1024 * 1024 * 1024, "%s exceeds the expanded package-size limit" % path)
            for info in infos:
                check(not (info.flag_bits & 0x1), "%s contains encrypted member %s" % (path, info.filename))
                check(info.file_size <= 512 * 1024 * 1024, "%s contains oversized member %s" % (path, info.filename))
            check(len(names) == len(set(names)), "%s contains duplicate ZIP members" % path)
            members = {name: archive.read(name) for name in names}
    except (BadZipFile, OSError) as exc:
        raise VerificationError("%s is not a readable OOXML package: %s" % (path, exc)) from exc
    check("[Content_Types].xml" in members, "%s lacks [Content_Types].xml" % path)
    check("ppt/presentation.xml" in members, "%s lacks ppt/presentation.xml" % path)
    return members

def xml(members: dict[str, bytes], member: str) -> ET.Element:
    check(member in members, "missing package member %s" % member)
    try:
        return ET.fromstring(members[member])
    except ET.ParseError as exc:
        raise VerificationError("invalid XML in %s: %s" % (member, exc)) from exc

def rels_member(owner: str) -> str:
    if not owner:
        return "_rels/.rels"
    return posixpath.join(posixpath.dirname(owner), "_rels", posixpath.basename(owner) + ".rels")

def resolve_target(owner: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))

def relationships(members: dict[str, bytes], owner: str) -> dict[str, tuple[str, str, str | None]]:
    member = rels_member(owner)
    if member not in members:
        return {}
    result = {}
    for rel in xml(members, member):
        rid = rel.get("Id")
        target = rel.get("Target", "")
        mode = rel.get("TargetMode")
        result[rid] = (rel.get("Type", ""), target if mode == "External" else resolve_target(owner, target), mode)
    return result

def owner_from_rels(member: str) -> str:
    if member == "_rels/.rels":
        return ""
    directory, filename = posixpath.split(member)
    check(posixpath.basename(directory) == "_rels" and filename.endswith(".rels"), "bad rels member")
    return posixpath.join(posixpath.dirname(directory), filename[:-5])

def check_relationship_integrity(members: dict[str, bytes]) -> None:
    for member in members:
        if not member.endswith(".rels"):
            continue
        owner = owner_from_rels(member)
        for rid, (kind, target, mode) in relationships(members, owner).items():
            if mode == "External":
                continue
            check(target in members, "dangling relationship %s in %s -> %s" % (rid, owner or "/", target))

def check_canonical_member_names(members: dict[str, bytes]) -> None:
    folded: dict[str, str] = {}
    for name in members:
        check(name != "" and not name.startswith("/"), "noncanonical absolute/empty ZIP member %r" % name)
        check("\\" not in name, "noncanonical backslash ZIP member %r" % name)
        normalized = posixpath.normpath(name)
        check(normalized == name and normalized not in {".", ".."}, "noncanonical ZIP member path %r" % name)
        check(not any(part in {"", ".", ".."} for part in name.split("/")), "unsafe ZIP member path %r" % name)
        lowered = name.casefold()
        check(lowered not in folded or folded[lowered] == name, "case-colliding ZIP members %r and %r" % (folded.get(lowered), name))
        folded[lowered] = name

def check_content_types(members: dict[str, bytes]) -> None:
    root = xml(members, "[Content_Types].xml")
    check(root.tag == qn(CT, "Types"), "[Content_Types].xml has the wrong root")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        if child.tag == qn(CT, "Default"):
            extension = child.get("Extension", "").casefold()
            check(bool(extension) and "/" not in extension, "invalid content-type Default extension")
            check(extension not in defaults, "duplicate content-type Default for .%s" % extension)
            defaults[extension] = child.get("ContentType", "")
        elif child.tag == qn(CT, "Override"):
            part = child.get("PartName", "")
            check(part.startswith("/") and part[1:] in members, "content-type Override targets missing part %r" % part)
            check(part not in overrides, "duplicate content-type Override for %s" % part)
            overrides[part] = child.get("ContentType", "")
        else:
            check(False, "unexpected child in [Content_Types].xml: %s" % local(child.tag))
    for name in members:
        if name == "[Content_Types].xml" or name.endswith("/"):
            continue
        extension = posixpath.basename(name).rsplit(".", 1)[-1].casefold() if "." in posixpath.basename(name) else ""
        check(("/" + name) in overrides or extension in defaults, "package member has no content type: %s" % name)

def _check_child_order(parent: ET.Element, order: dict[str, int], member: str) -> None:
    positions = [order[child.tag] for child in parent if child.tag in order]
    check(positions == sorted(positions), "invalid OOXML child order in %s <%s>: %s" % (member, local(parent.tag), [local(child.tag) for child in parent]))

def check_ooxml_child_order(members: dict[str, bytes]) -> None:
    bullet_types = {qn(A, "buNone"), qn(A, "buAutoNum"), qn(A, "buChar"), qn(A, "buBlip")}
    bullet_fonts = {qn(A, "buFontTx"), qn(A, "buFont")}
    paragraph_property_tags = {qn(A, "pPr")} | {qn(A, "lvl%dpPr" % level) for level in range(1, 10)}
    for member in sorted(name for name in members if name.endswith(".xml")):
        root = xml(members, member)
        if root.tag in _ROOT_ORDERS:
            _check_child_order(root, _ROOT_ORDERS[root.tag], member)
        for element in root.iter():
            if element.tag not in paragraph_property_tags:
                continue
            _check_child_order(element, _PPR_ORDER, member)
            check(sum(child.tag in bullet_types for child in element) <= 1, "multiple bullet-type children in %s" % member)
            check(sum(child.tag in bullet_fonts for child in element) <= 1, "multiple bullet-font children in %s" % member)

def check_relationship_references(members: dict[str, bytes]) -> None:
    check_relationship_integrity(members)
    for rels_name in sorted(name for name in members if name.endswith(".rels")):
        owner = owner_from_rels(rels_name)
        if owner:
            check(owner in members, "relationships part has no owner: %s" % rels_name)
        root = xml(members, rels_name)
        ids = [child.get("Id") for child in root]
        check(all(ids) and len(ids) == len(set(ids)), "duplicate/empty relationship IDs in %s" % rels_name)
        rels = relationships(members, owner)
        if not owner or owner not in members or not owner.endswith(".xml"):
            continue
        owner_root = xml(members, owner)
        referenced = []
        for element in owner_root.iter():
            for attr, value in element.attrib.items():
                if attr in {qn(R, "id"), qn(R, "embed"), qn(R, "link")}:
                    referenced.append(value)
        for rid in referenced:
            check(rid in rels, "%s references missing relationship %s" % (owner, rid))

def check_reachability(members: dict[str, bytes]) -> None:
    root_rels = relationships(members, "")
    office = [target for kind, target, mode in root_rels.values() if kind == R + "/officeDocument" and mode != "External"]
    check(len(office) == 1 and office[0] == "ppt/presentation.xml", "package must have exactly one root officeDocument relationship")
    reachable = set()
    pending = [target for kind, target, mode in root_rels.values() if mode != "External"]
    while pending:
        member = pending.pop()
        if member in reachable or member not in members:
            continue
        reachable.add(member)
        for _kind, target, mode in relationships(members, member).values():
            if mode != "External" and target not in reachable:
                pending.append(target)
    payload_members = {
        name
        for name in members
        if name != "[Content_Types].xml" and not name.endswith(".rels") and not name.endswith("/")
    }
    check(payload_members <= reachable, "package contains unreachable parts: %s" % sorted(payload_members - reachable))

def check_presentation_identifiers(members: dict[str, bytes]) -> None:
    records = slide_records(members)
    ids = [record["id"] for record in records]
    rids = [record["rid"] for record in records]
    targets = [record["member"] for record in records]
    check(len(ids) == len(set(ids)) and all(value >= 256 for value in ids), "presentation slide IDs are invalid or duplicated")
    check(len(rids) == len(set(rids)), "presentation slide relationship IDs are duplicated")
    check(len(targets) == len(set(targets)), "presentation aliases the same slide part more than once")

def check_editable_chart_ownership(members: dict[str, bytes], *, allow_shared_graphs: bool = False) -> None:
    chart_owners: defaultdict[str, list[str]] = defaultdict(list)
    for slide in slide_records(members):
        for shape in iter_shapes(slide["root"]):
            chart = shape.find(".//c:chart", NS)
            if chart is None:
                continue
            rid = chart.get(qn(R, "id"))
            rels = relationships(members, slide["member"])
            check(rid in rels and rels[rid][0] == REL_CHART, "chart frame %r has no chart relationship" % shape_name(shape))
            if rid in rels:
                chart_owners[rels[rid][1]].append("%s:%s" % (slide["member"], shape_name(shape)))
    workbook_owners: defaultdict[str, list[str]] = defaultdict(list)
    for chart_member in chart_owners:
        workbook = workbook_part(members, chart_member)
        if workbook:
            workbook_owners[workbook].append(chart_member)
    if not allow_shared_graphs:
        for chart_member, owners in chart_owners.items():
            check(len(owners) == 1, "shared editable chart part %s owned by %s" % (chart_member, owners))
        for workbook, owners in workbook_owners.items():
            check(len(owners) == 1, "shared embedded workbook %s owned by %s" % (workbook, owners))

def _xlsx_cell_values(workbook_bytes: bytes) -> tuple[dict[str, str], dict[str, dict[str, str]]]:
    try:
        with ZipFile(__import__("io").BytesIO(workbook_bytes)) as archive:
            names = archive.namelist()
            check("xl/workbook.xml" in names, "embedded workbook lacks xl/workbook.xml")
            workbook_members = {name: archive.read(name) for name in names}
    except (BadZipFile, OSError) as exc:
        raise VerificationError("embedded workbook is unreadable: %s" % exc) from exc
    wb_root = xml(workbook_members, "xl/workbook.xml")
    wb_rels = relationships(workbook_members, "xl/workbook.xml")
    shared_strings: list[str] = []
    if "xl/sharedStrings.xml" in workbook_members:
        shared_root = xml(workbook_members, "xl/sharedStrings.xml")
        for item in shared_root.findall("./x:si", NS):
            shared_strings.append("".join(node.text or "" for node in item.findall(".//x:t", NS)))
    sheet_targets: dict[str, str] = {}
    sheets: dict[str, dict[str, str]] = {}
    for sheet in wb_root.findall("./x:sheets/x:sheet", NS):
        name = sheet.get("name", "")
        rid = sheet.get(qn(R, "id"))
        check(rid in wb_rels, "workbook sheet %r has no relationship" % name)
        if rid not in wb_rels:
            continue
        target = wb_rels[rid][1]
        sheet_targets[name] = target
        root = xml(workbook_members, target)
        cells: dict[str, str] = {}
        for cell in root.findall(".//x:c", NS):
            ref = cell.get("r", "").replace("$", "")
            kind = cell.get("t")
            if kind == "inlineStr":
                value = "".join(node.text or "" for node in cell.findall(".//x:is/x:t", NS))
            else:
                value_node = cell.find("./x:v", NS)
                value = "" if value_node is None else (value_node.text or "")
                if kind == "s" and value.isdigit() and int(value) < len(shared_strings):
                    value = shared_strings[int(value)]
                elif kind == "b":
                    value = "TRUE" if value == "1" else "FALSE"
            cells[ref] = value
        sheets[name] = cells
    return sheet_targets, sheets

def _column_number(token: str) -> int:
    result = 0
    for char in token:
        result = result * 26 + ord(char) - 64
    return result

def _column_token(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result

def _range_values(sheets: dict[str, dict[str, str]], formula: str) -> list[str] | None:
    match = _CELL_RANGE.match(formula.strip())
    if match is None:
        return None
    sheet = (match.group(1) or match.group(2) or "").replace("''", "'")
    if sheet not in sheets:
        return None
    col1, row1 = _column_number(match.group(3)), int(match.group(4))
    col2, row2 = _column_number(match.group(5) or match.group(3)), int(match.group(6) or match.group(4))
    values = []
    for row in range(row1, row2 + 1):
        for col in range(col1, col2 + 1):
            values.append(sheets[sheet].get("%s%d" % (_column_token(col), row), ""))
    return values

def _normalized_scalar(value: str) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isfinite(number) and number.is_integer():
        return str(int(number))
    return format(number, ".15g")

def check_chart_cache_workbook_agreement(members: dict[str, bytes]) -> None:
    for chart_member in sorted(name for name in members if name.endswith(".xml") and "/charts/" in name and "/_rels/" not in name):
        workbook = workbook_part(members, chart_member)
        if workbook is None:
            continue
        check(workbook in members, "%s targets missing workbook %s" % (chart_member, workbook))
        if workbook not in members:
            continue
        _, sheets = _xlsx_cell_values(members[workbook])
        chart_root = xml(members, chart_member)
        for reference in list(chart_root.findall(".//c:numRef", NS)) + list(chart_root.findall(".//c:strRef", NS)):
            formula_node = reference.find("./c:f", NS)
            cache = reference.find("./c:numCache", NS)
            if cache is None:
                cache = reference.find("./c:strCache", NS)
            if formula_node is None or cache is None or not (formula_node.text or "").strip():
                continue
            expected = _range_values(sheets, formula_node.text or "")
            check(expected is not None, "%s contains an unsupported or broken workbook formula %r" % (chart_member, formula_node.text))
            if expected is None:
                continue
            points = {int(point.get("idx", "-1")): (point.findtext("./c:v", default="", namespaces=NS)) for point in cache.findall("./c:pt", NS)}
            actual = [points.get(index, "") for index in range(len(expected))]
            check([_normalized_scalar(value) for value in actual] == [_normalized_scalar(value) for value in expected], "%s cache disagrees with workbook range %s" % (chart_member, formula_node.text))

def validate_pptx_package(path: Path, *, allow_shared_graphs: bool = False) -> dict[str, bytes]:
    members = package(path)
    check_canonical_member_names(members)
    check_content_types(members)
    for member in sorted(name for name in members if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml"):
        xml(members, member)
    check_relationship_references(members)
    check_reachability(members)
    check_ooxml_child_order(members)
    check_presentation_identifiers(members)
    check_editable_chart_ownership(members, allow_shared_graphs=allow_shared_graphs)
    check_chart_cache_workbook_agreement(members)
    return members

def slide_records(members: dict[str, bytes]) -> list[dict]:
    pres = xml(members, "ppt/presentation.xml")
    rels = relationships(members, "ppt/presentation.xml")
    records = []
    for index, sld_id in enumerate(pres.findall("./p:sldIdLst/p:sldId", NS), 1):
        rid = sld_id.get(qn(R, "id"))
        check(rid in rels and rels[rid][0] == REL_SLIDE, "invalid presentation slide relationship")
        member = rels[rid][1]
        root = xml(members, member)
        records.append(
            {
                "index": index,
                "id": int(sld_id.get("id")),
                "rid": rid,
                "member": member,
                "root": root,
                "title": slide_title(root),
            }
        )
    return records

def iter_shapes(root: ET.Element) -> Iterable[ET.Element]:
    for element in root.iter():
        if element.tag in SHAPE_TAGS:
            yield element

def c_nv_pr(shape: ET.Element) -> ET.Element | None:
    paths = {
        qn(P, "sp"): "./p:nvSpPr/p:cNvPr",
        qn(P, "pic"): "./p:nvPicPr/p:cNvPr",
        qn(P, "graphicFrame"): "./p:nvGraphicFramePr/p:cNvPr",
        qn(P, "grpSp"): "./p:nvGrpSpPr/p:cNvPr",
    }
    return shape.find(paths[shape.tag], NS)

def shape_name(shape: ET.Element) -> str:
    node = c_nv_pr(shape)
    return "" if node is None else node.get("name", "")

def shape_placeholder(shape: ET.Element) -> ET.Element | None:
    if shape.tag != qn(P, "sp"):
        return None
    return shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)

def shape_text(shape: ET.Element) -> str:
    return "".join((node.text or "") for node in shape.findall(".//a:t", NS))

def slide_title(root: ET.Element) -> str:
    candidates = list(iter_shapes(root))
    for shape in candidates:
        ph = shape_placeholder(shape)
        if ph is not None and ph.get("type") in {"title", "ctrTitle"}:
            return shape_text(shape)
    for shape in candidates:
        if shape_name(shape).startswith("Slide Title"):
            return shape_text(shape)
    for shape in candidates:
        text = shape_text(shape)
        if text:
            return text
    return ""

def slide_by_title(members: dict[str, bytes], title: str) -> dict:
    matches = [record for record in slide_records(members) if record["title"] == title]
    check(len(matches) == 1, "expected one slide titled %r, found %d" % (title, len(matches)))
    return matches[0]

def shapes_named(root: ET.Element, name: str) -> list[ET.Element]:
    return [shape for shape in iter_shapes(root) if shape_name(shape) == name]

def one_shape(root: ET.Element, name: str) -> ET.Element:
    matches = shapes_named(root, name)
    check(len(matches) == 1, "expected one shape %r, found %d" % (name, len(matches)))
    return matches[0]

def node_signature(
    element: ET.Element,
    *,
    ignore_attributes: set[str] | None = None,
    ignore_tags: set[str] | None = None,
    ignore_text: bool = False,
):
    ignore_attributes = ignore_attributes or set()
    ignore_tags = ignore_tags or set()
    if element.tag in ignore_tags:
        return None
    attrs = tuple(sorted((key, value) for key, value in element.attrib.items() if key not in ignore_attributes))
    text = "" if ignore_text else (element.text or "")
    if element.tag != qn(A, "t") and text.isspace():
        text = ""
    children = []
    for child in element:
        signature = node_signature(
            child,
            ignore_attributes=ignore_attributes,
            ignore_tags=ignore_tags,
            ignore_text=ignore_text,
        )
        if signature is not None:
            children.append(signature)
    return element.tag, attrs, text, tuple(children)

def semantic_member_signature(members: dict[str, bytes], member: str):
    if member == "[Content_Types].xml":
        root = xml(members, member)
        defaults = {child.get("Extension", "").casefold(): child.get("ContentType", "") for child in root if child.tag == qn(CT, "Default")}
        overrides = {child.get("PartName", "").lstrip("/"): child.get("ContentType", "") for child in root if child.tag == qn(CT, "Override")}
        resolved = []
        for name in members:
            if name == "[Content_Types].xml" or name.endswith("/"):
                continue
            extension = posixpath.basename(name).rsplit(".", 1)[-1].casefold() if "." in posixpath.basename(name) else ""
            resolved.append((name, overrides.get(name, defaults.get(extension))))
        return tuple(sorted(resolved))
    if member.endswith(".rels"):
        owner = owner_from_rels(member)
        return tuple(sorted((rid, kind, target, mode) for rid, (kind, target, mode) in relationships(members, owner).items()))
    if member.endswith(".xml"):
        root = xml(members, member)
        return node_signature(root)
    return hashlib.sha256(members[member]).hexdigest()

def semantic_changed_members(before: dict[str, bytes], after: dict[str, bytes]) -> list[str]:
    return sorted(
        name
        for name in set(before) | set(after)
        if name not in before
        or name not in after
        or semantic_member_signature(before, name) != semantic_member_signature(after, name)
    )

def chart_part(members: dict[str, bytes], slide: dict, shape_name_value: str) -> str:
    shape = one_shape(slide["root"], shape_name_value)
    chart = shape.find(".//c:chart", NS)
    check(chart is not None, "%s is not a chart frame" % shape_name_value)
    rid = chart.get(qn(R, "id"))
    rels = relationships(members, slide["member"])
    check(rid in rels and rels[rid][0] == REL_CHART, "chart relationship missing for %s" % shape_name_value)
    return rels[rid][1]

def workbook_part(members: dict[str, bytes], chart_member: str) -> str | None:
    workbooks = [target for kind, target, mode in relationships(members, chart_member).values() if kind == REL_PACKAGE and mode != "External"]
    check(len(workbooks) <= 1, "%s owns multiple workbooks" % chart_member)
    return workbooks[0] if workbooks else None

def chart_series(members: dict[str, bytes], chart_member: str):
    root = xml(members, chart_member)
    series = root.findall(".//c:ser", NS)
    check(series, "%s has no chart series" % chart_member)
    result = []
    for ser in series:
        names = [(node.text or "") for node in ser.findall("./c:tx//c:v", NS)]
        cats = [(node.text or "") for node in ser.findall("./c:cat//c:v", NS)]
        if not cats:
            cats = [(node.text or "") for node in ser.findall("./c:xVal//c:v", NS)]
        vals = [(node.text or "") for node in ser.findall("./c:val//c:v", NS)]
        if not vals:
            vals = [(node.text or "") for node in ser.findall("./c:yVal//c:v", NS)]
        result.append({"name": names[-1] if names else "", "categories": cats,
                       "values": [_normalized_scalar(value) for value in vals]})
    return result

def chart_visual_signature(members: dict[str, bytes], chart_member: str):
    data_tags = {
        qn(C, "tx"),
        qn(C, "cat"),
        qn(C, "val"),
        qn(C, "xVal"),
        qn(C, "yVal"),
        qn(C, "bubbleSize"),
        qn(C, "externalData"),
    }
    return node_signature(xml(members, chart_member), ignore_tags=data_tags)

def check_chart_formula_workbook(members: dict[str, bytes], chart_member: str, series_name: str, categories: list[str], values: list[str]) -> None:
    workbook = workbook_part(members, chart_member)
    check(workbook is not None, "%s must retain an embedded workbook" % chart_member)
    if workbook is None:
        return
    root = xml(members, chart_member)
    _, sheets = _xlsx_cell_values(members[workbook])
    series = root.findall(".//c:ser", NS)
    check(len(series) == 1, "%s requires one editable series" % chart_member)
    if len(series) != 1:
        return
    for field, wanted, numeric in (("tx", [series_name], False), ("cat", categories, False), ("val", values, True)):
        formula = series[0].find("./c:%s//c:f" % field, NS)
        actual = _range_values(sheets, formula.text or "") if formula is not None else None
        if numeric and actual is not None:
            actual = [_normalized_scalar(value) for value in actual]
            wanted = [_normalized_scalar(value) for value in wanted]
        check(actual == wanted, "%s workbook %s range disagrees with requested chart data" % (chart_member, field))

def assert_one_series(members: dict[str, bytes], chart_member: str, name: str, categories: list[str], values: list[str]) -> None:
    series = chart_series(members, chart_member)
    check(len(series) == 1, "%s should contain one series" % chart_member)
    check(series[0]["name"] == name, "%s series name mismatch: %r" % (chart_member, series[0]["name"]))
    check(series[0]["categories"] == categories, "%s categories mismatch: %r" % (chart_member, series[0]["categories"]))
    check(series[0]["values"] == values, "%s values mismatch: %r" % (chart_member, series[0]["values"]))

def task_dir(root: Path, task: str) -> Path:
    return root / "evals" / task

def check_no_presentation_output(directory: Path) -> None:
    if not directory.exists():
        return
    package_suffixes = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot", ".potx", ".potm", ".zip"}
    forbidden = [
        path
        for path in directory.rglob("*")
        if (path.is_file() and (path.suffix.lower() in package_suffixes or path.name.endswith((".partial", ".tmp"))))
        or (path.is_file() and path.name == "[Content_Types].xml")
        or (path.is_dir() and path.name in {"ppt", "_rels"})
    ]
    check(not forbidden, "task must not emit a presentation/package: %s" % forbidden)

def run_powerpoint_open_gate(command_text: str, path: Path) -> dict:
    """Run a trusted evaluator command and require a real-PowerPoint attestation."""
    tokens = shlex.split(command_text)
    check(bool(tokens), "PowerPoint open-gate command is empty")
    if not tokens:
        return {}
    check(any("{pptx}" in token for token in tokens), "PowerPoint open-gate command must contain {pptx}")
    check(any("{result_json}" in token for token in tokens), "PowerPoint open-gate command must contain {result_json}")
    with tempfile.TemporaryDirectory(prefix="paper-pptx-powerpoint-gate-") as temporary:
        result_path = Path(temporary) / "attestation.json"
        command = [
            token.replace("{pptx}", str(path)).replace("{result_json}", str(result_path))
            for token in tokens
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=180)
        check(result.returncode == 0, "PowerPoint open gate rejected %s: %s" % (path, result.stderr[-2000:]))
        check(result_path.is_file(), "PowerPoint open gate did not emit its required attestation")
        try:
            attestation = json.loads(result_path.read_text())
        except (OSError, json.JSONDecodeError) as exc:
            raise VerificationError("invalid PowerPoint open-gate attestation: %s" % exc) from exc
    check(isinstance(attestation, dict), "PowerPoint open-gate attestation must be a JSON object")
    if not isinstance(attestation, dict):
        return {}
    check(attestation.get("schema") == "paper-pptx-powerpoint-open-gate", "PowerPoint attestation schema mismatch")
    check(attestation.get("version") == 1, "PowerPoint attestation version mismatch")
    check(attestation.get("application") == "Microsoft PowerPoint", "open gate did not attest Microsoft PowerPoint")
    check(attestation.get("platform") in {"macOS", "Windows"}, "PowerPoint gate platform must be macOS or Windows")
    check(attestation.get("pptx_sha256") == file_hash(path), "PowerPoint attestation is for different PPTX bytes")
    check(attestation.get("opened_without_repair") is True, "PowerPoint did not open the deck directly without Repair")
    check(attestation.get("repair_warning") is False, "PowerPoint displayed a repair warning")
    check(attestation.get("repair_accepted") is False, "PowerPoint Repair must never be accepted")
    return {
        "application": attestation.get("application"),
        "platform": attestation.get("platform"),
        "pptx_sha256": attestation.get("pptx_sha256"),
        "opened_without_repair": attestation.get("opened_without_repair"),
        "repair_warning": attestation.get("repair_warning"),
        "repair_accepted": attestation.get("repair_accepted"),
        "returncode": result.returncode,
        "stdout": result.stdout[-2000:],
        "stderr": result.stderr[-2000:],
    }

# BEGIN MATERIALIZED GRADER RUNTIME

def integrity_gate(function):
    """Validate package structure without awarding points for XML bookkeeping."""
    def guarded(*args, **kwargs):
        global _ACTIVE_SCORECARD
        previous = _ACTIVE_SCORECARD
        _ACTIVE_SCORECARD = None
        try:
            return function(*args, **kwargs)
        finally:
            _ACTIVE_SCORECARD = previous
    return guarded


for _gate_name in (
    "package", "xml", "check_relationship_integrity", "check_canonical_member_names",
    "check_content_types", "check_ooxml_child_order", "check_relationship_references",
    "check_reachability", "check_presentation_identifiers", "check_editable_chart_ownership",
    "check_chart_cache_workbook_agreement", "validate_pptx_package",
):
    globals()[_gate_name] = integrity_gate(globals()[_gate_name])

def _write_grading(path: Path | None, payload: dict) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")

def _write_reward(path: Path | None, reward: float) -> None:
    if path is None:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(("%.4f" % reward).rstrip("0").rstrip(".") + "\n")

def _result_destinations() -> tuple[Path | None, Path | None, int | None]:
    if len(sys.argv) == 4 and sys.argv[2] == "--result-fd":
        try:
            return None, None, int(sys.argv[3])
        except ValueError as exc:
            raise VerificationError("result FD must be an integer") from exc
    check(2 <= len(sys.argv) <= 4, "usage: grader.py ROOT [REWARD_PATH] [GRADING_PATH]")
    reward_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else None
    grading_path = Path(sys.argv[3]) if len(sys.argv) >= 4 else None
    return reward_path, grading_path, None

def _publish_result(payload: dict) -> None:
    reward_path, grading_path, result_fd = _result_destinations()
    if result_fd is not None:
        envelope = (
            json.dumps(
                {
                    "schema": "paper-pptx-verifier-result-envelope",
                    "version": 1,
                    "reward": payload.get("reward", 0.0),
                    "grading": payload,
                },
                sort_keys=True,
            ).encode()
            + b"\n"
        )
        view = memoryview(envelope)
        while view:
            written = os.write(result_fd, view)
            if written <= 0:
                raise VerificationError("could not write verifier result envelope")
            view = view[written:]
        return
    _write_reward(reward_path, float(payload.get("reward", 0.0)))
    _write_grading(grading_path, payload)

def _pattern_status(scorecard: Scorecard, patterns: tuple[str, ...]) -> dict[str, bool]:
    status = {}
    for pattern in patterns:
        matches = [
            item
            for item in scorecard.checks
            if pattern.casefold() in str(item["message"]).casefold()
        ]
        # A missing assertion is missing evidence, not a pass.
        status[pattern] = bool(matches) and all(item["passed"] for item in matches)
    return status

@dataclass(frozen=True)
class TaskConfig:
    task_id: str
    inputs: tuple[str, ...]
    outputs: tuple[str, ...]
    critical_effects: tuple[str, ...] = ()
    critical_collateral: tuple[str, ...] = ()

def hard_preflight(
    root: Path,
    config: TaskConfig,
) -> dict:
    task = config.task_id
    details = {"inputs": {}, "outputs": {}, "powerpoint_open_gates": []}
    for relative in config.inputs:
        path = root / relative
        assert_known_input(path)
        details["inputs"][relative] = file_hash(path)
    if not config.outputs:
        check_no_presentation_output(task_dir(root, task))
    open_gate = os.environ.get("PAPER_PPTX_OPEN_GATE_CMD")
    if config.outputs and os.environ.get("PAPER_PPTX_REQUIRE_POWERPOINT") == "1":
        check(
            bool(open_gate),
            "PowerPoint gate is required but PAPER_PPTX_OPEN_GATE_CMD is unset",
        )
    for filename in config.outputs:
        path = task_dir(root, task) / filename
        members = validate_pptx_package(path)
        details["outputs"][filename] = {
            "sha256": file_hash(path),
            "members": len(members),
            "slides": len(slide_records(members)),
        }
        if open_gate:
            details["powerpoint_open_gates"].append(
                run_powerpoint_open_gate(open_gate, path)
            )
    return details

def score_results(
    config: TaskConfig,
    scorecard: Scorecard,
    hard_failures: list[str],
) -> tuple[
    float, dict[str, float], dict[str, dict[str, bool]], list[dict[str, object]]
]:
    selected = [item for item in scorecard.checks if item["group"] == "task"]
    critical = {
        "effect": _pattern_status(scorecard, config.critical_effects),
        "collateral": _pattern_status(scorecard, config.critical_collateral),
    }
    effect_checks = [
        item for item in selected
        if any(pattern.casefold() in item["message"].casefold()
               for pattern in config.critical_effects)
    ]
    preservation_checks = [item for item in selected if item not in effect_checks]
    effect_ok = bool(effect_checks) and all(item["passed"] for item in effect_checks)
    if not config.critical_effects:
        effect_ok = bool(selected) and all(item["passed"] for item in selected)
    effect_ok = effect_ok and all(critical["effect"].values())
    preservation_ok = all(item["passed"] for item in preservation_checks)
    preservation_ok = preservation_ok and all(critical["collateral"].values())
    group_scores = {"effect": float(effect_ok), "preservation": float(preservation_ok)}
    reward = float(
        effect_ok
        and preservation_ok
        and not hard_failures
        and all(item["passed"] for item in scorecard.checks)
    )
    return reward, group_scores, critical, []

def run_task(
    config: TaskConfig,
    verifier: Callable[[Path, str], None],
) -> None:
    _result_destinations()
    root = Path(sys.argv[1]).resolve()
    task = config.task_id
    hard_failures: list[str] = []
    hard_details = {}
    try:
        hard_details = hard_preflight(root, config)
    except Exception as exc:
        hard_failures.append("preflight aborted: %s: %s" % (type(exc).__name__, exc))

    scorecard = Scorecard()
    if not hard_failures:
        try:
            with collect_checks(scorecard, "task"):
                verifier(root, task)
        except Exception as exc:
            hard_failures.append(
                "task verifier aborted: %s: %s" % (type(exc).__name__, exc)
            )
    reward, group_scores, critical, caps = score_results(
        config,
        scorecard,
        hard_failures,
    )
    payload = {
        "schema": "paper-pptx-harbor-grading",
        "version": 3,
        "task": task,
        "reward": reward,
        "task_success": reward == 1.0 and not hard_failures,
        "hard_failures": hard_failures,
        "hard_gate_details": hard_details,
        "group_scores": group_scores,
        "critical": critical,
        "caps": caps,
        "checks": scorecard.checks,
    }
    _publish_result(payload)
    print(
        json.dumps(
            {
                "schema": payload["schema"],
                "task": task,
                "reward": reward,
                "hard_failures": hard_failures,
                "failed_checks": [
                    item["message"] for item in scorecard.checks if not item["passed"]
                ],
            },
            sort_keys=True,
        )
    )

def grader_entrypoint(main_function: Callable[[], None]) -> None:
    try:
        main_function()
    except Exception as exc:
        _publish_result(
            {
                "schema": "paper-pptx-harbor-grading",
                "version": 2,
                "reward": 0.0,
                "hard_failures": ["grader crash: %s: %s" % (type(exc).__name__, exc)],
            }
        )
        print("FAIL:", exc, file=sys.stderr)

# END MATERIALIZED GRADER RUNTIME

_PPR_ORDER = {
    qn(A, "lnSpc"): 0,
    qn(A, "spcBef"): 1,
    qn(A, "spcAft"): 2,
    qn(A, "buClrTx"): 3,
    qn(A, "buClr"): 3,
    qn(A, "buSzTx"): 4,
    qn(A, "buSzPct"): 4,
    qn(A, "buSzPts"): 4,
    qn(A, "buFontTx"): 5,
    qn(A, "buFont"): 5,
    qn(A, "buNone"): 6,
    qn(A, "buAutoNum"): 6,
    qn(A, "buChar"): 6,
    qn(A, "buBlip"): 6,
    qn(A, "tabLst"): 7,
    qn(A, "defRPr"): 8,
    qn(A, "extLst"): 9,
}

_ROOT_ORDERS = {
    qn(P, "presentation"): {
        qn(P, "sldMasterIdLst"): 0,
        qn(P, "notesMasterIdLst"): 1,
        qn(P, "handoutMasterIdLst"): 2,
        qn(P, "sldIdLst"): 3,
        qn(P, "sldSz"): 4,
        qn(P, "notesSz"): 5,
        qn(P, "smartTags"): 6,
        qn(P, "embeddedFontLst"): 7,
        qn(P, "custShowLst"): 8,
        qn(P, "photoAlbum"): 9,
        qn(P, "custDataLst"): 10,
        qn(P, "kinsoku"): 11,
        qn(P, "defaultTextStyle"): 12,
        qn(P, "modifyVerifier"): 13,
        qn(P, "extLst"): 14,
    },
    qn(P, "sldMaster"): {
        qn(P, "cSld"): 0,
        qn(P, "clrMap"): 1,
        qn(P, "sldLayoutIdLst"): 2,
        qn(P, "transition"): 3,
        qn(P, "timing"): 4,
        qn(P, "hf"): 5,
        qn(P, "txStyles"): 6,
        qn(P, "extLst"): 7,
    },
    qn(P, "sldLayout"): {
        qn(P, "cSld"): 0,
        qn(P, "clrMapOvr"): 1,
        qn(P, "transition"): 2,
        qn(P, "timing"): 3,
        qn(P, "hf"): 4,
        qn(P, "extLst"): 5,
    },
    qn(P, "sld"): {
        qn(P, "cSld"): 0,
        qn(P, "clrMapOvr"): 1,
        qn(P, "transition"): 2,
        qn(P, "timing"): 3,
        qn(P, "extLst"): 4,
    },
    qn(P, "notes"): {
        qn(P, "cSld"): 0,
        qn(P, "clrMapOvr"): 1,
        qn(P, "extLst"): 2,
    },
    qn(P, "notesMaster"): {
        qn(P, "cSld"): 0,
        qn(P, "clrMap"): 1,
        qn(P, "hf"): 2,
        qn(P, "notesStyle"): 3,
        qn(P, "extLst"): 4,
    },
    qn(C, "chartSpace"): {
        qn(C, "date1904"): 0,
        qn(C, "lang"): 1,
        qn(C, "roundedCorners"): 2,
        qn(C, "style"): 3,
        qn(C, "clrMapOvr"): 4,
        qn(C, "pivotSource"): 5,
        qn(C, "protection"): 6,
        qn(C, "chart"): 7,
        qn(C, "spPr"): 8,
        qn(C, "txPr"): 9,
        qn(C, "externalData"): 10,
        qn(C, "printSettings"): 11,
        qn(C, "userShapes"): 12,
        qn(C, "extLst"): 13,
    },
}

_CELL_RANGE = re.compile(r"^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$")

SHAPE_TAGS = {qn(P, "sp"), qn(P, "pic"), qn(P, "graphicFrame"), qn(P, "grpSp")}

# Retained-task preservation helpers

def normalized_content(element, *, ignore_attributes=(), ignore_tags=()):
    """Compare formatted text independently of adjacent equal-style run splitting."""
    node = copy.deepcopy(element)
    for value in node.findall(".//c:numCache/c:pt/c:v", NS) + node.findall(".//c:numLit/c:pt/c:v", NS):
        value.text = _normalized_scalar(value.text or "")
    ignored = set(ignore_tags)
    for parent in reversed(list(node.iter())):
        for child in list(parent):
            if child.tag in ignored:
                parent.remove(child)
            elif (child.tag in {qn(A, "pPr"), qn(A, "rPr"), qn(A, "defRPr"), qn(A, "extLst"), qn(P, "extLst")}
                  and not child.attrib and not len(child) and not (child.text or "").strip()):
                parent.remove(child)
            elif local(child.tag) == "ext" and ignored and not len(child) and not (child.text or "").strip():
                parent.remove(child)
    for paragraph in node.iter(qn(A, "p")):
        previous = None
        for child in list(paragraph):
            if child.tag != qn(A, "r"):
                previous = None
                continue
            text = child.find(qn(A, "t"))
            props = child.find(qn(A, "rPr"))
            key = node_signature(props) if props is not None else None
            if previous is not None and previous[0] == key:
                previous[1].text = (previous[1].text or "") + (text.text or "")
                paragraph.remove(child)
            elif text is not None:
                previous = (key, text)
    return node_signature(node, ignore_attributes=set(ignore_attributes), ignore_tags=set(ignore_tags))


def replace_plain_span(shape, old, new):
    """Construct the expected replacement using the first affected run's style."""
    for paragraph in shape.findall(".//a:p", NS):
        runs = paragraph.findall("./a:r", NS)
        text = "".join(run.findtext("./a:t", default="", namespaces=NS) for run in runs)
        if old not in text:
            continue
        start, end = text.index(old), text.index(old) + len(old)
        position = 0
        inserted = False
        for run in runs:
            leaf = run.find("./a:t", NS)
            value = leaf.text or ""
            stop = position + len(value)
            if stop > start and position < end:
                prefix = value[:max(0, start - position)]
                suffix = value[max(0, end - position):]
                leaf.text = prefix + (new if not inserted else "") + suffix
                inserted = True
                if not leaf.text:
                    paragraph.remove(run)
            position = stop
        return
    raise VerificationError("expected replacement span is absent from source")


def preserved_members(before, after, mutable=()):
    """Protect all retained dependencies; mutable owners are checked by the caller."""
    mutable = set(mutable)
    mapping = {}
    after_slides = {slide["id"]: slide for slide in slide_records(after)}
    for slide in slide_records(before):
        if slide["id"] in after_slides:
            mapping[slide["member"]] = after_slides[slide["id"]]["member"]
    pending = list(mapping)
    seen = set()
    while pending:
        owner = pending.pop()
        if owner in seen or owner in mutable:
            continue
        seen.add(owner)
        current = mapping[owner]
        mapping[rels_member(owner)] = rels_member(current)
        old_rels, new_rels = relationships(before, owner), relationships(after, current)
        for rid, (kind, target, mode) in old_rels.items():
            candidate = new_rels.get(rid)
            if mode != "External" and candidate is not None and candidate[0] == kind and candidate[2] != "External":
                if target not in mapping:
                    mapping[target] = candidate[1]
                    pending.append(target)
    failures = []
    for name in before:
        if name in mutable or name == "[Content_Types].xml":
            continue
        target = mapping.get(name, name)
        expected = retained_member_signature(before, name)
        if name.endswith(".rels"):
            expected = tuple((rid, (kind, mapping.get(part, part) if mode != "External" else part, mode))
                             for rid, (kind, part, mode) in expected)
        if target not in after or expected != retained_member_signature(after, target):
            failures.append(name)
    check(not failures, "unrelated package content changed: %s" % failures)


def retained_member_signature(members, name):
    if name.endswith(".rels"):
        return tuple(sorted(relationships(members, owner_from_rels(name)).items()))
    if name.endswith(".xml"):
        return normalized_content(xml(members, name))
    if name.endswith(".xlsx"):
        from io import BytesIO
        with ZipFile(BytesIO(members[name])) as archive:
            return tuple(sorted((entry, normalized_content(ET.fromstring(archive.read(entry)))
                                 if entry.endswith((".xml", ".rels")) else archive.read(entry))
                                for entry in archive.namelist() if not entry.endswith("/")))
    return members[name]


if "one_shape" not in globals():
    def one_shape(root, name):
        matches = [shape for shape in iter_shapes(root) if shape_name(shape) == name]
        check(len(matches) == 1, "expected exactly one shape named %s" % name)
        return matches[0]


if "geometry_values" not in globals():
    def geometry_values(shape):
        transform = shape.find("./p:spPr/a:xfrm", NS)
        if transform is None:
            transform = shape.find("./p:xfrm", NS)
        if transform is None:
            return None
        offset, extent = transform.find("./a:off", NS), transform.find("./a:ext", NS)
        if offset is None or extent is None:
            return None
        return (int(offset.get("x", "0")), int(offset.get("y", "0")),
                int(extent.get("cx", "0")), int(extent.get("cy", "0")), int(transform.get("rot", "0")))


def relationship_content(members, member):
    """Compare relationship meaning without prescribing relationship identifiers."""
    return sorted(relationships(members, member).values())


def placement_matches(shape, expected):
    # A hundredth of a point tolerates unit conversion, not visible movement.
    actual = geometry_values(shape)
    return actual is not None and all(abs(a - b) <= 127 for a, b in zip(actual[:4], expected))


def shape_appearance(shape):
    return normalized_content(
        shape,
        ignore_attributes={"id", "name", qn(R, "id"), qn(R, "embed"), qn(R, "link")},
        ignore_tags={qn("http://schemas.microsoft.com/office/drawing/2014/main", "creationId"),
                     qn("http://schemas.microsoft.com/office/powerpoint/2010/main", "creationId")},
    )


def all_shapes_in_bounds(members, records):
    size = xml(members, "ppt/presentation.xml").find("./p:sldSz", NS)
    width, height = int(size.get("cx")), int(size.get("cy"))
    for record in records:
        check(record["root"].get("show", "1") not in {"0", "false"}, "required slide is hidden")
        boxes = []
        for shape in iter_shapes(record["root"]):
            geometry = geometry_values(shape)
            if geometry is None:
                placeholder = shape_placeholder(shape)
                if placeholder is not None:
                    layouts = slide_relationship_target(members, record, REL_LAYOUT)
                    if layouts:
                        for candidate in iter_shapes(xml(members, layouts[0])):
                            ph = shape_placeholder(candidate)
                            if ph is not None and ph.get("idx", "0") == placeholder.get("idx", "0"):
                                geometry = geometry_values(candidate)
                                break
            if geometry is not None:
                x, y, cx, cy, _ = geometry
                check(cx > 0 and cy > 0 and x >= 0 and y >= 0 and x + cx <= width and y + cy <= height,
                      "shape falls outside slide bounds: %s" % shape_name(shape))
                if shape_text(shape) or shape.find(".//a:tbl", NS) is not None or shape.find(".//c:chart", NS) is not None or shape.tag == qn(P, "pic"):
                    boxes.append((shape_name(shape), x, y, cx, cy))
            for props in shape.findall(".//a:rPr", NS):
                check(int(props.get("sz", "1200")) >= 1200, "text is too small to read: %s" % shape_name(shape))
        for index, (name, x, y, cx, cy) in enumerate(boxes):
            for other, ox, oy, ocx, ocy in boxes[index + 1:]:
                overlap = min(x + cx, ox + ocx) - max(x, ox) > 9144 and min(y + cy, oy + ocy) - max(y, oy) > 9144
                check(not overlap, "slide content overlaps: %s and %s" % (name, other))

KNOWN_HASHES.update({'eval_fixture_03_objects.pptx': 'ce8ce3fa8281d3a87898874a1d729023a1dec3db145c32e0da598b3fc77aa5d3'})

TASK = TaskConfig(
    task_id='chart-data-update-with-isolation',
    inputs=('eval_fixtures/pptx/eval_fixture_03_objects.pptx',),
    outputs=('output.pptx',),
    critical_effects=('should contain one series', 'series name mismatch', 'categories mismatch', 'values mismatch', 'range disagrees with requested chart data', 'Historical forecast gained a workbook'),
)

def verify_chart_matrix(root: Path, task: str) -> None:
    source = root / "eval_fixtures/pptx/eval_fixture_03_objects.pptx"
    before = package(source)
    out = validate_pptx_package(task_dir(root, task) / "output.pptx")
    specs = {
        "Current forecast": ("FY27", ["North", "South"], ["10", "20"]),
        "North forecast": ("FY27", ["North", "South"], ["11", "21"]),
        "South forecast": ("FY27", ["North", "South"], ["12", "22"]),
        "Historical forecast": ("FY27", ["North", "South"], ["3", "6"]),
    }
    slide4 = slide_by_title(out, "Regional forecast")
    slide5 = slide_by_title(out, "Regional comparison")
    slide_for = {name: slide4 for name in ("Current forecast", "North forecast")}
    slide_for.update({name: slide5 for name in ("South forecast", "Historical forecast")})
    mutable = set()
    for name, (series_name, cats, vals) in specs.items():
        chart = chart_part(out, slide_for[name], name)
        assert_one_series(out, chart, series_name, cats, vals)
        before_chart = chart_part(before, slide_by_title(before, slide_for[name]["title"]), name)
        mutable.update({before_chart, rels_member(before_chart), workbook_part(before, before_chart)})
        check(chart_visual_signature(out, chart) == chart_visual_signature(before, before_chart), "%s chart type/visual formatting changed" % name)
        if name != "Historical forecast":
            check_chart_formula_workbook(out, chart, series_name, cats, vals)
    check(workbook_part(out, chart_part(out, slide5, "Historical forecast")) is None, "Historical forecast gained a workbook")
    for out_slide, before_slide, name in ((slide4, slide_by_title(before, slide4["title"]), "North reference"), (slide5, slide_by_title(before, slide5["title"]), "South reference")):
        out_chart = chart_part(out, out_slide, name)
        before_chart = chart_part(before, before_slide, name)
        check(chart_series(out, out_chart) == chart_series(before, before_chart), "%s changed" % name)
        check(normalized_content(xml(out, out_chart)) == normalized_content(xml(before, before_chart)), "%s chart formatting changed" % name)
        check(retained_member_signature(out, workbook_part(out, out_chart)) == retained_member_signature(before, workbook_part(before, before_chart)), "%s workbook data or formatting changed" % name)
    for current in (slide4, slide5):
        baseline = slide_by_title(before, current["title"])
        check(normalized_content(current["root"], ignore_attributes={qn(R, "id")}) == normalized_content(baseline["root"], ignore_attributes={qn(R, "id")}), "chart update changed slide content or geometry")
        check(sorted(value for value in relationships(before, baseline["member"]).values() if value[0] != REL_CHART) == sorted(value for value in relationships(out, current["member"]).values() if value[0] != REL_CHART), "chart update changed an unrelated relationship")
        mutable.update({baseline["member"], rels_member(baseline["member"])})
    preserved_members(before, out, mutable)

def main() -> None:
    run_task(TASK, verify_chart_matrix)

if __name__ == "__main__":
    grader_entrypoint(main)
