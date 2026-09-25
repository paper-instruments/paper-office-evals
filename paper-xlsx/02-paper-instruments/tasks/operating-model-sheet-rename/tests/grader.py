#!/usr/bin/env python3
"""Standalone standard-library grader runtime copied into each Harbor task."""

from __future__ import annotations

# Replaced by materialize_xlsx_graders.py in each standalone task grader.
TASK = {'allowed_changed_parts': ['xl/charts/chart1.xml',
                           'xl/workbook.xml',
                           'xl/worksheets/sheet1.xml',
                           'xl/worksheets/sheet2.xml'],
 'chart_formulas_from_golden': True,
 'charts_from_golden': True,
 'defined_names_from_golden': True,
 'formula_cache_invalidation': {'refreshed_values': {
     'Operating Model!E2': 397, 'Operating Model!E3': 0.92,
     'Operating Model!E4': 21, 'Dashboard!B1': 397, 'Dashboard!B2': 21}},
 'golden': 'operating-model-golden.xlsx',
 'id': 'operating-model-sheet-rename',
 'inputs': [{'path': 'eval_fixtures/xlsx/operating-model.xlsx',
             'sha256': '1f1007150f8304f532f6292aad34c2a7dae55576fd0657dd65fc3ccc6ae440ef'}],
 'output': 'evals/operating-model-sheet-rename/output.xlsx',
 'part_not_contains': [{'part': 'xl/workbook.xml', 'text': 'Ops Model'}],
 'sheet_cells_from_golden': ['Operating Model', 'Dashboard'],
 'sheet_names_from_golden': True}

import hashlib
import json
import os
import posixpath
import pwd
import re
import signal
import sys
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "chart": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}

CHECK_WEIGHTS = {
    "critical_effect": 4.0,
    "critical_collateral": 3.0,
    "behavioral": 4.0,
    "supporting": 1.0,
}


class HardFailure(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def safe_package(path: Path, *, suffixes=(".xlsx", ".xlsm")) -> dict[str, bytes]:
    if path.suffix.lower() not in suffixes:
        raise HardFailure(f"unexpected workbook extension: {path.suffix}")
    if not path.is_file():
        raise HardFailure(f"missing workbook: {path}")
    if path.stat().st_size > 64 * 1024 * 1024:
        raise HardFailure("workbook exceeds 64 MiB package limit")
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            names = [info.filename for info in infos]
            if len(names) != len(set(names)):
                raise HardFailure("duplicate ZIP member")
            lowered = [name.casefold() for name in names]
            if len(lowered) != len(set(lowered)):
                raise HardFailure("case-colliding ZIP member")
            expanded = 0
            payloads: dict[str, bytes] = {}
            for info in infos:
                pure = PurePosixPath(info.filename)
                if pure.is_absolute() or ".." in pure.parts or "\\" in info.filename:
                    raise HardFailure(f"non-canonical ZIP member: {info.filename}")
                if info.flag_bits & 0x1:
                    raise HardFailure("encrypted ZIP member")
                expanded += info.file_size
                if expanded > 256 * 1024 * 1024:
                    raise HardFailure("expanded package exceeds 256 MiB")
                payloads[info.filename] = archive.read(info)
    except (zipfile.BadZipFile, OSError) as exc:
        raise HardFailure(f"invalid workbook ZIP: {exc}") from exc
    required = {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"}
    missing = sorted(required - payloads.keys())
    if missing:
        raise HardFailure(f"missing package roots: {missing}")
    for name, payload in payloads.items():
        if name.endswith((".xml", ".rels", ".vml")):
            try:
                ET.fromstring(payload)
            except ET.ParseError as exc:
                raise HardFailure(f"invalid XML in {name}: {exc}") from exc
    validate_relationships(payloads)
    return payloads


def rel_target(source_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def relationship_source(rel_name: str) -> str | None:
    if rel_name == "_rels/.rels":
        return ""
    marker = "/_rels/"
    if marker not in rel_name or not rel_name.endswith(".rels"):
        return None
    prefix, leaf = rel_name.split(marker, 1)
    return f"{prefix}/{leaf[:-5]}"


def validate_relationships(payloads: dict[str, bytes]) -> None:
    for name, payload in payloads.items():
        if not name.endswith(".rels"):
            continue
        source = relationship_source(name)
        if source is None:
            continue
        root = ET.fromstring(payload)
        ids: set[str] = set()
        for rel in root:
            rid = rel.attrib.get("Id", "")
            if not rid or rid in ids:
                raise HardFailure(f"invalid relationship id in {name}: {rid!r}")
            ids.add(rid)
            if rel.attrib.get("TargetMode") == "External":
                continue
            target = rel_target(source, rel.attrib.get("Target", ""))
            if target not in payloads:
                raise HardFailure(f"dangling relationship {name}:{rid} -> {target}")


def part_reachability(payloads: dict[str, bytes], part: str) -> dict:
    relationships = []
    for name, payload in payloads.items():
        if not name.endswith(".rels"):
            continue
        source = relationship_source(name)
        if source is None:
            continue
        for relation in ET.fromstring(payload):
            if relation.attrib.get("TargetMode") == "External":
                continue
            target = rel_target(source, relation.attrib.get("Target", ""))
            if target == part:
                relationships.append(
                    {
                        "source": source,
                        "relationship_part": name,
                        "id": relation.attrib.get("Id", ""),
                    }
                )

    content_type_entries = []
    content_types = ET.fromstring(payloads["[Content_Types].xml"])
    for node in content_types:
        if local_name(node.tag) != "Override":
            continue
        if node.attrib.get("PartName", "").lstrip("/") == part:
            content_type_entries.append(dict(sorted(node.attrib.items())))
    return {
        "part_present": part in payloads,
        "relationships": relationships,
        "content_type_entries": content_type_entries,
    }


def xml_signature(payload: bytes):
    def node_signature(node: ET.Element):
        text = node.text or ""
        if (
            not text.strip()
            and node.attrib.get("{http://www.w3.org/XML/1998/namespace}space")
            != "preserve"
        ):
            text = ""
        tail = node.tail or ""
        if not tail.strip():
            tail = ""
        return (
            node.tag,
            tuple(sorted(node.attrib.items())),
            text,
            tail,
            tuple(node_signature(child) for child in node),
        )

    return node_signature(ET.fromstring(payload))


_SIMPLE_SHEET_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")
_QUOTED_SHEET_REFERENCE = re.compile(r"'((?:[^']|'')+)'!")


def canonical_formula(value: str | None) -> str | None:
    """Normalize OOXML spellings that Excel evaluates identically."""

    if value is None:
        return None

    def unquote_simple_sheet(match: re.Match[str]) -> str:
        sheet = match.group(1).replace("''", "'")
        if _SIMPLE_SHEET_NAME.fullmatch(sheet):
            return f"{sheet}!"
        return match.group(0)

    formula = value[1:] if value.startswith("=") else value
    segments: list[str] = []
    start = 0
    index = 0
    in_string = False
    while index < len(formula):
        if formula[index] != '"':
            index += 1
            continue
        if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
            index += 2
            continue
        segment = formula[start:index]
        segments.append(
            segment
            if in_string
            else _QUOTED_SHEET_REFERENCE.sub(unquote_simple_sheet, segment)
        )
        segments.append('"')
        in_string = not in_string
        index += 1
        start = index
    segment = formula[start:]
    segments.append(
        segment
        if in_string
        else _QUOTED_SHEET_REFERENCE.sub(unquote_simple_sheet, segment)
    )
    result = "".join(segments)
    # Only unwrap a bare A1 operand, outside quoted strings/sheet names and
    # structured references. In particular, never unwrap a function's argument
    # list or a grouped arithmetic expression.
    protected = re.compile(r'("(?:[^\"]|\"\")*"|\'(?:[^\']|\'\')*\'|\[[^\]]*\])')
    atom = re.compile(r"(?<![\w.$])\((\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)\)")
    fragments = protected.split(result)
    for index in range(0, len(fragments), 2):
        previous = None
        while previous != fragments[index]:
            previous = fragments[index]
            fragments[index] = atom.sub(r"\1", fragments[index])
    result = "".join(fragments)
    # Parentheses around an entire expression do not change its value. Do not
    # simplify arithmetic, relative references, function calls, or quoted text.
    while result.startswith("(") and result.endswith(")"):
        depth = 0
        quoted = False
        closes_early = False
        for index, character in enumerate(result):
            if character == '"':
                quoted = not quoted
            if not quoted:
                depth += (character == "(") - (character == ")")
                if depth == 0 and index < len(result) - 1:
                    closes_early = True
                    break
        if closes_early or depth != 0 or quoted:
            break
        result = result[1:-1]
    return result


def canonical_chart_formula(value):
    # Chart references identify fixed ranges, not copy-relative cell formulas.
    value = canonical_formula(value)
    if value is None:
        return None
    prefix, separator, reference = value.rpartition("!")
    if separator and re.fullmatch(r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?", reference):
        return prefix + separator + reference.replace("$", "")
    return value


def xml_node_signature(node: ET.Element):
    text = node.text or ""
    if (
        not text.strip()
        and node.attrib.get("{http://www.w3.org/XML/1998/namespace}space") != "preserve"
    ):
        text = ""
    tail = node.tail or ""
    if not tail.strip():
        tail = ""
    return (
        node.tag,
        tuple(sorted(node.attrib.items())),
        text,
        tail,
        tuple(xml_node_signature(child) for child in node),
    )


_BOOLEAN_STYLE_NODES = {
    "b",
    "condense",
    "extend",
    "i",
    "outline",
    "shadow",
    "strike",
}
_UNORDERED_STYLE_CONTAINERS = {"border", "font", "patternFill"}


def style_node_signature(node: ET.Element):
    """Canonicalize visual style semantics without depending on writer syntax."""

    name = local_name(node.tag)
    attributes = dict(node.attrib)
    if name in _BOOLEAN_STYLE_NODES:
        attributes["val"] = "1" if truthy_xml(attributes.get("val", "1")) else "0"
    if name == "patternFill" and attributes.get("patternType") in {None, "none"}:
        attributes["patternType"] = "none"
    if name == "xf":
        attributes = {
            key: value
            for key, value in attributes.items()
            if not key.startswith("apply")
            and not (key in {"pivotButton", "quotePrefix"} and not truthy_xml(value))
        }
    children = [style_node_signature(child) for child in node]
    if name in _UNORDERED_STYLE_CONTAINERS:
        children.sort(key=repr)
    text = node.text or ""
    if not text.strip():
        text = ""
    return (
        node.tag,
        tuple(sorted(attributes.items())),
        text,
        tuple(children),
    )


def _component_signature(root: ET.Element | None, index: int) -> tuple | None:
    if root is None or index < 0 or index >= len(root):
        return None
    return style_node_signature(root[index])


def _integer_attribute(node: ET.Element, name: str) -> int:
    try:
        return int(node.attrib.get(name, "0"))
    except ValueError:
        return 0


def style_signatures(payloads: dict[str, bytes]) -> tuple[tuple, ...]:
    """Resolve cell XF indices to formatting semantics, not table positions."""

    payload = payloads.get("xl/styles.xml")
    if payload is None:
        return ((),)
    root = ET.fromstring(payload)
    fonts = root.find("main:fonts", NS)
    fills = root.find("main:fills", NS)
    borders = root.find("main:borders", NS)
    cell_style_xfs = root.find("main:cellStyleXfs", NS)
    cell_xfs = root.find("main:cellXfs", NS)
    number_formats = {
        node.attrib.get("numFmtId", ""): node.attrib.get("formatCode", "")
        for node in root.findall("main:numFmts/main:numFmt", NS)
    }

    def xf_signature(xf: ET.Element, *, include_base: bool) -> tuple:
        num_fmt_id = str(_integer_attribute(xf, "numFmtId"))
        base = None
        if include_base and cell_style_xfs is not None:
            base_index = _integer_attribute(xf, "xfId")
            if base_index < len(cell_style_xfs):
                base = xf_signature(cell_style_xfs[base_index], include_base=False)
        alignment = next(
            (
                style_node_signature(child)
                for child in xf
                if local_name(child.tag) == "alignment"
            ),
            None,
        )
        protection = next(
            (
                style_node_signature(child)
                for child in xf
                if local_name(child.tag) == "protection"
            ),
            None,
        )
        return (
            base,
            ("custom", number_formats[num_fmt_id]) if num_fmt_id in number_formats else ("builtin", num_fmt_id),
            _component_signature(fonts, _integer_attribute(xf, "fontId")),
            _component_signature(fills, _integer_attribute(xf, "fillId")),
            _component_signature(borders, _integer_attribute(xf, "borderId")),
            alignment,
            protection,
            truthy_xml(xf.attrib.get("quotePrefix")),
            truthy_xml(xf.attrib.get("pivotButton")),
        )

    if cell_xfs is None or not len(cell_xfs):
        return ((),)
    return tuple(xf_signature(xf, include_base=True) for xf in cell_xfs)


def chart_semantic_signature(payload: bytes, *, unspecified_style=False):
    """Retain chart structure while ignoring recalculable cache serialization."""

    cache_nodes = {"numCache", "strCache", "multiLvlStrCache"}

    def node_signature(node: ET.Element):
        if local_name(node.tag) in cache_nodes:
            return None
        if unspecified_style and node.tag == "{" + NS["chart"] + "}style":
            value = node.get("val", "")
            if set(node.attrib) == {"val"} and not len(node) and value.isdigit() and 1 <= int(value) <= 48:
                return None
        text = node.text or ""
        if local_name(node.tag) == "f":
            text = canonical_chart_formula(text) or ""
        elif not text.strip():
            text = ""
        children = tuple(
            item for child in node if (item := node_signature(child)) is not None
        )
        return (
            node.tag,
            tuple(sorted(node.attrib.items())),
            text,
            children,
        )

    return node_signature(ET.fromstring(payload))


def semantic_part_signature(part: str, payload: bytes):
    if part.startswith("xl/charts/") and part.endswith(".xml"):
        return chart_semantic_signature(payload)
    return xml_signature(payload)


def workbook_structure_signature(payload: bytes):
    """Ignore recalculation directives while retaining all workbook structure."""

    def node_signature(node: ET.Element):
        children = tuple(
            node_signature(child) for child in node if local_name(child.tag) != "calcPr"
        )
        text = node.text or ""
        if not text.strip():
            text = ""
        return (
            node.tag,
            tuple(sorted(node.attrib.items())),
            text,
            children,
        )

    return node_signature(ET.fromstring(payload))


def core_preservation_signature(payload: bytes):
    mutable = {"lastModifiedBy", "modified", "revision"}
    root = ET.fromstring(payload)
    return tuple(
        sorted(
            xml_node_signature(child)
            for child in root
            if local_name(child.tag) not in mutable
        )
    )


def unowned_xml_signature(part: str, payload: bytes):
    root = ET.fromstring(payload)
    if part == "xl/styles.xml":
        return style_node_signature(root)
    if part == "docProps/core.xml":
        return core_preservation_signature(payload)
    if part == "[Content_Types].xml" or part.endswith(".rels"):
        return (
            root.tag,
            tuple(sorted(root.attrib.items())),
            tuple(sorted((xml_node_signature(child) for child in root), key=repr)),
        )
    return xml_node_signature(root)


def unowned_part_equal(
    part: str,
    before: bytes | None,
    after: bytes | None,
    *,
    output_model: dict,
    before_package: dict | None = None,
    after_package: dict | None = None,
) -> bool:
    if before == after:
        return True
    if before is None or after is None:
        return False
    if before_package is not None and after_package is not None:
        if part == "[Content_Types].xml":
            return effective_content_types(before_package) == effective_content_types(after_package)
        if part.endswith((".xml", ".rels", ".vml")):
            before = resolved_xml(part, before, before_package)
            after = resolved_xml(part, after, after_package)
    if (
        part == "xl/workbook.xml"
        and recalc_posture(output_model)
        and workbook_structure_signature(before) == workbook_structure_signature(after)
    ):
        return True
    if part == "[Content_Types].xml" or part.endswith((".xml", ".rels", ".vml")):
        try:
            return unowned_xml_signature(part, before) == unowned_xml_signature(
                part, after
            )
        except ET.ParseError:
            return False
    return False


def effective_content_types(payloads):
    root = ET.fromstring(payloads["[Content_Types].xml"])
    defaults = {node.get("Extension"): node.get("ContentType") for node in root if local_name(node.tag) == "Default"}
    overrides = {node.get("PartName", "").lstrip("/"): node.get("ContentType") for node in root if local_name(node.tag) == "Override"}
    return {part: overrides.get(part, defaults.get(part.rsplit(".", 1)[-1])) for part in payloads if part != "[Content_Types].xml"}


def normalize_comment_part_names(reference, package):
    """Match comment parts by their complete incoming ownership, not filename."""
    def owners(payloads):
        result = {}
        for part, payload in payloads.items():
            if not part.endswith(".rels"):
                continue
            source = relationship_source(part)
            for edge in ET.fromstring(payload):
                if edge.get("Type", "").endswith("/comments") and edge.get("TargetMode") != "External":
                    target = rel_target(source, edge.get("Target", ""))
                    result.setdefault(target, []).append((source, edge.get("Type")))
        grouped = {}
        for target, edges in result.items():
            grouped.setdefault(tuple(sorted(edges)), []).append(target)
        return grouped

    before, after = owners(reference), owners(package)
    mapping = {targets[0]: before[owner][0] for owner, targets in after.items()
               if len(targets) == 1 and len(before.get(owner, [])) == 1
               and targets[0] in package and before[owner][0] in reference
               and targets[0] != before[owner][0] and before[owner][0] not in package}
    if not mapping:
        return package
    result = {mapping.get(part, part): payload for part, payload in package.items()}
    for part, payload in list(result.items()):
        if part.endswith(".rels"):
            tree = ET.fromstring(payload)
            source = relationship_source(part)
            changed = False
            for edge in tree:
                target = rel_target(source, edge.get("Target", ""))
                if edge.get("TargetMode") != "External" and target in mapping:
                    edge.set("Target", "/" + mapping[target])
                    changed = True
            if changed:
                result[part] = ET.tostring(tree)
        elif part == "[Content_Types].xml":
            tree = ET.fromstring(payload)
            for node in tree:
                target = node.get("PartName", "").lstrip("/")
                if target in mapping:
                    node.set("PartName", "/" + mapping[target])
            result[part] = ET.tostring(tree)
    return result


def validation_attributes(node):
    attributes = dict(node.attrib)
    attributes.pop("sqref", None)
    for key in ("allowBlank", "showDropDown", "showInputMessage", "showErrorMessage"):
        value = attributes.get(key, "0")
        attributes[key] = {"false": "0", "true": "1"}.get(value, value)
    for key, default in (("type", "none"), ("errorStyle", "stop"), ("imeMode", "noControl"), ("operator", "between")):
        attributes.setdefault(key, default)
    return tuple(sorted(attributes.items()))


def resolved_xml(part, payload, package):
    """Resolve relationship references; keep target ownership and edge types."""
    root = ET.fromstring(payload)
    if part.endswith(".rels"):
        source = relationship_source(part)
        for node in root:
            node.attrib.pop("Id", None)
            if node.get("TargetMode") != "External":
                node.set("Target", rel_target(source, node.get("Target", "")))
                node.attrib.pop("TargetMode", None)
        root[:] = sorted(root, key=lambda node: repr(xml_node_signature(node)))
    else:
        rel_payload = package.get(relationship_part(part))
        relations = {} if rel_payload is None else {node.get("Id"): node for node in ET.fromstring(rel_payload)}
        for node in root.iter():
            for key, value in list(node.attrib.items()):
                if namespace_uri(key) == NS["rel"]:
                    edge = relations.get(value)
                    if edge is None:
                        raise HardFailure(f"unresolved relationship {part}:{value}")
                    external = edge.get("TargetMode") == "External"
                    target = edge.get("Target", "") if external else rel_target(part, edge.get("Target", ""))
                    node.set(key, repr((edge.get("Type"), external, target)))
        if part == "xl/styles.xml":
            formats = {node.get("numFmtId"): node.get("formatCode") for node in root.findall("main:numFmts/main:numFmt", NS)}
            for node in root.iter():
                if node.get("numFmtId") in formats:
                    node.set("numFmtId", "custom:" + formats[node.get("numFmtId")])
    return ET.tostring(root)


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def namespace_uri(tag: str) -> str:
    return tag[1:].split("}", 1)[0] if tag.startswith("{") else ""


def relationship_part(source_part: str) -> str:
    directory, leaf = posixpath.split(source_part)
    return posixpath.join(directory, "_rels", f"{leaf}.rels")


def normalized_refs(value: str | None) -> tuple[str, ...]:
    return tuple(sorted((value or "").split()))


def range_bounds(reference: str) -> tuple[int, int, int, int]:
    def coordinate(address):
        match = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", address)
        if match is None:
            raise HardFailure(f"unsupported cell reference {address!r}")
        column = 0
        for character in match[1].upper():
            column = column * 26 + ord(character) - 64
        return column, int(match[2])
    first, _, last = reference.partition(":")
    left, top = coordinate(first)
    right, bottom = coordinate(last or first)
    return left, top, right, bottom


def ranges_overlap(first: str, second: str) -> bool:
    for a in first.split():
        al, at, ar, ab = range_bounds(a)
        for b in second.split():
            bl, bt, br, bb = range_bounds(b)
            if max(al, bl) <= min(ar, br) and max(at, bt) <= min(ab, bb):
                return True
    return False


def descendant_text(node: ET.Element, name: str) -> str | None:
    for child in node.iter():
        if local_name(child.tag) == name:
            return "".join(child.itertext())
    return None


def rule_model(node: ET.Element, differential_formats: tuple = ()) -> dict:
    formulas = [
        canonical_formula(child.text) or ""
        for child in node.iter()
        if local_name(child.tag) in {"f", "formula", "formula1", "formula2"}
    ]
    # Numeric priority, not XML order, determines evaluation precedence.
    # workbook_model records relative precedence for overlapping ranges below.
    attributes = tuple(
        sorted(
            (key, value)
            for key, value in node.attrib.items()
            if local_name(key) not in {"priority", "dxfId"}
        )
    )
    try:
        dxf_index = int(node.attrib.get("dxfId", "-1"))
    except ValueError:
        dxf_index = -1
    return {
        "namespace": namespace_uri(node.tag),
        "tag": local_name(node.tag),
        "attributes": attributes,
        "differential_format": differential_formats[dxf_index]
        if 0 <= dxf_index < len(differential_formats)
        else None,
        "formulas": formulas,
    }


def core_properties(payloads: dict[str, bytes]) -> dict[str, str]:
    payload = payloads.get("docProps/core.xml")
    if payload is None:
        return {}
    return {
        local_name(node.tag): node.text or ""
        for node in ET.fromstring(payload)
        if node.text is not None
    }


def workbook_model(payloads: dict[str, bytes]) -> dict:
    workbook = ET.fromstring(payloads["xl/workbook.xml"])
    styles = style_signatures(payloads)
    style_root = ET.fromstring(payloads["xl/styles.xml"]) if "xl/styles.xml" in payloads else None
    differential_formats = ()
    if style_root is not None:
        dxfs = style_root.find("main:dxfs", NS)
        if dxfs is not None:
            differential_formats = tuple(style_node_signature(node) for node in dxfs)
    rels = ET.fromstring(payloads["xl/_rels/workbook.xml.rels"])
    rel_map = {
        rel.attrib["Id"]: rel_target("xl/workbook.xml", rel.attrib["Target"])
        for rel in rels
        if rel.attrib.get("TargetMode") != "External"
    }
    shared: list[str] = []
    if "xl/sharedStrings.xml" in payloads:
        root = ET.fromstring(payloads["xl/sharedStrings.xml"])
        for item in root.findall("main:si", NS):
            shared.append(
                "".join(
                    node.text or "" for node in item.iter() if node.tag.endswith("}t")
                )
            )

    sheets: dict[str, dict] = {}
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        title = sheet.attrib["name"]
        rid = sheet.attrib[f"{{{NS['rel']}}}id"]
        part = rel_map[rid]
        cells: dict[str, dict] = {}
        root = ET.fromstring(payloads[part])
        for cell in root.findall(".//main:c", NS):
            address = cell.attrib.get("r")
            if not address:
                continue
            formula = cell.find("main:f", NS)
            value = cell.find("main:v", NS)
            inline = cell.find("main:is", NS)
            if value is None:
                cache_state = "missing"
            elif value.text is None or (
                not value.text.strip()
                and value.attrib.get("{http://www.w3.org/XML/1998/namespace}space")
                != "preserve"
            ):
                cache_state = "empty"
            else:
                cache_state = "value"
            raw = value.text if value is not None else None
            if (
                raw is not None
                and not raw.strip()
                and value.attrib.get("{http://www.w3.org/XML/1998/namespace}space")
                != "preserve"
            ):
                raw = None
            cell_type = cell.attrib.get("t")
            if cell_type == "s" and raw is not None:
                try:
                    parsed = shared[int(raw)]
                except (ValueError, IndexError):
                    parsed = raw
            elif cell_type == "inlineStr" and inline is not None:
                parsed = "".join(
                    node.text or "" for node in inline.iter() if node.tag.endswith("}t")
                )
            elif cell_type == "b" and raw is not None:
                parsed = raw == "1"
            else:
                parsed = raw
                if raw is not None and formula is None:
                    try:
                        parsed = float(raw) if "." in raw else int(raw)
                    except ValueError:
                        pass
            cells[address] = {
                "value": parsed,
                "formula": canonical_formula(formula.text)
                if formula is not None
                else None,
                "formula_type": formula.attrib.get("t")
                if formula is not None
                else None,
                "formula_ref": formula.attrib.get("ref")
                if formula is not None
                else None,
                "formula_attributes": tuple(sorted(formula.attrib.items()))
                if formula is not None
                else (),
                "cache_state": cache_state if formula is not None else None,
                "style": styles[_integer_attribute(cell, "s")]
                if _integer_attribute(cell, "s") < len(styles)
                else None,
                "style_id": cell.attrib.get("s"),
                "type": cell_type,
                "empty_unstyled": formula is None and parsed is None and inline is None
                and cell_type in {None, "n"} and _integer_attribute(cell, "s") == 0
                and not (set(cell.attrib) - {"r", "s", "t"})
                and all(child.tag == "{" + NS["main"] + "}v" and not child.attrib
                        and not (child.text or "").strip() for child in cell),
            }
        merged_cells = sorted(
            node.attrib.get("ref", "")
            for node in root.findall("main:mergeCells/main:mergeCell", NS)
        )
        conditional_formatting = []
        priority_rules = []
        for node in root.iter():
            if local_name(node.tag) != "conditionalFormatting":
                continue
            sqref = node.attrib.get("sqref") or descendant_text(node, "sqref")
            conditional_formatting.append(
                {
                    "namespace": namespace_uri(node.tag),
                    "sqref": normalized_refs(sqref),
                    "rules": [
                        rule_model(child, differential_formats)
                        for child in node
                        if local_name(child.tag) == "cfRule"
                    ],
                }
            )
            for child, model in zip(
                (child for child in node if local_name(child.tag) == "cfRule"),
                conditional_formatting[-1]["rules"],
            ):
                priority_rules.append((sqref or "", child.attrib.get("priority"), model))
        for ranges, priority, model in priority_rules:
            model["precedes"] = tuple(sorted(
                repr({key: value for key, value in other.items() if key != "precedes"})
                + "@" + " ".join(normalized_refs(other_ranges))
                for other_ranges, other_priority, other in priority_rules
                if other is not model and ranges_overlap(ranges, other_ranges)
                and priority is not None and other_priority is not None
                and int(priority) < int(other_priority)
            ))
        for group in conditional_formatting:
            group["rules"].sort(key=repr)
        conditional_formatting.sort(key=lambda item: (item["namespace"], item["sqref"]))

        data_validations = []
        for node in root.iter():
            if local_name(node.tag) != "dataValidation":
                continue
            sqref = node.attrib.get("sqref") or descendant_text(node, "sqref")
            data_validations.append(
                {
                    "namespace": namespace_uri(node.tag),
                    "sqref": normalized_refs(sqref),
                    "attributes": validation_attributes(node),
                    "formula1": canonical_formula(descendant_text(node, "formula1")),
                    "formula2": canonical_formula(descendant_text(node, "formula2")),
                }
            )
        data_validations.sort(key=lambda item: (item["namespace"], item["sqref"]))

        sparkline_groups = []
        for group in root.iter():
            if local_name(group.tag) != "sparklineGroup":
                continue
            colors = []
            sparklines = []
            for child in group:
                child_name = local_name(child.tag)
                if child_name.startswith("color"):
                    colors.append(
                        {
                            "tag": child_name,
                            "attributes": tuple(sorted(child.attrib.items())),
                        }
                    )
                elif child_name == "sparklines":
                    for sparkline in child:
                        if local_name(sparkline.tag) != "sparkline":
                            continue
                        sparklines.append(
                            {
                                "formula": canonical_formula(
                                    descendant_text(sparkline, "f")
                                ),
                                "sqref": normalized_refs(
                                    descendant_text(sparkline, "sqref")
                                ),
                            }
                        )
            sparkline_groups.append(
                {
                    "namespace": namespace_uri(group.tag),
                    "attributes": tuple(sorted(group.attrib.items())),
                    "colors": tuple(colors),
                    "sparklines": tuple(sparklines),
                }
            )

        auto_filters = [
            {
                "ref": node.attrib.get("ref", ""),
                "signature": tuple(rule_model(child) for child in node),
            }
            for node in root
            if local_name(node.tag) == "autoFilter"
        ]
        sheet_protection_node = next(
            (node for node in root if local_name(node.tag) == "sheetProtection"),
            None,
        )
        sheet_protection = (
            None
            if sheet_protection_node is None
            else tuple(sorted(sheet_protection_node.attrib.items()))
        )
        row_breaks = []
        for breaks in root:
            if local_name(breaks.tag) != "rowBreaks":
                continue
            row_breaks.extend(
                tuple(sorted(child.attrib.items()))
                for child in breaks
                if local_name(child.tag) == "brk"
            )
        x14_regions = []
        for node in root.iter():
            if namespace_uri(node.tag) != "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main":
                continue
            refs = tuple(
                sorted(
                    (local_name(key), value)
                    for key, value in node.attrib.items()
                    if local_name(key) in {"ref", "sqref"}
                )
            )
            x14_regions.append((local_name(node.tag), refs, (node.text or "").strip()))
        sheet_format = root.find("main:sheetFormatPr", NS)
        default_height = float(sheet_format.get("defaultRowHeight", "15")) if sheet_format is not None else 15.0
        row_heights = {0: (default_height, ())}
        for node in root.findall("main:sheetData/main:row", NS):
            height = float(node.get("ht", str(default_height)))
            attributes = {key: value for key, value in node.attrib.items() if key not in {"r", "ht", "spans"}}
            if height == default_height:
                attributes.pop("customHeight", None)
            for key in ("hidden", "collapsed", "thickTop", "thickBot", "customFormat", "customHeight"):
                if attributes.get(key) in {"0", "false"}:
                    attributes.pop(key)
                elif attributes.get(key) == "true":
                    attributes[key] = "1"
            if attributes.get("outlineLevel") == "0":
                attributes.pop("outlineLevel")
            if height != default_height or attributes:
                row_heights[int(node.attrib["r"])] = (height, tuple(sorted(attributes.items())))
        column_widths = {}
        for node in root.findall("main:cols/main:col", NS):
            if "min" not in node.attrib or "max" not in node.attrib:
                continue
            for index in range(int(node.attrib["min"]), int(node.attrib["max"]) + 1):
                column_widths[index] = node.attrib.get("width")

        sheet_rels: dict[str, tuple[str, str]] = {}
        external_rels: dict[str, tuple[str, str]] = {}
        rels_name = relationship_part(part)
        if rels_name in payloads:
            for relation in ET.fromstring(payloads[rels_name]):
                if relation.attrib.get("TargetMode") == "External":
                    external_rels[relation.attrib.get("Id", "")] = (
                        relation.attrib.get("Type", ""),
                        relation.attrib.get("Target", ""),
                    )
                    continue
                sheet_rels[relation.attrib.get("Id", "")] = (
                    relation.attrib.get("Type", ""),
                    rel_target(part, relation.attrib.get("Target", "")),
                )

        for hyperlink in root.findall("main:hyperlinks/main:hyperlink", NS):
            address = hyperlink.attrib.get("ref", "")
            if address not in cells:
                continue
            rid = hyperlink.attrib.get(f"{{{NS['rel']}}}id")
            if rid and rid in external_rels:
                cells[address]["hyperlink"] = external_rels[rid][1]
            else:
                cells[address]["hyperlink"] = hyperlink.attrib.get("location")

        comments = []
        vml_drawings = []
        for relation_type, target in sheet_rels.values():
            if relation_type.endswith("/comments") and target in payloads:
                comments_root = ET.fromstring(payloads[target])
                authors = [
                    node.text or ""
                    for node in comments_root.findall("main:authors/main:author", NS)
                ]
                for comment in comments_root.findall(
                    "main:commentList/main:comment", NS
                ):
                    author_id = comment.attrib.get("authorId", "")
                    try:
                        index = int(author_id)
                        if not 0 <= index < len(authors):
                            raise IndexError
                        author = authors[index]
                    except (ValueError, IndexError):
                        author = ("invalid-author-index", author_id)
                    comments.append(
                        {
                            "part": target,
                            "ref": comment.attrib.get("ref", ""),
                            "author": author,
                            "text": "".join(
                                child.text or ""
                                for child in comment.iter()
                                if local_name(child.tag) == "t"
                            ),
                        }
                    )
            elif relation_type.endswith("/vmlDrawing") and target in payloads:
                vml_drawings.append(
                    {"part": target, "signature": xml_signature(payloads[target])}
                )
        comments.sort(key=lambda item: (item["ref"], repr(item["author"]), item["text"]))
        vml_drawings.sort(key=lambda item: item["part"])

        sheets[title] = {
            "part": part,
            "state": sheet.attrib.get("state", "visible"),
            "cells": cells,
            "merged_cells": merged_cells,
            "conditional_formatting": conditional_formatting,
            "data_validations": data_validations,
            "sparkline_groups": sparkline_groups,
            "auto_filters": auto_filters,
            "sheet_protection": sheet_protection,
            "row_breaks": row_breaks,
            "x14_regions": x14_regions,
            "row_heights": row_heights,
            "hidden_rows": (
                truthy_xml(sheet_format.get("zeroHeight")) if sheet_format is not None else False,
                tuple(sorted(int(node.attrib["r"]) for node in root.findall("main:sheetData/main:row", NS)
                             if truthy_xml(node.get("hidden")) or float(node.get("ht", str(default_height))) == 0)),
            ),
            "column_widths": column_widths,
            "comments": comments,
            "vml_drawings": vml_drawings,
            "xml": root,
        }

    names = {
        node.attrib.get("name", ""): canonical_formula(node.text) or ""
        for node in workbook.findall("main:definedNames/main:definedName", NS)
    }
    table_refs: dict[str, str] = {}
    tables: dict[str, dict] = {}
    for name, payload in payloads.items():
        if name.startswith("xl/tables/") and name.endswith(".xml"):
            root = ET.fromstring(payload)
            table_name = root.attrib.get("name", name)
            table_refs[table_name] = root.attrib.get("ref", "")
            auto_filter = next(
                (node for node in root if local_name(node.tag) == "autoFilter"),
                None,
            )
            columns = []
            for column in root.iter():
                if local_name(column.tag) != "tableColumn":
                    continue
                columns.append(
                    {
                        "attributes": tuple(sorted(column.attrib.items())),
                        "calculated_formula": canonical_formula(
                            descendant_text(column, "calculatedColumnFormula")
                        ),
                        "totals_formula": canonical_formula(
                            descendant_text(column, "totalsRowFormula")
                        ),
                    }
                )
            style = next(
                (node for node in root if local_name(node.tag) == "tableStyleInfo"),
                None,
            )
            tables[table_name] = {
                "attributes": tuple(sorted(root.attrib.items())),
                "auto_filter": None
                if auto_filter is None
                else {
                    "ref": auto_filter.attrib.get("ref", ""),
                    "rules": tuple(rule_model(child) for child in auto_filter),
                },
                "columns": columns,
                "style": None if style is None else tuple(sorted(style.attrib.items())),
            }
    chart_formulas: dict[str, list[str]] = {}
    charts: dict[str, tuple] = {}
    for name, payload in payloads.items():
        if name.startswith("xl/charts/") and name.endswith(".xml"):
            root = ET.fromstring(payload)
            chart_formulas[name] = [
                canonical_chart_formula(node.text) or ""
                for node in root.findall(".//chart:f", NS)
            ]
            charts[name] = chart_semantic_signature(payload)
    calc_pr = workbook.find("main:calcPr", NS)
    return {
        "sheet_names": list(sheets),
        "sheets": sheets,
        "defined_names": names,
        "table_refs": table_refs,
        "tables": tables,
        "chart_formulas": chart_formulas,
        "charts": charts,
        "calc_properties": {}
        if calc_pr is None
        else dict(sorted(calc_pr.attrib.items())),
        "core_properties": core_properties(payloads),
    }


def empty_workbook_model() -> dict:
    return {
        "sheet_names": [],
        "sheets": {},
        "defined_names": {},
        "table_refs": {},
        "tables": {},
        "chart_formulas": {},
        "charts": {},
        "calc_properties": {},
        "core_properties": {},
    }


def chart_series(root: ET.Element) -> list[ET.Element]:
    return [node for node in root.iter() if local_name(node.tag) == "ser"]


def chart_reference_formula(series: ET.Element, component: str = "val") -> str | None:
    for node in series.iter():
        if local_name(node.tag) != component:
            continue
        for descendant in node.iter():
            if local_name(descendant.tag) == "f":
                return canonical_chart_formula(descendant.text)
    return None


def chart_values_formula(series: ET.Element) -> str | None:
    return chart_reference_formula(series)


def chart_values_has_cache(series: ET.Element) -> bool:
    for node in series.iter():
        if local_name(node.tag) == "val":
            return any(local_name(descendant.tag).endswith("Cache") for descendant in node.iter())
    return False


def chart_cache_is_fresh(series: ET.Element, values: list, component: str = "val") -> bool:
    caches = [(parent, child) for node in series.iter() if local_name(node.tag) == component
              for parent in node.iter() for child in parent if local_name(child.tag) in {"numCache", "strCache"}]
    if len(caches) != 1:
        return False
    parent, cache = caches[0]
    if (local_name(parent.tag), local_name(cache.tag)) not in {("numRef", "numCache"), ("strRef", "strCache")}:
        return False
    points = [node for node in cache if local_name(node.tag) == "pt"]
    count = next((node.get("val") for node in cache if local_name(node.tag) == "ptCount"), None)
    if count is None or int(count) != len(values) or len(points) != len(values):
        return False
    actual = {int(node.get("idx", "-1")): descendant_text(node, "v") for node in points}
    if local_name(cache.tag) == "numCache":
        try:
            for value in actual.values():
                float(value)
        except (TypeError, ValueError):
            return False
    return set(actual) == set(range(len(values))) and all(equivalent_cached_value(actual[index], value) for index, value in enumerate(values))


def clear_chart_values_cache(series: ET.Element, component: str = "val") -> None:
    for node in series.iter():
        if local_name(node.tag) == component:
            for parent in node.iter():
                for child in list(parent):
                    if local_name(child.tag) in {"numCache", "strCache"}:
                        parent.remove(child)


def normalize_category_reference(series: ET.Element) -> None:
    # Both reference encodings can point at the same worksheet labels. A fresh
    # string cache uses strRef; a cache-free writer may retain numRef.
    for node in series.iter():
        if local_name(node.tag) == "cat":
            for child in node:
                if local_name(child.tag) == "numRef":
                    child.tag = "{" + NS["chart"] + "}strRef"








def record(
    checks: list[dict],
    name: str,
    passed: bool,
    detail: str = "",
    *,
    kind: str = "supporting",
    score: float | None = None,
    effect_gate: bool = False,
) -> None:
    if kind not in CHECK_WEIGHTS:
        raise ValueError(f"unknown check kind: {kind}")
    normalized_score = float(bool(passed)) if score is None else float(score)
    if not 0.0 <= normalized_score <= 1.0:
        raise ValueError(f"check score must be between zero and one: {score}")
    checks.append(
        {
            "name": name,
            "passed": normalized_score == 1.0,
            "score": normalized_score,
            "detail": detail,
            "kind": kind,
            "effect_gate": effect_gate,
        }
    )


def recalc_posture(model: dict) -> bool:
    properties = model.get("calc_properties", {})
    return properties.get("calcMode", "auto") in {"auto", "autoNoTable"} and any(
        truthy_xml(properties.get(name)) for name in ("fullCalcOnLoad", "forceFullCalc")
    )


def normalized_cell_type(cell: dict) -> str:
    cell_type = cell.get("type")
    if cell_type in {None, "n"}:
        return "numeric"
    if cell_type in {"inlineStr", "s", "str"}:
        return "string"
    return cell_type


def normalized_formula_type(cell: dict) -> str:
    return (
        "normal"
        if cell.get("formula_type") in {None, "normal"}
        else cell["formula_type"]
    )


def cells_semantically_equal(
    actual: dict | None,
    expected: dict | None,
    *,
    actual_model: dict,
) -> bool:
    if actual is None or expected is None:
        present = actual if actual is not None else expected
        return present is None or (present.get("empty_unstyled", False) and not present.get("hyperlink"))
    if normalized_cell_type(actual) != normalized_cell_type(expected):
        return False
    if actual.get("style") != expected.get("style"):
        return False
    if actual.get("hyperlink") != expected.get("hyperlink"):
        return False

    actual_formula = canonical_formula(actual.get("formula"))
    expected_formula = canonical_formula(expected.get("formula"))
    if actual_formula != expected_formula:
        return False
    actual_has_formula = actual_formula is not None or actual.get("formula_type") is not None
    expected_has_formula = expected_formula is not None or expected.get("formula_type") is not None
    if actual_has_formula != expected_has_formula:
        return False
    if not actual_has_formula:
        return equivalent_cached_value(actual.get("value"), expected.get("value"))
    if normalized_formula_type(actual) != normalized_formula_type(expected):
        return False
    if actual.get("formula_attributes", ()) != expected.get("formula_attributes", ()):
        return False
    if canonical_formula(actual.get("formula_ref")) != canonical_formula(
        expected.get("formula_ref")
    ):
        return False

    actual_cache = actual.get("cache_state")
    expected_cache = expected.get("cache_state")
    if actual_cache in {"missing", "empty"} and expected_cache in {"missing", "empty"}:
        return True
    if equivalent_cached_value(actual.get("value"), expected.get("value")):
        return True
    return actual_cache in {"missing", "empty"} and recalc_posture(actual_model)


def compare_cell(
    checks,
    output_model,
    golden_model,
    primary_model,
    qualified: str,
) -> None:
    title, address = qualified.rsplit("!", 1)
    actual = output_model["sheets"].get(title, {}).get("cells", {}).get(address)
    expected = golden_model["sheets"].get(title, {}).get("cells", {}).get(address)
    before = primary_model.get("sheets", {}).get(title, {}).get("cells", {}).get(address)
    kind = (
        "critical_effect"
        if not cells_semantically_equal(expected, before, actual_model=golden_model)
        else "critical_collateral"
    )
    record(
        checks,
        f"cell {qualified}",
        cells_semantically_equal(actual, expected, actual_model=output_model),
        f"expected={expected!r}; actual={actual!r}",
        kind=kind,
        effect_gate=kind == "critical_effect",
    )


def compare_sheet_cells(
    checks,
    output_model,
    golden_model,
    primary_model,
    title: str,
) -> None:
    actual_cells = output_model["sheets"].get(title, {}).get("cells", {})
    expected_cells = golden_model["sheets"].get(title, {}).get("cells", {})
    primary_cells = primary_model.get("sheets", {}).get(title, {}).get("cells", {})
    addresses = sorted(set(actual_cells) | set(expected_cells) | set(primary_cells))
    effect_addresses: list[str] = []
    collateral_addresses: list[str] = []
    correct_effects = 0
    correct_collateral = 0
    for address in addresses:
        expected = expected_cells.get(address)
        primary = primary_cells.get(address)
        actual = actual_cells.get(address)
        changed_by_oracle = not cells_semantically_equal(
            expected,
            primary,
            actual_model=golden_model,
        )
        correct = cells_semantically_equal(
            actual,
            expected,
            actual_model=output_model,
        )
        if changed_by_oracle:
            effect_addresses.append(address)
            correct_effects += int(correct)
        else:
            collateral_addresses.append(address)
            correct_collateral += int(correct)

    if effect_addresses:
        score = correct_effects / len(effect_addresses)
        record(
            checks,
            f"required cell effects on {title}",
            score == 1.0,
            f"correct={correct_effects}/{len(effect_addresses)}; addresses={effect_addresses}",
            kind="critical_effect",
            score=score,
            effect_gate=True,
        )
    if collateral_addresses:
        score = correct_collateral / len(collateral_addresses)
        record(
            checks,
            f"unowned cell preservation on {title}",
            score == 1.0,
            f"correct={correct_collateral}/{len(collateral_addresses)}",
            kind="critical_collateral",
            score=score,
        )


def compare_sheet_feature(
    checks,
    output_model,
    reference_model,
    title: str,
    feature: str,
    *,
    kind: str,
    effect_gate: bool = False,
) -> None:
    actual = output_model["sheets"].get(title, {}).get(feature)
    expected = reference_model["sheets"].get(title, {}).get(feature)
    record(
        checks,
        f"{title} {feature.replace('_', ' ')}",
        actual == expected,
        f"expected={expected!r}; actual={actual!r}",
        kind=kind,
        effect_gate=effect_gate,
    )


def formula_topology(model: dict, title: str) -> dict:
    cells = model["sheets"].get(title, {}).get("cells", {})
    return {
        address: {
            "formula": cell["formula"],
            "formula_type": normalized_formula_type(cell),
            "formula_ref": canonical_formula(cell["formula_ref"]),
            "formula_attributes": cell.get("formula_attributes", ()),
        }
        for address, cell in cells.items()
        if cell["formula"] is not None
    }


def truthy_xml(value: str | None) -> bool:
    return str(value or "").lower() in {"1", "true", "on"}


def equivalent_cached_value(actual, expected) -> bool:
    try:
        return float(actual) == float(expected)
    except (TypeError, ValueError):
        return actual == expected


# Default indexed palette, ECMA-376 18.8.27. Workbooks may override these in
# styles.xml; indices 64/65 are the system foreground/background colors.
INDEXED_COLORS = (
    "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
    "000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF "
    "800000 008000 000080 808000 800080 008080 C0C0C0 808080 "
    "9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF "
    "000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF "
    "00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 "
    "3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 "
    "003366 339966 003300 333300 993300 993366 333399 333333"
).split()


def numeric_format_section(format_code: str, value: float) -> str:
    """Select the number-format section that actually displays this value."""
    sections = []
    start = 0
    quoted = False
    index = 0
    while index < len(format_code):
        character = format_code[index]
        if character == "\\":
            index += 2
            continue
        if character == '"':
            quoted = not quoted
        elif character == ";" and not quoted:
            sections.append(format_code[start:index])
            start = index + 1
        index += 1
    sections.append(format_code[start:])
    conditions = [re.search(r"\[(<=|>=|<>|=|<|>)(-?\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?)\]", section) for section in sections[:2]]
    if any(conditions):
        for section, condition in zip(sections[:2], conditions):
            if condition is None:
                return section
            operator, threshold = condition[1], float(condition[2])
            if {"<": value < threshold, ">": value > threshold, "<=": value <= threshold,
                ">=": value >= threshold, "=": value == threshold, "<>": value != threshold}[operator]:
                return section
        return sections[2] if len(sections) > 2 else ""
    selected = 1 if value < 0 and len(sections) > 1 else 2 if value == 0 and len(sections) > 2 else 0
    return sections[selected]


def grade_creation_styles(checks, payloads, model, config):
    styles = ET.fromstring(payloads.get("xl/styles.xml", b"<styleSheet/>"))
    xfs = styles.find("main:cellXfs", NS)
    fonts = styles.find("main:fonts", NS)
    fills = styles.find("main:fills", NS)
    custom_formats = {node.get("numFmtId"): node.get("formatCode", "")
                      for node in styles.findall("main:numFmts/main:numFmt", NS)}
    builtins = {9: "0%", 10: "0.00%", 5: '"$"#,##0', 6: '"$"#,##0',
                7: '"$"#,##0.00', 8: '"$"#,##0.00', 44: '_("$"* #,##0.00_)'}
    indexed_colors = [node.get("rgb", "")[-6:] for node in styles.findall("main:colors/main:indexedColors/main:rgbColor", NS)] or INDEXED_COLORS
    theme_colors = []
    if "xl/theme/theme1.xml" in payloads:
        for node in ET.fromstring(payloads["xl/theme/theme1.xml"]).iter():
            if local_name(node.tag) == "clrScheme":
                theme_colors = [next(iter(child)).get("lastClr") or next(iter(child)).get("val") for child in node]
                # Spreadsheet theme indices use light1, dark1, light2, dark2.
                if len(theme_colors) >= 4:
                    theme_colors[:4] = [theme_colors[1], theme_colors[0], theme_colors[3], theme_colors[2]]
                break

    def color(node, default):
        if node is None:
            return default
        value = node.get("rgb")
        if value:
            value = value[-6:]
        elif node.get("theme") is not None:
            index = int(node.get("theme"))
            value = theme_colors[index] if index < len(theme_colors) else None
        elif node.get("indexed") is not None:
            index = int(node.get("indexed"))
            value = indexed_colors[index] if 0 <= index < min(len(indexed_colors), 64) else {64: "000000", 65: "FFFFFF"}.get(index)
        else:
            return default if truthy_xml(node.get("auto")) else None
        if not value or not re.fullmatch("[A-Fa-f0-9]{6}", value):
            return None
        channels = [int(value[index:index + 2], 16) / 255 for index in (0, 2, 4)]
        tint = float(node.get("tint", "0"))
        return [channel * (1 + tint) if tint < 0 else channel + (1 - channel) * tint for channel in channels]

    def cell_style(qualified):
        title, address = qualified.rsplit("!", 1)
        cell = model["sheets"].get(title, {}).get("cells", {}).get(address)
        if cell is None or xfs is None:
            return None
        index = int(cell.get("style_id") or 0)
        return xfs[index] if index < len(xfs) else None

    for qualified in config.get("headers", []):
        xf = cell_style(qualified)
        valid = False
        if xf is not None and fonts is not None and fills is not None:
            font = fonts[_integer_attribute(xf, "fontId")]
            fill = fills[_integer_attribute(xf, "fillId")]
            bold = font.find("main:b", NS)
            pattern = fill.find("main:patternFill", NS)
            foreground = color(font.find("main:color", NS), [0, 0, 0])
            background = color(pattern.find("main:fgColor", NS), [1, 1, 1]) if pattern is not None and pattern.get("patternType") == "solid" else [1, 1, 1]
            def luminance(channels):
                linear = [value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4 for value in channels]
                return sum(value * weight for value, weight in zip(linear, (0.2126, 0.7152, 0.0722)))
            if foreground is not None and background is not None:
                levels = sorted((luminance(foreground), luminance(background)))
                valid = bold is not None and truthy_xml(bold.get("val", "1")) and (levels[1] + .05) / (levels[0] + .05) >= 3
        record(checks, f"readable emphasized header {qualified}", valid, kind="critical_effect", effect_gate=True)
    for key, marker in (("currency_cells", "$"), ("percent_cells", "%")):
        for qualified in config.get(key, []):
            xf = cell_style(qualified)
            format_code = "" if xf is None else custom_formats.get(xf.get("numFmtId"), builtins.get(_integer_attribute(xf, "numFmtId"), ""))
            title, address = qualified.rsplit("!", 1)
            value = model["sheets"].get(title, {}).get("cells", {}).get(address, {}).get("value")
            # Fresh formulas may not yet have a cache. These report formulas
            # calculate positive gross profit, so use the positive section.
            try:
                numeric_value = float(value)
            except (TypeError, ValueError):
                numeric_value = 1.0
            format_code = numeric_format_section(format_code, numeric_value)
            # Currency symbols may be quoted or locale-prefixed. Percent must
            # be an active formatting token, not a quoted display label.
            active = re.sub(r'"[^\"]*"|\\.|\[[^\]]*\]|[_*].', "", format_code)
            valid = bool(re.search(r"[0#]", active)) and (marker in format_code if marker == "$" else marker in active)
            if marker == "$":
                valid = valid and "%" not in active
            record(checks, f"{key.replace('_cells', '')} format {qualified}", valid, kind="critical_effect", effect_gate=True)


def grade_image_replacement(checks, before, after, config, root):
    """Resolve pictures by their drawing anchor, never by a media filename."""
    part = config["drawing_part"]
    relations_part = relationship_part(part)
    left, top, _, _ = range_bounds(config["anchor"])
    replacement = (root / config["replacement_path"]).read_bytes()

    def drawing(payloads, *, expected=False):
        if part not in payloads or relations_part not in payloads:
            raise HardFailure("image replacement lost drawing or relationships")
        relations = {node.get("Id"): node for node in ET.fromstring(payloads[relations_part])}
        drawing_root = ET.fromstring(payloads[part])
        selected_parts = []
        selected_ids = []
        for anchor in drawing_root:
            origin = next((node for node in anchor if local_name(node.tag) == "from"), None)
            selected = origin is not None and descendant_text(origin, "col") == str(left - 1) and descendant_text(origin, "row") == str(top - 1)
            pictures = [node for node in anchor.iter() if local_name(node.tag) == "blip"]
            if selected and len(pictures) != 1:
                raise HardFailure("target image anchor is missing or ambiguous")
            for node in anchor.iter():
                for key, value in list(node.attrib.items()):
                    if namespace_uri(key) != NS["rel"]:
                        continue
                    relationship = relations.get(value)
                    if relationship is None:
                        raise HardFailure(f"unresolved drawing relationship {value}")
                    target = rel_target(part, relationship.get("Target", ""))
                    if relationship.get("TargetMode") == "External":
                        resolved = relationship.get("Target", "")
                    elif relationship.get("Type", "").endswith("/image"):
                        content = payloads[target]
                        if selected and local_name(node.tag) == "blip":
                            selected_parts.append(target)
                            selected_ids.append(value)
                            if expected:
                                content = replacement
                        resolved = hashlib.sha256(content).hexdigest()
                    else:
                        resolved = target
                    node.set(key, repr((relationship.get("Type"), relationship.get("TargetMode"), resolved)))
        if len(selected_parts) != 1:
            raise HardFailure("target image anchor is missing or ambiguous")
        used_ids = {value for node in ET.fromstring(payloads[part]).iter()
                    for key, value in node.attrib.items() if namespace_uri(key) == NS["rel"]}
        relation_signatures = []
        for rid, relationship in relations.items():
            external = relationship.get("TargetMode") == "External"
            target = relationship.get("Target", "") if external else rel_target(part, relationship.get("Target", ""))
            if not external and relationship.get("Type", "").endswith("/image"):
                content = replacement if expected and rid in selected_ids else payloads[target]
                target = hashlib.sha256(content).hexdigest()
                if rid in used_ids:
                    continue  # Each used image is already checked at its anchor.
            relation_signatures.append((relationship.get("Type", ""), "External" if external else "Internal", target))
        return drawing_root, selected_parts[0], sorted(relation_signatures)

    expected_drawing, old_part, expected_relationships = drawing(before, expected=True)
    actual_drawing, new_part, actual_relationships = drawing(after)
    record(checks, "replacement image at requested anchor", after[new_part] == replacement,
           kind="critical_effect", effect_gate=True)
    record(checks, "drawing geometry and other objects preserved",
           xml_node_signature(actual_drawing) == xml_node_signature(expected_drawing),
           kind="critical_collateral")
    media_before = {name for name in before if name.startswith("xl/media/")}
    media_after = {name for name in after if name.startswith("xl/media/")}
    record(checks, "original media retained", before[old_part] in [after[name] for name in media_after], kind="critical_collateral")
    referenced_media = set()
    for rel_part, payload in after.items():
        if rel_part.endswith(".rels"):
            source = relationship_source(rel_part)
            referenced_media.update(rel_target(source, edge.get("Target", "")) for edge in ET.fromstring(payload)
                                    if edge.get("TargetMode") != "External" and edge.get("Type", "").endswith("/image"))
    record(checks, "media inventory preserved with one replacement",
           {hashlib.sha256(before[name]).hexdigest() for name in media_before}
           == {hashlib.sha256(after[name]).hexdigest() for name in media_after if name != new_part}
           and (media_after - media_before) <= referenced_media,
           kind="critical_collateral")
    record(checks, "unrelated drawing relationships preserved", actual_relationships == expected_relationships,
           kind="critical_collateral")
    content_root = ET.fromstring(after["[Content_Types].xml"])
    declared_type = next((node.get("ContentType") for node in content_root
                          if node.get("PartName", "").lstrip("/") == new_part), None)
    if declared_type is None:
        declared_type = next((node.get("ContentType") for node in content_root
                              if node.get("Extension", "").lower() == new_part.rsplit(".", 1)[-1].lower()), None)
    expected_type = ("image/png" if replacement.startswith(b"\x89PNG\r\n\x1a\n") else
                     "image/jpeg" if replacement.startswith(b"\xff\xd8\xff") else
                     "image/gif" if replacement.startswith((b"GIF87a", b"GIF89a")) else None)
    record(checks, "replacement media has correct content type", expected_type is not None and declared_type == expected_type,
           kind="critical_collateral")
    # Media names/deduplication are storage choices. Ownership and geometry are
    # checked above, and unreferenced new media are never exempted.
    def content_types(payloads, ignored):
        return set((hashlib.sha256(payloads[name]).hexdigest() if name.startswith("xl/media/") else name, kind)
                   for name, kind in effective_content_types(payloads).items() if name not in ignored)
    record(checks, "unrelated content types preserved",
           content_types(before, set()) == content_types(after, {new_part} if new_part not in before else set()),
           kind="critical_collateral")
    return {part, relations_part, "[Content_Types].xml"} | media_before | media_after


def pivot_records(payloads, config):
    """Decode the small typed pivot-record matrix, including shared-item indices."""
    definition = ET.fromstring(payloads[config["cache_part"]])
    fields = definition.findall("main:cacheFields/main:cacheField", NS)
    root = ET.fromstring(payloads[config["records_part"]])
    if int(root.get("count", "-1")) != len(root):
        return None
    result = []
    for row in root:
        if local_name(row.tag) != "r" or len(row) != len(fields):
            return None
        values = []
        for index, node in enumerate(row):
            if local_name(node.tag) == "x":
                shared = fields[index].find("main:sharedItems", NS)
                try:
                    shared_index = int(node.get("v"))
                    if shared is None or not 0 <= shared_index < len(shared):
                        return None
                    node = shared[shared_index]
                except (TypeError, ValueError, IndexError):
                    return None
            kind = local_name(node.tag)
            if kind not in {"n", "s", "b", "m", "e", "d"}:
                return None
            value = node.get("v")
            if kind == "n":
                try:
                    value = float(value)
                except (TypeError, ValueError):
                    return None
            values.append(value)
        result.append(values)
    return result


def pivot_definition_signature(part, payload, package, *, refreshed=False):
    root = ET.fromstring(resolved_xml(part, payload, package))
    root.attrib.pop("refreshOnLoad", None)
    if refreshed:
        for key in ("refreshedDate", "refreshedDateIso", "refreshedBy"):
            root.attrib.pop(key, None)
        # Pivot item indices also address shared dictionaries. Preserve those
        # dictionaries even when records were consistently reindexed; otherwise
        # unchanged pivot item ordering would silently acquire new meanings.
        indexed_fields = set()
        for name, content in package.items():
            if name.startswith("xl/pivotTables/") and name.endswith(".xml"):
                table = ET.fromstring(content)
                for index, field in enumerate(table.findall("main:pivotFields/main:pivotField", NS)):
                    if field.find("main:items", NS) is not None:
                        indexed_fields.add(index)
        for index, field in enumerate(root.findall("main:cacheFields/main:cacheField", NS)):
            if index in indexed_fields or field.find("main:fieldGroup", NS) is not None:
                continue
            for child in list(field):
                if local_name(child.tag) == "sharedItems":
                    field.remove(child)
    return xml_node_signature(root)


def grade_chart_caches(checks, payloads, model, task):
    """Check present caches when their bounded local ranges have known values."""
    fresh = task.get("formula_cache_invalidation", {}).get("refreshed_values", {})
    for part, payload in payloads.items():
        if not (part.startswith("xl/charts/") and part.endswith(".xml")):
            continue
        root = ET.fromstring(payload)
        for index, series in enumerate(chart_series(root)):
            for component in ("val", "cat", "tx"):
                formula = chart_reference_formula(series, component)
                if formula is None or "!" not in formula:
                    continue
                title, reference = formula.rsplit("!", 1)
                title = title.strip("'").replace("''", "'")
                if title not in model["sheets"] or not re.fullmatch(r"[A-Za-z]+[1-9][0-9]*(?::[A-Za-z]+[1-9][0-9]*)?", reference):
                    continue
                left, top, right, bottom = range_bounds(reference)
                if (left != right and top != bottom) or (right - left + 1) * (bottom - top + 1) > 10000:
                    continue
                expected = []
                for row in range(top, bottom + 1):
                    for col in range(left, right + 1):
                        letters = ""
                        value = col
                        while value:
                            value, remainder = divmod(value - 1, 26)
                            letters = chr(65 + remainder) + letters
                        address = f"{letters}{row}"
                        cell = model["sheets"][title]["cells"].get(address, {})
                        expected.append(fresh.get(f"{title}!{address}", cell.get("value")))
                if any(value is None for value in expected):
                    continue  # A cleared formula cache is not an answer key.
                caches = [node for holder in series if local_name(holder.tag) == component
                          for node in holder.iter() if local_name(node.tag) in {"numCache", "strCache"}]
                if caches:
                    record(checks, f"chart cache {part} series {index} {component} matches referenced cells",
                           chart_cache_is_fresh(series, expected, component), kind="critical_collateral")


def probe_identity() -> tuple[int, int] | None:
    if os.geteuid() != 0:
        return None
    try:
        account = pwd.getpwnam("agent")
    except KeyError:
        return None
    return account.pw_uid, account.pw_gid


def cleanup_candidate_processes() -> None:
    identity = probe_identity()
    proc = Path("/proc")
    if identity is None or not proc.is_dir():
        return
    uid, _ = identity
    for entry in proc.iterdir():
        if not entry.name.isdigit() or int(entry.name) == os.getpid():
            continue
        try:
            status = (entry / "status").read_text()
            uid_line = next(
                line for line in status.splitlines() if line.startswith("Uid:")
            )
            real_uid = int(uid_line.split()[1])
            if real_uid == uid:
                os.kill(int(entry.name), signal.SIGKILL)
        except (FileNotFoundError, PermissionError, ProcessLookupError, StopIteration):
            continue


def _grade(root: Path, task: dict) -> dict:
    hard_failures: list[str] = []
    checks: list[dict] = []
    inputs: dict[str, tuple[Path, dict[str, bytes]]] = {}
    for item in task.get("inputs", []):
        path = root / item["path"]
        if not path.is_file():
            hard_failures.append(f"missing input {item['path']}")
            continue
        digest = sha256(path)
        if digest != item["sha256"]:
            hard_failures.append(f"input hash mismatch for {item['path']}: {digest}")
            continue
        try:
            inputs[item["path"]] = (path, safe_package(path))
        except HardFailure as exc:
            hard_failures.append(str(exc))
    for item in task.get("aux_inputs", []):
        path = root / item["path"]
        if not path.is_file():
            hard_failures.append(f"missing auxiliary input {item['path']}")
        elif sha256(path) != item["sha256"]:
            hard_failures.append(f"auxiliary input hash mismatch for {item['path']}")

    output_path = root / task["output"] if task.get("output") else None
    inspection = task.get("inspection_contract")
    if inspection:
        for forbidden in task.get("forbidden_outputs", []):
            if (root / forbidden).exists():
                hard_failures.append(f"forbidden output exists: {forbidden}")
        if output_path is None or not output_path.is_file():
            hard_failures.append("missing error-inspection JSON")
        elif output_path.stat().st_size > 64 * 1024:
            hard_failures.append("error-inspection JSON exceeds 64 KiB")
        else:
            try:
                report = json.loads(output_path.read_text())
                def error_set(items):
                    if not isinstance(items, list):
                        raise ValueError("errors must be an array")
                    values = []
                    for item in items:
                        if not isinstance(item, dict) or set(item) != {"sheet", "cell", "error"}:
                            raise ValueError("each error must contain sheet, cell and error")
                        if not all(isinstance(value, str) for value in item.values()):
                            raise ValueError("error fields must be strings")
                        values.append((item["sheet"], item["cell"].upper(), item["error"]))
                    if len(values) != len(set(values)):
                        raise ValueError("duplicate error cells")
                    return set(values)
                if not isinstance(report, dict) or set(report) != {"errors"}:
                    raise ValueError("report must contain an errors array")
                actual = error_set(report["errors"])
                expected = error_set(inspection["errors"])
                record(checks, "exact spreadsheet-error report", actual == expected,
                       f"missing={sorted(expected - actual)!r}; extra={sorted(actual - expected)!r}",
                       kind="critical_effect", effect_gate=True)
            except (OSError, UnicodeError, ValueError, TypeError) as exc:
                hard_failures.append(f"invalid error-inspection JSON: {exc}")
        return {"reward": float(not hard_failures and bool(checks) and all(item["passed"] for item in checks)),
                "hard_failures": hard_failures, "checks": checks}

    output_payloads = None
    output_model = None
    if output_path is not None:
        try:
            output_payloads = safe_package(output_path)
            output_model = workbook_model(output_payloads)
        except HardFailure as exc:
            hard_failures.append(str(exc))
    for forbidden in task.get("forbidden_outputs", []):
        if (root / forbidden).exists():
            hard_failures.append(f"forbidden output exists: {forbidden}")

    if hard_failures:
        return {"reward": 0.0, "hard_failures": hard_failures, "checks": checks}

    primary_payloads = (
        inputs[task["inputs"][0]["path"]][1] if task.get("inputs") else {}
    )
    normalized_output = normalize_comment_part_names(primary_payloads, output_payloads)
    if normalized_output is not output_payloads:
        output_payloads = normalized_output
        output_model = workbook_model(output_payloads)
    primary_model = workbook_model(primary_payloads) if primary_payloads else empty_workbook_model()
    golden_payloads = None
    golden_model = None
    if task.get("golden"):
        golden_path = Path(__file__).resolve().parent / task["golden"]
        golden_payloads = safe_package(golden_path)
        golden_payloads = normalize_comment_part_names(primary_payloads, golden_payloads)
        golden_model = workbook_model(golden_payloads)
        for title in golden_model["sheets"]:
            compare_sheet_feature(checks, output_model, golden_model, title, "hidden_rows", kind="critical_collateral")

    # Compare formula content/style independently of the two permitted cache
    # outcomes. The dedicated contract below still rejects all stale values.
    if golden_model is not None:
        for qualified, value in task.get("formula_cache_invalidation", {}).get("refreshed_values", {}).items():
            title, address = qualified.rsplit("!", 1)
            expected = golden_model["sheets"].get(title, {}).get("cells", {}).get(address)
            if expected is not None and expected.get("formula") is not None:
                expected["value"] = value
                expected["cache_state"] = "value"

    pivot_config = task.get("pivot_refresh_contract", {})
    pivot_refreshed = False
    if pivot_config.get("records_part"):
        pivot_refreshed = pivot_records(output_payloads, pivot_config) == pivot_config["refreshed_records"]
        if pivot_refreshed:
            for qualified, value in pivot_config.get("refreshed_report_values", {}).items():
                title, address = qualified.rsplit("!", 1)
                golden_model["sheets"][title]["cells"][address]["value"] = value

    creation_style = task.get("creation_style_contract")
    if creation_style:
        grade_creation_styles(checks, output_payloads, output_model, creation_style)
        # New-workbook aesthetics are not a preservation contract. Content and
        # formulas remain checked against the reference; styles have explicit
        # business-level assertions instead of one writer's font/fill choices.
        for title in task.get("sheet_cells_from_golden", []):
            actual_cells = output_model["sheets"].get(title, {}).get("cells", {})
            expected_cells = golden_model["sheets"].get(title, {}).get("cells", {})
            for address, actual in actual_cells.items():
                if address in expected_cells:
                    actual["style"] = expected_cells[address].get("style")

    for qualified in task.get("cells_from_golden", []):
        compare_cell(
            checks,
            output_model,
            golden_model,
            primary_model,
            qualified,
        )
    for title in task.get("sheet_cells_from_golden", []):
        compare_sheet_cells(
            checks,
            output_model,
            golden_model,
            primary_model,
            title,
        )
    if task.get("sheet_names_from_golden"):
        changed = primary_model["sheet_names"] != golden_model["sheet_names"]
        record(
            checks,
            "sheet names",
            output_model["sheet_names"] == golden_model["sheet_names"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("defined_names_from_golden"):
        changed = primary_model["defined_names"] != golden_model["defined_names"]
        record(
            checks,
            "defined names",
            output_model["defined_names"] == golden_model["defined_names"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("table_refs_from_golden"):
        changed = primary_model["table_refs"] != golden_model["table_refs"]
        record(
            checks,
            "table references",
            output_model["table_refs"] == golden_model["table_refs"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("tables_from_golden"):
        changed = primary_model["tables"] != golden_model["tables"]
        record(
            checks,
            "table models",
            output_model["tables"] == golden_model["tables"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("chart_formulas_from_golden"):
        changed = primary_model["chart_formulas"] != golden_model["chart_formulas"]
        record(
            checks,
            "chart formulas",
            output_model["chart_formulas"] == golden_model["chart_formulas"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("charts_from_golden"):
        if task.get("creation_chart_style_unspecified"):
            for model, payloads in ((output_model, output_payloads), (golden_model, golden_payloads)):
                model["charts"] = {part: chart_semantic_signature(payloads[part], unspecified_style=True)
                                   for part in model["charts"]}
        changed = primary_model["charts"] != golden_model["charts"]
        record(
            checks,
            "chart models",
            output_model["charts"] == golden_model["charts"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("charts_from_golden") or task.get("chart_cache_contract"):
        grade_chart_caches(checks, output_payloads, output_model, task)
    for part in task.get("drawing_parts_from_golden", []):
        actual = output_payloads.get(part)
        expected = golden_payloads.get(part)
        record(checks, f"drawing geometry and content {part}",
               actual is not None and expected is not None and
               xml_signature(resolved_xml(part, actual, output_payloads)) ==
               xml_signature(resolved_xml(part, expected, golden_payloads)),
               kind="critical_collateral")
    if task.get("core_properties_from_golden"):
        changed = primary_model["core_properties"] != golden_model["core_properties"]
        record(
            checks,
            "core properties",
            output_model["core_properties"] == golden_model["core_properties"],
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    for title in task.get("formula_topology_from_golden", []):
        actual = formula_topology(output_model, title)
        expected = formula_topology(golden_model, title)
        before = formula_topology(primary_model, title)
        changed = expected != before
        record(
            checks,
            f"{title} formula topology",
            actual == expected,
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    for qualified in task.get("cell_values_from_golden", []):
        title, address = qualified.rsplit("!", 1)
        actual = output_model["sheets"].get(title, {}).get("cells", {}).get(address)
        expected = golden_model["sheets"].get(title, {}).get("cells", {}).get(address)
        before = primary_model["sheets"].get(title, {}).get("cells", {}).get(address)
        changed = expected is not None and (
            before is None
            or not equivalent_cached_value(expected["value"], before["value"])
        )
        record(
            checks,
            f"cell value {qualified}",
            actual is not None
            and expected is not None
            and equivalent_cached_value(actual["value"], expected["value"]),
            f"expected={expected!r}; actual={actual!r}",
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )

    for title, features in task.get("sheet_features_from_golden", {}).items():
        for feature in features:
            expected = golden_model["sheets"].get(title, {}).get(feature)
            before = primary_model["sheets"].get(title, {}).get(feature)
            compare_sheet_feature(
                checks,
                output_model,
                golden_model,
                title,
                feature,
                kind="critical_effect" if expected != before else "critical_collateral",
                effect_gate=expected != before,
            )
    for title, features in task.get("sheet_features_equal_input", {}).items():
        for feature in features:
            compare_sheet_feature(
                checks,
                output_model,
                primary_model,
                title,
                feature,
                kind="critical_collateral",
            )
            if feature == "sparkline_groups":
                actual = output_model["sheets"].get(title, {}).get(feature)
                expected = primary_model["sheets"].get(title, {}).get(feature)
                if actual != expected:
                    hard_failures.append(
                        f"sparkline semantics changed on {title}: "
                        f"expected={expected!r}; actual={actual!r}"
                    )
    for title, ignored_addresses in task.get(
        "sheet_cells_equal_input_except", {}
    ).items():
        ignored = set(ignored_addresses)
        actual = {
            address: cell
            for address, cell in output_model["sheets"]
            .get(title, {})
            .get("cells", {})
            .items()
            if address not in ignored
        }
        expected = {
            address: cell
            for address, cell in primary_model["sheets"]
            .get(title, {})
            .get("cells", {})
            .items()
            if address not in ignored
        }
        addresses = sorted(set(actual) | set(expected))
        correct = sum(
            cells_semantically_equal(
                actual.get(address),
                expected.get(address),
                actual_model=output_model,
            )
            for address in addresses
        )
        score = correct / len(addresses) if addresses else 1.0
        record(
            checks,
            f"unowned cells on {title}",
            score == 1.0,
            f"correct={correct}/{len(addresses)}",
            kind="critical_collateral",
            score=score,
        )

    allowed_changed_parts = set(task.get("allowed_changed_parts", []))
    if task.get("image_replacement_contract"):
        allowed_changed_parts.update(grade_image_replacement(
            checks, primary_payloads, output_payloads, task["image_replacement_contract"], root))
    if allowed_changed_parts:
        all_parts = set(primary_payloads) | set(output_payloads)
        unexpected = sorted(
            part
            for part in all_parts
            if part not in allowed_changed_parts
            and not unowned_part_equal(
                part,
                primary_payloads.get(part),
                output_payloads.get(part),
                output_model=output_model,
                before_package=primary_payloads,
                after_package=output_payloads,
            )
        )
        record(
            checks,
            "unowned package parts byte-identical",
            not unexpected,
            f"unexpected changed parts={unexpected}",
            kind="critical_collateral",
        )

    for part in task.get("parts_equal_input", []):
        actual = output_payloads.get(part)
        expected = primary_payloads.get(part)
        record(
            checks,
            f"preserve input part {part}",
            actual is not None
            and expected is not None
            and unowned_part_equal(
                part,
                expected,
                actual,
                output_model=output_model,
                before_package=primary_payloads,
                after_package=output_payloads,
            ),
            kind="critical_collateral",
        )
    for part in task.get("parts_equal_golden", []):
        actual = output_payloads.get(part)
        expected = golden_payloads.get(part)
        before = primary_payloads.get(part)
        if actual is None or expected is None:
            passed = False
        elif part.endswith((".xml", ".rels", ".vml")):
            passed = xml_signature(resolved_xml(part, actual, output_payloads)) == xml_signature(resolved_xml(part, expected, golden_payloads))
        else:
            passed = actual == expected
        record(
            checks,
            f"match golden part {part}",
            passed,
            kind="critical_effect" if expected != before else "critical_collateral",
            effect_gate=expected != before,
        )
    for part in task.get("parts_semantically_equal_golden", []):
        actual = output_payloads.get(part)
        expected = golden_payloads.get(part)
        before = primary_payloads.get(part)
        if actual is None or expected is None:
            passed = False
        else:
            passed = semantic_part_signature(part, actual) == semantic_part_signature(
                part, expected
            )
        changed = before is None or semantic_part_signature(
            part, before
        ) != semantic_part_signature(part, expected)
        record(
            checks,
            f"semantically match golden part {part}",
            passed,
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    for part in task.get("parts_present", []):
        record(
            checks,
            f"required part {part}",
            part in output_payloads,
            kind="critical_collateral",
        )
    for part in task.get("parts_absent", []):
        changed = part in primary_payloads
        record(
            checks,
            f"absent part {part}",
            part not in output_payloads,
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    for rule in task.get("parts_unreachable", []):
        part = rule["part"]
        evidence = part_reachability(output_payloads, part)
        relationship_source = rule.get("relationship_source")
        source_relationships = [
            item
            for item in evidence["relationships"]
            if item["source"] == relationship_source
        ]
        record(
            checks,
            f"absent private part {part}",
            not evidence["part_present"],
            f"evidence={evidence!r}",
            kind="critical_effect",
            effect_gate=True,
        )
        record(
            checks,
            f"absent workbook relationship to {part}",
            not source_relationships,
            f"relationships={source_relationships!r}",
            kind="critical_effect",
            effect_gate=True,
        )
        record(
            checks,
            f"absent content type for {part}",
            not evidence["content_type_entries"],
            f"entries={evidence['content_type_entries']!r}",
            kind="critical_effect",
            effect_gate=True,
        )
        record(
            checks,
            f"no package relationship reaches {part}",
            not evidence["relationships"],
            f"relationships={evidence['relationships']!r}",
            kind="critical_effect",
            effect_gate=True,
        )
        if (
            evidence["part_present"]
            or source_relationships
            or evidence["content_type_entries"]
            or evidence["relationships"]
        ):
            hard_failures.append(
                f"private package part remains reachable: {part}; evidence={evidence!r}"
            )
    for rule in task.get("part_contains", []):
        payload = output_payloads.get(rule["part"], b"")
        token = rule["text"].encode()
        record(
            checks,
            f"{rule['part']} contains {rule['text']}",
            rule["part"] in output_payloads and token in payload,
            kind="critical_collateral",
        )
    for rule in task.get("part_not_contains", []):
        payload = output_payloads.get(rule["part"], b"")
        before_payload = primary_payloads.get(rule["part"], b"")
        token = rule["text"].encode()
        changed = token in before_payload
        record(
            checks,
            f"{rule['part']} omits {rule['text']}",
            rule["part"] in output_payloads and token not in payload,
            kind="critical_effect" if changed else "critical_collateral",
            effect_gate=changed,
        )
    if task.get("output_equals_input"):
        source_path = inputs[task["inputs"][0]["path"]][0]
        record(
            checks,
            "output byte-identical to input",
            output_path.read_bytes() == source_path.read_bytes(),
            kind="critical_collateral",
        )

    dimension_config = task.get("dimension_contract")
    if dimension_config:
        title = dimension_config["sheet"]
        actual_sheet = output_model["sheets"].get(title, {})
        before_sheet = primary_model["sheets"].get(title, {})
        for row in dimension_config.get("row_heights", []):
            record(
                checks,
                f"preserve row {row} height",
                actual_sheet.get("row_heights", {}).get(row, actual_sheet.get("row_heights", {}).get(0))
                == before_sheet.get("row_heights", {}).get(row, before_sheet.get("row_heights", {}).get(0)),
                kind="critical_collateral",
            )
        for column in dimension_config.get("column_widths", []):
            index = 0
            for character in column.upper():
                index = index * 26 + ord(character) - 64
            record(
                checks,
                f"preserve column {column} width",
                actual_sheet.get("column_widths", {}).get(index)
                == before_sheet.get("column_widths", {}).get(index),
                kind="critical_collateral",
            )

    chart_config = task.get("chart_cache_contract")
    if chart_config:
        part = chart_config["part"]
        actual_payload = output_payloads.get(part)
        expected_payload = golden_payloads.get(part)
        before_payload = primary_payloads.get(part)
        if actual_payload is None or expected_payload is None or before_payload is None:
            hard_failures.append(f"chart cache contract missing part {part}")
        else:
            actual_root = ET.fromstring(actual_payload)
            expected_root = ET.fromstring(expected_payload)
            before_root = ET.fromstring(before_payload)
            index = chart_config["series_index"]
            actual_series = chart_series(actual_root)
            expected_series = chart_series(expected_root)
            before_series = chart_series(before_root)
            valid_index = index < len(actual_series) == len(expected_series) == len(before_series)
            record(checks, "target chart series exists", valid_index, kind="behavioral")
            if valid_index:
                formula = chart_values_formula(actual_series[index])
                record(
                    checks,
                    "target values-series formula",
                    formula == canonical_chart_formula(chart_config["values_formula"]),
                    f"actual={formula!r}",
                    kind="critical_effect",
                    effect_gate=True,
                )
                record(
                    checks,
                    "target chart cache fresh or invalidated",
                    not chart_values_has_cache(actual_series[index]) or (
                        "refreshed_values" in chart_config
                        and chart_cache_is_fresh(actual_series[index], chart_config["refreshed_values"])
                    ),
                    kind="critical_effect",
                    effect_gate=True,
                )
                if "categories_formula" in chart_config:
                    record(checks, "target categories-series formula",
                           chart_reference_formula(actual_series[index], "cat") == canonical_chart_formula(chart_config["categories_formula"]),
                           kind="critical_effect", effect_gate=True)
                    has_categories_cache = any(local_name(child.tag) in {"numCache", "strCache"}
                                               for node in actual_series[index].iter() if local_name(node.tag) == "cat"
                                               for child in node.iter())
                    record(checks, "target category cache fresh or invalidated",
                           not has_categories_cache or ("refreshed_categories" in chart_config and chart_cache_is_fresh(
                               actual_series[index], chart_config["refreshed_categories"], "cat")),
                           kind="critical_effect", effect_gate=True)
                    clear_chart_values_cache(actual_series[index], "cat")
                    clear_chart_values_cache(expected_series[index], "cat")
                    normalize_category_reference(actual_series[index])
                    normalize_category_reference(expected_series[index])
                for series_index in range(len(actual_series)):
                    if series_index == index:
                        continue
                    record(
                        checks,
                        f"non-target chart series {series_index} preserved",
                        xml_node_signature(actual_series[series_index])
                        == xml_node_signature(before_series[series_index]),
                        kind="critical_collateral",
                    )
                clear_chart_values_cache(actual_series[index])
                clear_chart_values_cache(expected_series[index])
                record(
                    checks,
                    "complete chart semantics match oracle",
                    chart_semantic_signature(ET.tostring(actual_root)) == chart_semantic_signature(ET.tostring(expected_root)),
                    kind="critical_collateral",
                )

    pivot_config = task.get("pivot_refresh_contract")
    if pivot_config:
        part = pivot_config["cache_part"]
        actual_payload = output_payloads.get(part)
        expected_payload = golden_payloads.get(part)
        if actual_payload is None or expected_payload is None:
            hard_failures.append(f"pivot refresh contract missing {part}")
        else:
            actual_root = ET.fromstring(actual_payload)
            record(
                checks,
                "affected pivot selected for refresh on open",
                pivot_refreshed or truthy_xml(actual_root.attrib.get("refreshOnLoad")),
                kind="behavioral",
                effect_gate=True,
            )
            record(
                checks,
                "pivot cache definition retains unrelated content",
                pivot_definition_signature(part, actual_payload, output_payloads, refreshed=pivot_refreshed)
                == pivot_definition_signature(part, expected_payload, golden_payloads, refreshed=pivot_refreshed),
                kind="critical_collateral",
            )
        for preserve_part in pivot_config.get("preserve_parts", []):
            record(
                checks,
                f"preserve pivot part {preserve_part}",
                unowned_part_equal(preserve_part, primary_payloads.get(preserve_part), output_payloads.get(preserve_part),
                                   output_model=output_model, before_package=primary_payloads, after_package=output_payloads),
                kind="critical_collateral",
            )
        if pivot_config.get("records_part"):
            records_part = pivot_config["records_part"]
            deferred_records = unowned_part_equal(records_part, primary_payloads.get(records_part), output_payloads.get(records_part),
                                                 output_model=output_model, before_package=primary_payloads, after_package=output_payloads)
            record(checks, "affected pivot records correctly refreshed or safely deferred", pivot_refreshed or deferred_records,
                   kind="critical_effect", effect_gate=True)

    cache_config = task.get("formula_cache_invalidation")
    if cache_config:
        recalculate_on_open = recalc_posture(output_model)
        for qualified in cache_config.get("formula_cells", []):
            title, address = qualified.rsplit("!", 1)
            actual = output_model["sheets"].get(title, {}).get("cells", {}).get(address)
            before = (
                primary_model["sheets"].get(title, {}).get("cells", {}).get(address)
            )
            formula_preserved = (
                actual is not None
                and before is not None
                and canonical_formula(actual["formula"])
                == canonical_formula(before["formula"])
            )
            record(
                checks,
                f"formula preserved at {qualified}",
                formula_preserved,
                kind="critical_collateral",
            )
            expected_value = cache_config.get("refreshed_values", {}).get(qualified)
            if actual is None:
                cache_mode = "missing-formula-cell"
            elif actual["cache_state"] == "missing":
                cache_mode = "cleared"
            elif actual["cache_state"] == "empty":
                cache_mode = "explicitly-invalidated"
            elif expected_value is not None and equivalent_cached_value(
                actual["value"], expected_value
            ):
                cache_mode = "refreshed"
            else:
                cache_mode = "stale"
            cache_safe = cache_mode == "refreshed" or (
                cache_mode in {"cleared", "explicitly-invalidated"} and recalculate_on_open
            )
            record(
                checks,
                f"safe cache state at {qualified}",
                cache_safe,
                f"mode={cache_mode}; before={before!r}; actual={actual!r}; "
                f"expected_refreshed={expected_value!r}; recalc={recalculate_on_open}",
                kind="critical_effect",
            )
            if not formula_preserved or not cache_safe:
                hard_failures.append(
                    f"unsafe formula-cache state at {qualified}: "
                    f"formula_preserved={formula_preserved}; mode={cache_mode}; "
                    f"recalc={recalculate_on_open}"
                )
        for qualified in cache_config.get("unrelated_formula_cells", []):
            title, address = qualified.rsplit("!", 1)
            actual = output_model["sheets"].get(title, {}).get("cells", {}).get(address)
            before = primary_model["sheets"].get(title, {}).get("cells", {}).get(address)
            preserved = (
                actual is not None
                and before is not None
                and canonical_formula(actual.get("formula")) == canonical_formula(before.get("formula"))
                and actual.get("cache_state") == before.get("cache_state")
                and equivalent_cached_value(actual.get("value"), before.get("value"))
            )
            record(
                checks,
                f"unrelated formula cache preserved at {qualified}",
                preserved,
                f"before={before!r}; actual={actual!r}",
                kind="critical_collateral",
            )
            if not preserved:
                hard_failures.append(f"unrelated formula cache changed at {qualified}")

    if not checks:
        hard_failures.append("task has no assertions")
    reward = float(
        not hard_failures
        and bool(checks)
        and all(check["passed"] for check in checks)
    )
    return {
        "reward": round(reward, 6),
        "hard_failures": hard_failures,
        "checks": checks,
    }


def grade(root: Path, task: dict) -> dict:
    try:
        return _grade(root, task)
    except Exception as exc:
        return {
            "reward": 0.0,
            "hard_failures": [
                f"verifier abort: {type(exc).__name__}: {str(exc)[:1000]}"
            ],
            "checks": [],
        }


def main() -> int:
    if len(sys.argv) == 2 and sys.argv[1] == "--cleanup-candidate-processes":
        cleanup_candidate_processes()
        return 0
    if len(sys.argv) != 4:
        print("usage: grader.py ROOT REWARD_PATH GRADING_PATH", file=sys.stderr)
        return 2
    root = Path(sys.argv[1])
    reward_path = Path(sys.argv[2])
    grading_path = Path(sys.argv[3])
    result = grade(root, globals()["TASK"])
    reward_path.parent.mkdir(parents=True, exist_ok=True)
    grading_path.parent.mkdir(parents=True, exist_ok=True)
    reward_path.write_text(f"{result['reward']:.6f}\n")
    grading_path.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
