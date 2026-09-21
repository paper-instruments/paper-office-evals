#!/usr/bin/env python3
"""Standalone Harbor grader for FM22."""

from __future__ import annotations

import hashlib
import json
import os
import posixpath
import re
import stat
import sys
import unicodedata
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET
from preservation import changed_parts as semantic_changed_parts
from preservation import equivalent_part

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
PKG_R = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REPORT_MAX_BYTES = 2 * 1024 * 1024
ARCHIVE_MAX_BYTES = 100 * 1024 * 1024
ARCHIVE_MAX_ENTRIES = 4096
ARCHIVE_MAX_MEMBER_BYTES = 64 * 1024 * 1024
ARCHIVE_MAX_EXPANDED_BYTES = 256 * 1024 * 1024
ARCHIVE_MAX_XML_BYTES = 16 * 1024 * 1024
ARCHIVE_MAX_TOTAL_XML_BYTES = 64 * 1024 * 1024
ARCHIVE_MAX_RATIO = 200
XML_MAX_ELEMENTS = 200000
XML_MAX_DEPTH = 128
OFFICE_DOCUMENT_REL_TYPES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
}
RELATIONSHIPS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"
ALLOWED_EXTERNAL_SCHEMES = {"http", "https", "mailto"}
EXPECTED_HASHES = {'eval_fixture_02_edit_surgery.docx': '7cdbe138dbc22ff848fc3f05968cb03b9dc4a302a6504bda922006b012344c8e', 'eval_fixture_04_review_and_hygiene.docx': '1c5f932267adec14e52de9121b30b6f0c83e23af4a32dbba64cb5f00c1d01101'}
TASK_INPUTS = {'cross-document-composition': ('eval_fixture_02_edit_surgery.docx', 'eval_fixture_04_review_and_hygiene.docx')}
TASK_OUTPUTS = {'cross-document-composition': ('evals/cross-document-composition/output.docx',)}
NO_DOCX_OUTPUT_TASKS = set()
SUPPLIED_ARCHIVES = {}

class VerificationError(AssertionError):
    pass

@dataclass
class CheckResult:
    name: str
    weight: float
    passed: bool
    hard: bool
    message: Optional[str] = None

@dataclass
class ScoreCard:
    task: str
    checks: list[CheckResult] = field(default_factory=list)

    def hard(self, name: str, operation: Any) -> bool:
        try:
            operation()
        except Exception as exc:
            self.checks.append(CheckResult(name, 0.0, False, True, f"{type(exc).__name__}: {exc}"))
            return False
        self.checks.append(CheckResult(name, 0.0, True, True))
        return True

    def weighted(self, name: str, weight: float, operation: Any) -> bool:
        try:
            operation()
        except VerificationError as exc:
            self.checks.append(
                CheckResult(name, weight, False, False, f"{type(exc).__name__}: {exc}")
            )
            return False
        except Exception as exc:
            self.checks.append(
                CheckResult(name, 0.0, False, True, f"verifier abort: {type(exc).__name__}: {exc}")
            )
            return False
        self.checks.append(CheckResult(name, weight, True, False))
        return True

    @property
    def hard_failed(self) -> bool:
        return any((item.hard and (not item.passed) for item in self.checks))

    @property
    def score(self) -> float:
        semantic = [item for item in self.checks if not item.hard]
        return float(
            bool(semantic)
            and all(item.passed for item in self.checks)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": "docx-harbor-grade",
            "version": 1,
            "task": self.task,
            "score": round(self.score, 6),
            "hard_failed": self.hard_failed,
            "checks": [
                {
                    "name": item.name,
                    "weight": item.weight,
                    "passed": item.passed,
                    "hard": item.hard,
                    "message": item.message,
                }
                for item in self.checks
            ],
        }

_ACTIVE_SEMANTIC_CARD: Optional[ScoreCard] = None
_ACTIVE_EFFECT_PATTERNS: tuple[str, ...] = ()
_ACTIVE_COLLATERAL_PATTERNS: tuple[str, ...] = ()
_SEMANTIC_CHECKS_SUPPRESSED = 0

def check(condition: bool, message: str) -> None:
    if _ACTIVE_SEMANTIC_CARD is None or _SEMANTIC_CHECKS_SUPPRESSED:
        if not condition:
            raise VerificationError(message)
        return

    folded = message.casefold()
    if any(pattern.casefold() in folded for pattern in _ACTIVE_EFFECT_PATTERNS):
        role, weight = "critical-effect", 4.0
    elif any(
        pattern.casefold() in folded for pattern in _ACTIVE_COLLATERAL_PATTERNS
    ):
        role, weight = "critical-collateral", 3.0
    else:
        role, weight = "supporting", 1.0
    _ACTIVE_SEMANTIC_CARD.checks.append(
        CheckResult(f"{role}: {message}", weight, bool(condition), False)
    )

def collect_semantic_checks(
    card: ScoreCard,
    operation: Any,
    *,
    effect_patterns: tuple[str, ...],
    collateral_patterns: tuple[str, ...],
) -> None:
    global _ACTIVE_SEMANTIC_CARD
    global _ACTIVE_EFFECT_PATTERNS
    global _ACTIVE_COLLATERAL_PATTERNS

    _ACTIVE_SEMANTIC_CARD = card
    _ACTIVE_EFFECT_PATTERNS = effect_patterns
    _ACTIVE_COLLATERAL_PATTERNS = collateral_patterns
    try:
        operation()
    except Exception as exc:
        card.checks.append(
            CheckResult(
                "semantic verifier execution",
                0.0,
                False,
                True,
                f"verifier abort: {type(exc).__name__}: {exc}",
            )
        )
    finally:
        _ACTIVE_SEMANTIC_CARD = None
        _ACTIVE_EFFECT_PATTERNS = ()
        _ACTIVE_COLLATERAL_PATTERNS = ()

    semantic_names = [item.name.casefold() for item in card.checks if not item.hard]
    for role, patterns in (
        ("critical-effect", effect_patterns),
        ("critical-collateral", collateral_patterns),
    ):
        for pattern in patterns:
            expected = f"{role}: "
            if not any(
                name.startswith(expected) and pattern.casefold() in name
                for name in semantic_names
            ):
                card.checks.append(
                    CheckResult(
                        f"{role} assertion coverage: {pattern}",
                        0.0,
                        False,
                        True,
                        "configured semantic assertion did not execute",
                    )
                )

def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]

def root_for(part: bytes) -> ET.Element:
    return ET.fromstring(part)

def element_text(element: ET.Element, *, include_instr: bool = True) -> str:
    accepted = {W + "t", W + "delText"}
    if include_instr:
        accepted.add(W + "instrText")
    return "".join((node.text or "" for node in element.iter() if node.tag in accepted))

def _parse_xml_limited(payload: bytes, *, label: str) -> ET.Element:
    check(len(payload) <= ARCHIVE_MAX_XML_BYTES, f"XML part exceeds limit: {label}")
    probe = payload.replace(b"\x00", b"").lower()
    check(
        b"<!doctype" not in probe and b"<!entity" not in probe,
        f"DTD/entity is forbidden in {label}",
    )
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        raise VerificationError(f"malformed XML in {label}: {exc}") from exc
    count = 0
    stack = [(root, 1)]
    while stack:
        node, depth = stack.pop()
        count += 1
        check(count <= XML_MAX_ELEMENTS, f"XML element limit exceeded in {label}")
        check(depth <= XML_MAX_DEPTH, f"XML depth limit exceeded in {label}")
        stack.extend(((child, depth + 1) for child in node))
    return root

def _member_name(name: str) -> str:
    check(name and len(name.encode("utf-8")) <= 1024, "ZIP member name is empty or too long")
    check(
        not any((ord(char) < 32 or ord(char) == 127 for char in name)),
        f"unsafe control character in ZIP member {name!r}",
    )
    check("\\" not in name and "\x00" not in name, f"unsafe ZIP member path {name!r}")
    check(not name.startswith(("/", "//")), f"absolute ZIP member path {name!r}")
    check(not re.match("^[A-Za-z]:", name), f"drive-qualified ZIP member path {name!r}")
    check("//" not in name, f"repeated separator in ZIP member path {name!r}")
    raw_segments = name.rstrip("/").split("/")
    check(
        all((segment not in {"", ".", ".."} for segment in raw_segments)),
        f"traversal ZIP member path {name!r}",
    )
    decoded = unquote(name)
    check("\\" not in decoded, f"percent-encoded separator in ZIP member path {name!r}")
    check(
        not decoded.startswith("/") and (not re.match("^[A-Za-z]:", decoded)),
        f"encoded absolute ZIP member path {name!r}",
    )
    decoded_segments = decoded.rstrip("/").split("/")
    check(
        all((segment not in {"", ".", ".."} for segment in decoded_segments)),
        f"encoded traversal ZIP member path {name!r}",
    )
    return name.rstrip("/")

def _rels_source_part(name: str) -> Optional[str]:
    if name == "_rels/.rels":
        return None
    marker = "/_rels/"
    check(marker in name and name.endswith(".rels"), f"invalid relationship part location: {name}")
    prefix, leaf = name.rsplit(marker, 1)
    check(bool(prefix) and leaf != ".rels", f"invalid relationship part location: {name}")
    return f"{prefix}/{leaf[:-5]}"

def _internal_target(source: Optional[str], target: str) -> str:
    parsed = urlsplit(target)
    check(
        not parsed.scheme and (not parsed.netloc) and (not parsed.query),
        f"invalid internal relationship target {target!r}",
    )
    decoded = unquote(parsed.path)
    check(
        not re.search("%(?:2f|5c)", parsed.path, re.IGNORECASE),
        f"encoded separator in relationship target {target!r}",
    )
    check("\\" not in decoded and "\x00" not in decoded, f"unsafe relationship target {target!r}")
    check(decoded, "empty internal relationship target")
    if decoded.startswith("/"):
        combined = decoded.lstrip("/")
    else:
        base = "" if source is None else posixpath.dirname(source)
        combined = posixpath.join(base, decoded)
    normalized = posixpath.normpath(combined)
    check(
        normalized not in {"", ".", ".."} and (not normalized.startswith("../")),
        f"relationship target escapes package: {target!r}",
    )
    return normalized

def _content_type_for(name: str, defaults: dict[str, str], overrides: dict[str, str]) -> Optional[str]:
    override = overrides.get("/" + name.casefold())
    if override is not None:
        return override
    leaf = name.rsplit("/", 1)[-1]
    if "." not in leaf:
        return None
    return defaults.get(leaf.rsplit(".", 1)[-1].casefold())

def validate_opc_package(path: Path) -> tuple[list[zipfile.ZipInfo], dict[str, bytes]]:
    check(path.is_file(), f"missing required DOCX: {path.name}")
    check(path.stat().st_size <= ARCHIVE_MAX_BYTES, f"archive exceeds size limit: {path.name}")
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            check(len(infos) <= ARCHIVE_MAX_ENTRIES, f"too many ZIP members in {path.name}")
            exact: set[str] = set()
            folded: dict[str, str] = {}
            total = 0
            xml_total = 0
            parts: dict[str, bytes] = {}
            for info in infos:
                normalized = _member_name(info.filename)
                check(info.filename not in exact, f"duplicate ZIP member name {info.filename!r}")
                exact.add(info.filename)
                key = unicodedata.normalize("NFC", info.filename).casefold()
                check(key not in folded, f"case-ambiguous ZIP member name {info.filename!r}")
                folded[key] = info.filename
                check(not info.flag_bits & 1, f"encrypted ZIP member {info.filename!r}")
                check(
                    info.compress_type in {zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED},
                    f"unsupported compression for {info.filename!r}",
                )
                mode = info.external_attr >> 16 & 61440
                check(
                    mode in {0, stat.S_IFREG, stat.S_IFDIR}, f"special ZIP member {info.filename!r}"
                )
                check(
                    info.file_size <= ARCHIVE_MAX_MEMBER_BYTES,
                    f"ZIP member exceeds limit: {info.filename}",
                )
                total += info.file_size
                check(
                    total <= ARCHIVE_MAX_EXPANDED_BYTES, f"expanded ZIP exceeds limit: {path.name}"
                )
                if info.file_size and info.compress_size:
                    check(
                        info.file_size / info.compress_size <= ARCHIVE_MAX_RATIO,
                        f"compression ratio exceeds limit: {info.filename}",
                    )
                if info.is_dir():
                    continue
                with archive.open(info) as stream:
                    chunks: list[bytes] = []
                    seen = 0
                    while True:
                        chunk = stream.read(min(1024 * 1024, ARCHIVE_MAX_MEMBER_BYTES + 1 - seen))
                        if not chunk:
                            break
                        seen += len(chunk)
                        check(
                            seen <= ARCHIVE_MAX_MEMBER_BYTES,
                            f"streamed member exceeds limit: {info.filename}",
                        )
                        chunks.append(chunk)
                payload = b"".join(chunks)
                check(len(payload) == info.file_size, f"ZIP size mismatch for {info.filename}")
                parts[normalized] = payload
                if normalized.endswith((".xml", ".rels")) or normalized == "[Content_Types].xml":
                    xml_total += len(payload)
                    check(
                        xml_total <= ARCHIVE_MAX_TOTAL_XML_BYTES,
                        f"total XML exceeds limit: {path.name}",
                    )
            bad = archive.testzip()
            check(bad is None, f"corrupt ZIP entry in {path.name}: {bad}")
    except zipfile.BadZipFile as exc:
        raise VerificationError(f"{path.name} is not a valid ZIP package: {exc}") from exc
    check("[Content_Types].xml" in parts, f"{path.name} lacks [Content_Types].xml")
    check("_rels/.rels" in parts, f"{path.name} lacks _rels/.rels")
    parsed_xml: dict[str, ET.Element] = {}
    for name, payload in parts.items():
        if name.endswith((".xml", ".rels")) or name == "[Content_Types].xml":
            parsed_xml[name] = _parse_xml_limited(payload, label=f"{path.name}:{name}")
    ct_root = parsed_xml["[Content_Types].xml"]
    check(ct_root.tag == CT + "Types", "[Content_Types].xml has wrong root")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in ct_root:
        if child.tag == CT + "Default":
            extension = (child.get("Extension") or "").casefold()
            content_type = child.get("ContentType") or ""
            check(
                extension and content_type and (extension not in defaults),
                f"duplicate/empty content-type default {extension!r}",
            )
            defaults[extension] = content_type
        elif child.tag == CT + "Override":
            part_name = child.get("PartName") or ""
            content_type = child.get("ContentType") or ""
            check(
                part_name.startswith("/") and content_type,
                f"invalid content-type override {part_name!r}",
            )
            key = "/" + _member_name(part_name.lstrip("/")).casefold()
            check(key not in overrides, f"duplicate content-type override {part_name!r}")
            check(
                key[1:] in {name.casefold() for name in parts},
                f"orphan content-type override {part_name!r}",
            )
            overrides[key] = content_type
        else:
            raise VerificationError(f"unexpected content-type child {local(child.tag)!r}")
    check(
        defaults.get("rels") == RELATIONSHIPS_CONTENT_TYPE,
        "missing/incorrect .rels content-type default",
    )
    for name in parts:
        if name == "[Content_Types].xml":
            continue
        check(
            _content_type_for(name, defaults, overrides) is not None,
            f"package part lacks content type: {name}",
        )
    root_office_targets: list[str] = []
    for rel_name, root in parsed_xml.items():
        if not rel_name.endswith(".rels"):
            continue
        check(root.tag == PKG_R + "Relationships", f"relationship part has wrong root: {rel_name}")
        source = _rels_source_part(rel_name)
        if source is not None:
            check(source in parts, f"relationship source part is missing: {source}")
        seen_ids: set[str] = set()
        for rel in root:
            check(rel.tag == PKG_R + "Relationship", f"unexpected relationship child in {rel_name}")
            ident = rel.get("Id") or ""
            rel_type = rel.get("Type") or ""
            target = rel.get("Target") or ""
            mode = rel.get("TargetMode")
            check(ident and ident not in seen_ids, f"duplicate/empty relationship Id in {rel_name}")
            seen_ids.add(ident)
            check(
                urlsplit(rel_type).scheme in {"http", "https"},
                f"relationship Type must be an absolute URI in {rel_name}",
            )
            check(target, f"empty relationship target in {rel_name}:{ident}")
            check(mode in {None, "External"}, f"invalid TargetMode in {rel_name}:{ident}")
            if mode == "External":
                parsed = urlsplit(target)
                check(
                    parsed.scheme.casefold() in ALLOWED_EXTERNAL_SCHEMES
                    and (not parsed.username)
                    and (not parsed.password),
                    f"unsafe external relationship target {target!r}",
                )
                check(
                    not any((ord(char) < 32 for char in target)),
                    f"control character in external target {target!r}",
                )
                continue
            resolved = _internal_target(source, target)
            check(
                resolved in parts,
                f"relationship {ident!r} targets missing package part /{resolved}",
            )
            if source is None and rel_type in OFFICE_DOCUMENT_REL_TYPES:
                root_office_targets.append(resolved)
    check(
        len(root_office_targets) == 1,
        "package must have exactly one internal root officeDocument relationship",
    )
    main = root_office_targets[0]
    check(main == "word/document.xml", f"unexpected Word main document target: {main}")
    check(
        _content_type_for(main, defaults, overrides)
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "Word main document has wrong or macro-enabled content type",
    )
    check(
        not any((name.casefold().endswith(("vbaproject.bin", "vbadata.xml")) for name in parts)),
        "macro content is forbidden in DOCX output",
    )
    return (infos, parts)

class Docx:
    def __init__(self, path: Path):
        global _SEMANTIC_CHECKS_SUPPRESSED
        _SEMANTIC_CHECKS_SUPPRESSED += 1
        try:
            self.path = path
            self.infos, self.parts = validate_opc_package(path)
            check("word/document.xml" in self.parts, f"{path.name} lacks word/document.xml")
            self.document = root_for(self.parts["word/document.xml"])
        finally:
            _SEMANTIC_CHECKS_SUPPRESSED -= 1

    def root(self, name: str) -> ET.Element:
        check(name in self.parts, f"{self.path.name} lacks {name}")
        return root_for(self.parts[name])

    def text(self, name: str = "word/document.xml") -> str:
        return element_text(self.root(name))

    def all_xml_text(self) -> str:
        values = []
        for name in sorted(self.parts):
            if name.endswith(".xml"):
                values.append(element_text(root_for(self.parts[name])))
        return "\n".join(values)

    def body(self) -> ET.Element:
        body = self.document.find(W + "body")
        check(body is not None, "document has no body")
        return body

    def paragraphs(self) -> list[ET.Element]:
        return list(self.document.iter(W + "p"))

    def find_paragraph(self, needle: str) -> ET.Element:
        matches = [p for p in self.paragraphs() if needle in element_text(p)]
        check(
            len(matches) == 1, f"expected one paragraph containing {needle!r}, found {len(matches)}"
        )
        return matches[0]

def paragraph_style(paragraph: ET.Element) -> Optional[str]:
    node = paragraph.find("./" + W + "pPr/" + W + "pStyle")
    return None if node is None else node.get(W + "val")

def bookmarks(doc: Docx) -> dict[str, str]:
    result: dict[str, str] = {}
    starts = {node.get(W + "id"): node for node in doc.document.iter(W + "bookmarkStart")}
    parents: dict[ET.Element, ET.Element] = {}
    for parent in doc.document.iter():
        for child in parent:
            parents[child] = parent
    for ident, start in starts.items():
        name = start.get(W + "name")
        if not name:
            continue
        parent = parents.get(start)
        result[name] = "" if parent is None else element_text(parent, include_instr=False)
    return result

def require_source_hashes(app: Path, task: str) -> None:
    for filename in TASK_INPUTS[task]:
        path = app / "eval_fixtures" / "docx" / filename
        check(path.is_file(), f"missing immutable input {filename}")
        check(
            sha256(path) == EXPECTED_HASHES[filename], f"immutable input was modified: {filename}"
        )

def require_declared_outputs(app: Path, task: str) -> None:
    for name in TASK_OUTPUTS[task]:
        path = app / name
        check(
            path.is_file() and stat.S_ISREG(path.lstat().st_mode) and (not path.is_symlink()),
            f"missing regular output {name}",
        )

def require_all_output_packages_valid(app: Path, task: str) -> None:
    for name in TASK_OUTPUTS[task]:
        if name.casefold().endswith(".docx"):
            Docx(app / name)

def protection_enforced(doc: Docx) -> bool:
    settings = doc.root("word/settings.xml")
    node = settings.find(".//" + W + "documentProtection")
    return (
        node is not None
        and node.get(W + "enforcement") in {"1", "true"}
        and (node.get(W + "edit") == "trackedChanges")
    )

def relationship_targets(doc: Docx, *, external_only: bool = False) -> list[str]:
    targets: list[str] = []
    for name, payload in doc.parts.items():
        if not name.endswith(".rels"):
            continue
        root = root_for(payload)
        for rel in root.findall("./" + PKG_R + "Relationship"):
            if not external_only or rel.get("TargetMode") == "External":
                targets.append(rel.get("Target", ""))
    return targets

def paragraph_numbering(paragraph: ET.Element) -> Optional[tuple[str, str]]:
    num_pr = paragraph.find("./" + W + "pPr/" + W + "numPr")
    if num_pr is None:
        return None
    num_id = num_pr.find("./" + W + "numId")
    level = num_pr.find("./" + W + "ilvl")
    if num_id is None or level is None:
        return None
    return num_id.get(W + "val", ""), level.get(W + "val", "")

def verify_cross_document_composition(app: Path) -> None:
    from preservation import signature
    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    destination = Docx(app / "eval_fixtures/docx/eval_fixture_04_review_and_hygiene.docx")
    output = Docx(app / 'evals/cross-document-composition/output.docx')
    body = list(output.body())
    anchor = next(
        (
            i
            for i, node in enumerate(body)
            if 'Supplier schedule follows' in element_text(node)
        ),
        -1,
    )
    check(anchor >= 0, "destination anchor is missing")
    expected = [
        'The supplier will maintain the agreed service levels.',
        'Provide monthly service reports.',
        'Current supplier badge:',
        'Supplier portal: composition link',
        'Defined service: managed support',
        "REF ",
        'Service definition: managed support',
        'Northstar Services',
        "SourceValueAlpha1",
        "AdjacentTableBeta2",
    ]
    inserted_text = [element_text(node) for node in body[anchor + 1 : anchor + 11]]
    check(len(inserted_text) == 10, "composition did not insert ten top-level blocks")
    remaining = body[:anchor + 1] + body[anchor + 11:]
    # Word needs a separator between the appended table and the host's next table.
    boundary = anchor + 1
    if (len(remaining) > boundary + 1 and remaining[boundary].tag == W + "p"
            and len(remaining[boundary]) == 0 and not remaining[boundary].attrib
            and remaining[boundary + 1].tag == W + "tbl"):
        remaining.pop(boundary)
    check([signature(n) for n in remaining] == [signature(n) for n in destination.body()], "composition changed destination content or included material outside the source range")
    for needle, text in zip(expected, inserted_text):
        check(needle in text, f"composed block order/content is wrong near {needle!r}")
    check(
        "bound composition value" not in output.text(), "unsafe data-bound range leaked into output"
    )
    destination_para = output.find_paragraph("Destination paragraph using conflicting SharedStyle.")
    source_para = output.find_paragraph('The supplier will maintain the agreed service levels.')
    check(
        paragraph_style(destination_para) == "SharedStyle",
        "destination SharedStyle identity changed",
    )
    check(
        paragraph_style(source_para) not in {None, "SharedStyle"},
        "colliding source style was not renamed",
    )
    numbered = output.find_paragraph('Provide monthly service reports.')
    numbering = paragraph_numbering(numbered)
    check(numbering is not None, "composed numbered paragraph lost numPr")
    numbering_root = output.root("word/numbering.xml")
    check(any(node.get(W + "numId") == numbering[0] for node in numbering_root.findall("./" + W + "num")), "remapped numbering instance is undefined")
    names = set(bookmarks(output))
    check(
        "SharedTerm" in names
        and bool(names - set(bookmarks(destination))),
        "bookmark collision was not remapped",
    )
    def bookmark_text(doc, name):
        ident = next(n.get(W + 'id') for n in doc.document.iter(W + 'bookmarkStart') if n.get(W + 'name') == name)
        active, chunks = False, []
        for n in doc.document.iter():
            if n.tag == W + 'bookmarkStart' and n.get(W + 'id') == ident: active = True
            elif n.tag == W + 'bookmarkEnd' and n.get(W + 'id') == ident: active = False
            elif active and n.tag == W + 't': chunks.append(n.text or '')
        return ''.join(chunks)
    candidates = [name for name in names - set(bookmarks(destination)) if bookmark_text(output, name) == bookmark_text(source, 'SharedTerm')]
    check(len(candidates) == 1, "imported bookmark must preserve its target text")
    remapped = candidates[0]
    instructions = ["".join(node.text or "" for node in paragraph.iter(W + "instrText")) for paragraph in output.paragraphs()]
    check(any(remapped in value for value in instructions), "REF field did not follow the bookmark remap")
    anchors = [node.get(W + "anchor") for node in output.document.iter(W + "hyperlink") if node.get(W + "anchor")]
    check(remapped in anchors, "internal hyperlink did not follow the bookmark remap")
    ids = [node.get(W + "val") for node in output.document.iter(W + "id")]
    check(
        len([value for value in ids if value is not None])
        == len(set((value for value in ids if value is not None))),
        "content-control IDs collide",
    )
    source_media = {name for name in source.parts if name.startswith("word/media/")}
    destination_media = {name for name in destination.parts if name.startswith("word/media/")}
    output_media = {name for name in output.parts if name.startswith("word/media/")}
    check(
        len(output_media) == len(destination_media) + 1 and source_media,
        "embedded image was not copied exactly once",
    )
    new_media = output_media - destination_media
    check(len(new_media) == 1 and output.parts[next(iter(new_media))] in {source.parts[name] for name in source_media}, "copied media bytes do not match the source")
    import posixpath
    def picture_bytes(doc):
        p = doc.find_paragraph('Current supplier badge:')
        relationships = {n.get('Id'): n for n in doc.root('word/_rels/document.xml.rels')}
        values = []
        for n in p.iter():
            ident = n.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if ident:
                target = relationships[ident].get('Target')
                values.append(doc.parts[posixpath.normpath(posixpath.join('word', target))])
        return values
    check(bool(picture_bytes(source)) and picture_bytes(output) == picture_bytes(source), "imported image must remain attached to its drawing")
    external = relationship_targets(output, external_only=True)
    check(
        any(("example" in target or "composition" in target for target in external)),
        "external hyperlink was not preserved",
    )
    check(protection_enforced(output), "destination protection did not persist")
    for name in destination.parts:
        if name.startswith(("word/header", "word/footer")):
            check(equivalent_part(output.parts.get(name), destination.parts[name]), f"destination story changed: {name}")
    allowed_exact = {"word/document.xml", "word/styles.xml", "word/numbering.xml", "word/settings.xml", "word/_rels/document.xml.rels", "[Content_Types].xml", "docProps/core.xml", *new_media}
    changed = semantic_changed_parts(destination.parts, output.parts)
    check(changed <= allowed_exact, f"composition exceeded its changed-part budget: {sorted(changed - allowed_exact)}")

def _write_score(card: ScoreCard) -> None:
    channel = {
        "schema": "docx-grader-score-channel",
        "version": 1,
        "score": round(card.score, 6),
        "grading": card.to_dict(),
    }
    print(json.dumps(channel, sort_keys=True, separators=(",", ":")), flush=True)

def main() -> int:
    if len(sys.argv) != 2 or sys.argv[1] != 'cross-document-composition':
        print("usage: grader.py cross-document-composition", file=sys.stderr)
        return 2
    task = 'cross-document-composition'
    app = Path(os.environ.get("HARBOR_WORKDIR", "/app"))
    if not app.exists():
        app = Path.cwd()
    card = ScoreCard(task)
    card.hard("immutable fixture hashes", lambda: require_source_hashes(app, task))
    card.hard("required output artifacts", lambda: require_declared_outputs(app, task))
    card.hard(
        "universal emitted-package validity",
        lambda: require_all_output_packages_valid(app, task),
    )
    if not card.hard_failed:
        collect_semantic_checks(
            card,
            lambda: verify_cross_document_composition(app),
            effect_patterns=('composition did not insert ten top-level blocks', 'colliding source style was not renamed', 'bookmark collision was not remapped', 'embedded image was not copied exactly once', 'external hyperlink was not preserved'),
            collateral_patterns=('destination story changed', 'destination protection did not persist'),
        )
    card.hard("post-execution immutable fixture hashes", lambda: require_source_hashes(app, task))
    _write_score(card)
    summary = card.to_dict()
    if card.score == 1.0 and not card.hard_failed:
        print(f"PASS {task} score=1.000000", file=sys.stderr)
        return 0
    print(f"FAIL {task} score={card.score:.6f}", file=sys.stderr)
    for item in summary["checks"]:
        if not item["passed"]:
            print(f"- {item['name']}: {item['message']}", file=sys.stderr)
    return 1

if __name__ == "__main__":
    raise SystemExit(main())
