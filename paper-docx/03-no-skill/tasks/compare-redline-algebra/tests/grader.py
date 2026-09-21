#!/usr/bin/env python3
"""Standalone Harbor grader for FM21."""

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
EXPECTED_HASHES = {'eval_fixture_compare_original.docx': 'ce98235807f13a63fbaa17cc8e5c2981527d30652cbec60e518bab6993cc0e23', 'eval_fixture_compare_revised.docx': '6555cf6b482cbee402a3534cafbc6f35ba7c220ed1b50d789060ab438b512d77'}
TASK_INPUTS = {'compare-redline-algebra': ('eval_fixture_compare_original.docx', 'eval_fixture_compare_revised.docx')}
TASK_OUTPUTS = {'compare-redline-algebra': ('evals/compare-redline-algebra/redline.docx',)}
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

def visible_text(element: ET.Element) -> str:
    chunks: list[str] = []

    def walk(node: ET.Element, hidden: bool = False) -> None:
        now_hidden = hidden or node.tag in {W + "del", W + "moveFrom"}
        if node.tag == W + "t" and (not now_hidden):
            chunks.append(node.text or "")
        for child in node:
            walk(child, now_hidden)

    walk(element)
    return "".join(chunks)

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

def revision_nodes(doc: Docx) -> list[ET.Element]:
    names = {
        "ins",
        "del",
        "moveFrom",
        "moveTo",
        "rPrChange",
        "pPrChange",
        "tblPrChange",
        "tblPrExChange",
        "trPrChange",
        "tcPrChange",
    }
    return [node for node in doc.document.iter() if local(node.tag) in names]

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

def story_shape(doc: Docx) -> tuple[tuple[str, str], ...]:
    return tuple((local(node.tag), tuple(tuple(visible_text(cell) for cell in row.findall(W + "tc")) for row in node.findall(W + "tr")) if node.tag == W + "tbl" else visible_text(node)) for node in doc.body())

def projected_revision_text(node: ET.Element, *, accept: bool) -> str:
    chunks: list[str] = []

    def walk(current: ET.Element, hidden: bool = False) -> None:
        if current.tag == W + "p":
            inserted = current.find("./" + W + "pPr/" + W + "rPr/" + W + "ins") is not None
            deleted = current.find("./" + W + "pPr/" + W + "rPr/" + W + "del") is not None
            hidden = hidden or (inserted and not accept) or (deleted and accept)
        if current.tag == W + "tr":
            inserted = current.find("./" + W + "trPr/" + W + "ins") is not None
            deleted = current.find("./" + W + "trPr/" + W + "del") is not None
            hidden = hidden or (inserted and not accept) or (deleted and accept)
        if current.tag in {W + "ins", W + "moveTo"}:
            hidden = hidden or not accept
        elif current.tag in {W + "del", W + "moveFrom"}:
            hidden = hidden or accept
        if not hidden and current.tag in {W + "t", W + "delText"}:
            chunks.append(current.text or "")
        for child in current:
            walk(child, hidden)

    walk(node)
    return "".join(chunks)

def projected_story_shape(doc: Docx, *, accept: bool) -> tuple[tuple[str, str], ...]:
    result: list[tuple[str, str]] = []
    for node in doc.body():
        text = projected_revision_text(node, accept=accept)
        inserted = node.tag == W + "p" and node.find("./" + W + "pPr/" + W + "rPr/" + W + "ins") is not None
        deleted = node.tag == W + "p" and node.find("./" + W + "pPr/" + W + "rPr/" + W + "del") is not None
        if (inserted and not accept) or (deleted and accept):
            continue
        if node.tag == W + "tbl":
            rows = []
            for row in node.findall(W + "tr"):
                inserted = row.find("./" + W + "trPr/" + W + "ins") is not None
                deleted = row.find("./" + W + "trPr/" + W + "del") is not None
                if (inserted and not accept) or (deleted and accept):
                    continue
                rows.append(tuple(projected_revision_text(cell, accept=accept) for cell in row.findall(W + "tc")))
            if not rows:
                continue
            text = tuple(rows)
        result.append((local(node.tag), text))
    return tuple(result)

def verify_compare_redline_algebra(app: Path) -> None:
    from preservation import equivalent_part, project_revisions, run_normalized_signature
    original = Docx(app / "eval_fixtures/docx/eval_fixture_compare_original.docx")
    revised = Docx(app / "eval_fixtures/docx/eval_fixture_compare_revised.docx")
    redline = Docx(app / 'evals/compare-redline-algebra/redline.docx')
    dated = revision_nodes(redline)
    check(
        dated
        and all(
            n.get(W + "author") == 'Review Lead'
            and n.get(W + "date") == "2026-07-13T09:00:00Z"
            for n in dated
        ),
        "redline author/date is wrong",
    )
    check(
        {local(n.tag) for n in dated} >= {"ins", "del"},
        "redline lacks insertion/deletion revisions",
    )
    ids = [node.get(W + "id") for node in dated]
    for reference, decision in ((original, 'reject'), (revised, 'accept')):
        projected = project_revisions(redline.document, {i: decision for i in ids})
        check(run_normalized_signature(projected) == run_normalized_signature(reference.document), f"{decision} projection loses formatting or structure")
    unchanged = {visible_text(p) for p in original.body() if p.tag == W + 'p'} & {visible_text(p) for p in revised.body() if p.tag == W + 'p'}
    for p in redline.body():
        if p.tag == W + 'p' and projected_revision_text(p, accept=True) in unchanged and projected_revision_text(p, accept=False) in unchanged:
            check(not any(n.tag in {W + 'ins', W + 'del'} for n in p.iter()), "unchanged clause has needless revisions")
    check(
        all((ident is not None for ident in ids)) and len(ids) == len(set(ids)),
        "redline revision IDs are missing or duplicated",
    )
    parents: dict[ET.Element, ET.Element] = {}
    for parent in redline.document.iter():
        for child in parent:
            parents[child] = parent

    def ancestors(node: ET.Element) -> set[str]:
        names: set[str] = set()
        while node in parents:
            node = parents[node]
            names.add(local(node.tag))
        return names

    check(
        any(("trPr" in ancestors(node) for node in dated)),
        "redline lacks tracked table-row revisions",
    )
    check(
        any(({"pPr", "rPr"} & ancestors(node) for node in dated)),
        "redline lacks tracked paragraph-mark revisions",
    )
    check(
        any((local(node.tag) == "del" and list(node.iter(W + "delText")) for node in dated))
        and any((local(node.tag) == "ins" and list(node.iter(W + "t")) for node in dated)),
        "redline lacks a word-level tracked replacement",
    )
    check(
        projected_story_shape(redline, accept=True) == story_shape(revised),
        "accepted materialization does not equal revised story",
    )
    check(
        projected_story_shape(redline, accept=False) == story_shape(original),
        "rejected materialization does not equal original story",
    )
    for reference, label in ((revised, "revised"), (original, "original")):
        for name in set(reference.parts) - {"word/document.xml", "docProps/core.xml"}:
            check(equivalent_part(redline.parts.get(name), reference.parts[name]), f"redline changed non-story part from {label}: {name}")

def _write_score(card: ScoreCard) -> None:
    channel = {
        "schema": "docx-grader-score-channel",
        "version": 1,
        "score": round(card.score, 6),
        "grading": card.to_dict(),
    }
    print(json.dumps(channel, sort_keys=True, separators=(",", ":")), flush=True)

def main() -> int:
    if len(sys.argv) != 2 or sys.argv[1] != 'compare-redline-algebra':
        print("usage: grader.py compare-redline-algebra", file=sys.stderr)
        return 2
    task = 'compare-redline-algebra'
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
            lambda: verify_compare_redline_algebra(app),
            effect_patterns=('redline lacks insertion/deletion revisions', 'accepted materialization does not equal revised story', 'rejected materialization does not equal original story'),
            collateral_patterns=('redline changed non-story part',),
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
