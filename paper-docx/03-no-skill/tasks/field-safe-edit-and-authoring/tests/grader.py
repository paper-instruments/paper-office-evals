#!/usr/bin/env python3
"""Standalone Harbor grader for FM10."""

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
EXPECTED_HASHES = {'eval_fixture_02_edit_surgery.docx': '7cdbe138dbc22ff848fc3f05968cb03b9dc4a302a6504bda922006b012344c8e'}
TASK_INPUTS = {'field-safe-edit-and-authoring': ('eval_fixture_02_edit_surgery.docx',)}
TASK_OUTPUTS = {'field-safe-edit-and-authoring': ('evals/field-safe-edit-and-authoring/output.docx',)}
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

def canonical(element: ET.Element) -> tuple[Any, ...]:
    attrs = tuple(sorted(((local(key), value) for key, value in element.attrib.items())))
    return (
        local(element.tag),
        attrs,
        element.text or "",
        tuple((canonical(child) for child in element)),
    )

def instruction_texts(doc: Docx) -> list[str]:
    """Return normalized complex-field instructions, joining split instrText runs."""
    stack: list[dict[str, Any]] = []
    instructions: list[str] = []
    for node in doc.document.iter():
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"parts": [], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                record = stack.pop()
                instructions.append(" ".join("".join(record["parts"]).split()))
        elif node.tag == W + "instrText" and stack and (not stack[-1]["separated"]):
            stack[-1]["parts"].append(node.text or "")
    instructions.extend(
        " ".join((node.get(W + "instr") or "").split())
        for node in doc.document.iter(W + "fldSimple")
    )
    return instructions

def paragraph_with_instruction(doc: Docx, instruction: str) -> ET.Element:
    matches = [
        paragraph
        for paragraph in doc.paragraphs()
        if instruction
        in " ".join("".join((node.text or "" for node in paragraph.iter(W + "instrText"))).split())
    ]
    check(
        len(matches) == 1,
        f"expected one paragraph containing field instruction {instruction!r}, found {len(matches)}",
    )
    return matches[0]

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

def field_result_text(paragraph: ET.Element) -> str:
    simple = [
        "".join(node.text or "" for node in field.iter(W + "t"))
        for field in paragraph.iter(W + "fldSimple")
    ]
    if simple:
        return "".join(simple).strip()
    result: list[str] = []
    inside_result = False
    for node in paragraph.iter():
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "separate":
                inside_result = True
            elif kind == "end":
                inside_result = False
        elif inside_result and node.tag == W + "t":
            result.append(node.text or "")
    return "".join(result).strip()

def verify_field_safe_edit_and_authoring(app: Path) -> None:
    from preservation import field_records, run_normalized_signature, preserves_formatting, paragraph_properties
    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    output = Docx(app / 'evals/field-safe-edit-and-authoring/output.docx')
    source_field = paragraph_with_instruction(source, "REF BatchResult")
    output_field = paragraph_with_instruction(output, "REF BatchResult")
    check(
        element_text(output_field).startswith("Cross-reference result:"),
        "ordinary field-adjacent prose was not updated",
    )
    check("Field result:" not in element_text(output_field), "old ordinary prose prefix remains")
    check(
        field_records(source_field) == field_records(output_field),
        "existing field instruction or cached result changed",
    )
    check(preserves_formatting(source_field, output_field) and paragraph_properties(source_field) == paragraph_properties(output_field), "field-adjacent edit changed existing formatting")
    source_instructions = instruction_texts(source)
    instructions = instruction_texts(output)
    joined = "\n".join(instructions)
    check(
        len(instructions) == len(source_instructions) + 5,
        "expected exactly five newly authored fields",
    )
    slot_expectations = {
        'Current page:': ("PAGE",),
        'Total pages:': ("NUMPAGES",),
        'Report date:': ("DATE", "MMMM d, yyyy"),
        'Defined term reference:': ("REF EvalDependent",),
    }
    excluded = {'Field result:', 'Cross-reference result:', *slot_expectations}
    def unchanged(doc):
        return [run_normalized_signature(p) for p in doc.paragraphs()
                if not any(element_text(p).startswith(prefix) for prefix in excluded)
                and not any('TOC' in (n.text or '') for n in p.iter(W + 'instrText'))]
    check(unchanged(source) == unchanged(output), "field authoring changed unrelated content or formatting")
    for slot, tokens in slot_expectations.items():
        paragraph = output.find_paragraph(slot)
        records = field_records(paragraph)
        check(len(records) == 1 and records[0][0][0] == tokens[0].split()[0], f"{slot} has wrong field type")
        instruction = " ".join(
            (
                "".join((node.text or "" for node in paragraph.iter(W + "instrText")))
                + " "
                + " ".join(node.get(W + "instr") or "" for node in paragraph.iter(W + "fldSimple"))
            ).split()
        )
        for token in tokens:
            check(token in instruction, f"{slot} lacks authored field instruction {token!r}")
        kinds = [node.get(W + "fldCharType") for node in paragraph.iter(W + "fldChar")]
        simple = list(paragraph.iter(W + "fldSimple"))
        check(
            len(simple) == 1
            or kinds in (["begin", "end"], ["begin", "separate", "end"]),
            f"{slot} field shape is neither one simple field nor one balanced complex field",
        )
    body = list(output.body())
    toc_slot = next((i for i, node in enumerate(body) if 'Contents' in element_text(node)), -1)
    check(
        toc_slot >= 0 and toc_slot + 1 < len(body) and (local(body[toc_slot + 1].tag) == "p"),
        "TOC field paragraph is not immediately after its slot",
    )
    toc_instruction = " ".join(
        "".join((node.text or "" for node in body[toc_slot + 1].iter(W + "instrText"))).split()
    )
    check(
        "TOC" in toc_instruction and '\\o "1-3"' in toc_instruction,
        "TOC field is not limited to heading levels 1-3",
    )
    field_records(output.document)
    field_types = [node.get(W + "fldCharType") for node in output.document.iter(W + "fldChar")]
    check(
        field_types.count("begin") == field_types.count("end"),
        "document field markers are unbalanced",
    )
    for token in ("PAGE", "NUMPAGES", "DATE", "REF EvalDependent", "TOC"):
        check(token in joined, f"missing authored field instruction {token!r}")
    settings = output.root("word/settings.xml")
    update = settings.find(".//" + W + "updateFields")
    check(
        update is not None and update.get(W + "val", "true") not in {"0", "false"},
        "updateFields-on-open is not enabled",
    )
    check("batch omega token" in element_text(output_field), "existing cached field result changed")
    changed = semantic_changed_parts(source.parts, output.parts)
    check(
        changed <= {"word/document.xml", "word/settings.xml", "docProps/core.xml"},
        "field authoring exceeded its changed-part budget",
    )

def _write_score(card: ScoreCard) -> None:
    channel = {
        "schema": "docx-grader-score-channel",
        "version": 1,
        "score": round(card.score, 6),
        "grading": card.to_dict(),
    }
    print(json.dumps(channel, sort_keys=True, separators=(",", ":")), flush=True)

def main() -> int:
    if len(sys.argv) != 2 or sys.argv[1] != 'field-safe-edit-and-authoring':
        print("usage: grader.py field-safe-edit-and-authoring", file=sys.stderr)
        return 2
    task = 'field-safe-edit-and-authoring'
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
            lambda: verify_field_safe_edit_and_authoring(app),
            effect_patterns=('ordinary field-adjacent prose was not updated', 'expected exactly five newly authored fields', 'lacks authored field instruction', 'field shape is neither one simple field nor one balanced complex field', 'TOC field paragraph is not immediately after its slot', 'TOC field is not limited to heading levels 1-3', 'document field markers are unbalanced', 'updateFields-on-open is not enabled'),
            collateral_patterns=('existing field instruction or cached result changed', 'field-adjacent edit changed existing formatting', 'existing cached field result changed', 'field authoring exceeded its changed-part budget'),
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
