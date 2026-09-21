"""Package-neutral graders for the refreshed paper-docx tasks."""

from __future__ import annotations

import copy
import hashlib
import json
import os
import posixpath
import stat
import sys
import zipfile
from pathlib import Path
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_R = "{http://schemas.openxmlformats.org/package/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
W15 = "{http://schemas.microsoft.com/office/word/2012/wordml}"

TASKS = {
    "numbering-restart": {
        "inputs": {
            "eval_fixture_03_structured_surfaces.docx": "8a48c765775068997dba55bcf75dd217c6cf2dbdafa91adc2ea9511ada81fd15"
        },
        "output": "evals/numbering-restart/output.docx",
    },
    "rich-block-insertion": {
        "inputs": {
            "eval_fixture_03_structured_surfaces.docx": "10a12ba89c48b5d2bfd54961b3722ce68cec39152cd863f0122c56870b3525f8"
        },
        "output": "evals/rich-block-insertion/output.docx",
    },
    "normalized-context-targeting": {
        "inputs": {
            "eval_fixture_01_perception_gauntlet.docx": "294517c9bee2f9c7efd22ce3465380e7b82267b62fa009020bbf1dea4d9fce1b"
        },
        "output": "evals/normalized-context-targeting/output.docx",
    },
    "data-bound-control-sync": {
        "inputs": {
            "bound-control-source.docx": "eeb036087c71b9898360463920d93337f2a709bfc38a73985f098e651689a2ad"
        },
        "output": "evals/data-bound-control-sync/output.docx",
    },
    "isolated-picture-replacement": {
        "inputs": {
            "eval_fixture_02_edit_surgery.docx": "2dda293583307a149dd12307478f8af11fa0f7e7f8ccc0917efb88f73836debb",
            "replacement.jpeg": "96367138dc44ce09bf2c8f0f8e49348a1478d2c5c0af69bbc2bbc38b63cdcead",
        },
        "output": "evals/isolated-picture-replacement/output.docx",
    },
    "hyperlink-lifecycle": {
        "inputs": {
            "eval_fixture_02_edit_surgery.docx": "6b8b89da13284471fdaf61773a3ff3410b31542a0f6746d414aa323998040935"
        },
        "output": "evals/hyperlink-lifecycle/output.docx",
    },
    "footnote-endnote-authoring": {
        "inputs": {
            "eval_fixture_02_edit_surgery.docx": "64aa3836b61cacc34c16b5e2abed3dd59891c90308512bdecf23eee6dc87174d"
        },
        "output": "evals/footnote-endnote-authoring/output.docx",
    },
    "caption-cross-reference": {
        "inputs": {
            "eval_fixture_02_edit_surgery.docx": "7cdbe138dbc22ff848fc3f05968cb03b9dc4a302a6504bda922006b012344c8e"
        },
        "output": "evals/caption-cross-reference/output.docx",
    },
    "comment-thread-deletion": {
        "inputs": {
            "eval_fixture_02_edit_surgery.docx": "34f7f6c41477bef98bc78f9bd7173120260e856d00ccf5c71c33593a6e6886f9"
        },
        "output": "evals/comment-thread-deletion/output.docx",
    },
    "restrict-editing-authoring": {
        "inputs": {
            "eval_fixture_01_perception_gauntlet.docx": "07920c72f6c5e7986bbd776ea26b5b0baf3f5d3e7d97248116f331b11f00d815"
        },
        "output": "evals/restrict-editing-authoring/output.docx",
    },
    "append-source-letterhead": {
        "inputs": {
            "letterhead-destination.docx": "8d215d7080326a77bcfb1334a7bda91bf32f2ffd97a7b2bcc79eba92896f4518",
            "letterhead-source.docx": "3a34d67db736d7f7a8f0c78a9cfccd9a97985eeb20c0ac5c0d75d23cd7454b03",
        },
        "output": "evals/append-source-letterhead/output.docx",
    },
    "create-review-ready-document": {
        "inputs": {},
        "output": "evals/create-review-ready-document/output.docx",
    },
    "multi-edit-contract-amendment": {
        "inputs": {
            "services-agreement.docx": "5ffa91066339fb6d1f0180f677c9a6d2f2cabfbb5a9e131f3a8803c552c20e81"
        },
        "output": "evals/multi-edit-contract-amendment/output.docx",
    },
    "review-decisions-with-dependent-content": {
        "inputs": {
            "review-decisions.docx": "76fdde252cdd2071597c1d211af0ef436e7c8b8a7faf767fe3e73e1e07d6e33f"
        },
        "output": "evals/review-decisions-with-dependent-content/output.docx",
    },
    "tracked-terminology-update": {
        "inputs": {
            "portal-agreement.docx": "e9f76c165d9dee60b6bb94648ec307a8d06223f4cbb5a23a17d0c2b910e4e385"
        },
        "output": "evals/tracked-terminology-update/output.docx",
    },
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def xml(payload: bytes, name: str) -> ET.Element:
    if len(payload) > 16 * 1024 * 1024:
        raise ValueError(f"XML part too large: {name}")
    probe = payload.lower().replace(b"\x00", b"")
    if b"<!doctype" in probe or b"<!entity" in probe:
        raise ValueError(f"DTD or entity declaration in {name}")
    return ET.fromstring(payload)


def text(node: ET.Element) -> str:
    return "".join(
        item.text or ""
        for item in node.iter()
        if item.tag in {W + "t", W + "delText", W + "instrText"}
    )


def field_instructions(node: ET.Element) -> str:
    chunks = [item.text or "" for item in node.iter(W + "instrText")]
    chunks.extend(item.get(W + "instr", "") for item in node.iter(W + "fldSimple"))
    return "".join(chunks)


class Docx:
    def __init__(self, path: Path):
        if (
            not path.is_file()
            or path.is_symlink()
            or not stat.S_ISREG(path.lstat().st_mode)
        ):
            raise ValueError(f"not a regular DOCX: {path}")
        if path.stat().st_size > 100 * 1024 * 1024:
            raise ValueError("DOCX exceeds size limit")
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 4096:
                raise ValueError("DOCX has too many members")
            expanded = 0
            parts = {}
            for info in infos:
                name = info.filename
                if (
                    not name
                    or name.startswith(("/", "\\"))
                    or "\\" in name
                    or any(
                        piece in {"", ".", ".."}
                        for piece in name.rstrip("/").split("/")
                    )
                ):
                    raise ValueError(f"unsafe DOCX member path: {name!r}")
                expanded += info.file_size
                if info.file_size > 64 * 1024 * 1024 or expanded > 256 * 1024 * 1024:
                    raise ValueError("DOCX expanded-size limit exceeded")
                if not info.is_dir():
                    if name in parts:
                        raise ValueError(f"duplicate DOCX member: {name}")
                    parts[name] = archive.read(info)
            bad = archive.testzip()
            if bad:
                raise ValueError(f"DOCX CRC failure: {bad}")
        for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
            if required not in parts:
                raise ValueError(f"DOCX lacks {required}")
        self.path = path
        self.parts = parts
        self.roots = {
            name: xml(payload, name)
            for name, payload in parts.items()
            if name.endswith((".xml", ".rels"))
        }
        self._validate_relationships()
        self._validate_note_references()
        self._validate_xml_references()

    def _validate_xml_references(self):
        for name, root in self.roots.items():
            if not name.endswith(".xml"):
                continue
            attrs = [
                value
                for node in root.iter()
                for key, value in node.attrib.items()
                if key in {R + "id", R + "embed", R + "link"}
            ]
            if not attrs:
                continue
            rel_name = posixpath.join(
                posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels"
            )
            relations = rels(self, rel_name)
            if any(value not in relations for value in attrs):
                raise ValueError(f"unresolved XML relationship in {name}")

    def _validate_note_references(self):
        for kind in ("footnote", "endnote"):
            refs = [
                node
                for name, root in self.roots.items()
                if name.startswith("word/")
                for node in root.iter(W + kind + "Reference")
            ]
            if not refs:
                continue
            notes = self.root(f"word/{kind}s.xml").findall(W + kind)
            ids = [node.get(W + "id") for node in notes]
            if len(ids) != len(set(ids)) or any(
                ref.get(W + "id") not in ids for ref in refs
            ):
                raise ValueError(f"unresolved or duplicate {kind} identity")

    def _validate_relationships(self) -> None:
        for rels_name, root in self.roots.items():
            if not rels_name.endswith(".rels"):
                continue
            if rels_name == "_rels/.rels":
                source = ""
            else:
                prefix, leaf = rels_name.rsplit("/_rels/", 1)
                source = posixpath.join(prefix, leaf[:-5])
            for rel in root.findall("./" + PKG_R + "Relationship"):
                target = rel.get("Target", "")
                if rel.get("TargetMode") == "External":
                    if urlsplit(target).scheme.casefold() not in {
                        "http",
                        "https",
                        "mailto",
                    }:
                        raise ValueError(f"unsafe external relationship: {target!r}")
                    continue
                resolved = posixpath.normpath(
                    posixpath.join(posixpath.dirname(source), target.split("#", 1)[0])
                ).lstrip("/")
                if resolved not in self.parts:
                    raise ValueError(f"dangling relationship {rels_name}: {target!r}")

    def root(self, name: str) -> ET.Element:
        if name not in self.roots:
            raise ValueError(f"DOCX lacks XML part {name}")
        return self.roots[name]

    def all_text(self) -> str:
        return "\n".join(
            text(root) for name, root in self.roots.items() if name.endswith(".xml")
        )


def paragraphs(doc: Docx, part: str = "word/document.xml") -> list[ET.Element]:
    return list(doc.root(part).iter(W + "p"))


def paragraph(doc: Docx, needle: str, part: str = "word/document.xml") -> ET.Element:
    matches = [item for item in paragraphs(doc, part) if needle in text(item)]
    if len(matches) != 1:
        raise ValueError(
            f"expected one paragraph containing {needle!r}, found {len(matches)}"
        )
    return matches[0]


def changed_parts(before, after):
    from preservation import changed_parts as semantic_changes

    return semantic_changes(before.parts, after.parts)


def rels(
    doc: Docx, name: str = "word/_rels/document.xml.rels"
) -> dict[str, ET.Element]:
    return {
        item.get("Id", ""): item
        for item in doc.root(name).findall("./" + PKG_R + "Relationship")
    }


def num_id(item: ET.Element) -> str | None:
    node = item.find("./" + W + "pPr/" + W + "numPr/" + W + "numId")
    return None if node is None else node.get(W + "val")


def signature(node):
    from preservation import run_normalized_signature

    return run_normalized_signature(node)


def visible(node):
    return "".join(item.text or "" for item in node.iter(W + "t"))


def body(doc):
    return list(doc.root("word/document.xml").find(W + "body"))


def style(node):
    prop = node.find("./" + W + "pPr/" + W + "pStyle")
    return None if prop is None else prop.get(W + "val")


def note_at(doc, phrase, kind, expected):
    node = paragraph(doc, phrase)
    refs = list(node.iter(W + kind + "Reference"))
    if len(refs) != 1:
        return False
    ident = refs[0].get(W + "id")
    notes = [
        n
        for n in doc.root(f"word/{kind}s.xml").findall(W + kind)
        if n.get(W + "id") == ident
    ]
    prefix = ""
    for item in node.iter():
        if item is refs[0]:
            break
        if item.tag == W + "t":
            prefix += item.text or ""
    return (
        len(notes) == 1
        and visible(notes[0]).strip() == expected
        and prefix.endswith(phrase)
    )


def comment_anchor(doc, ident):
    active, chunks, starts, ends, refs = False, [], 0, 0, 0
    for node in doc.root("word/document.xml").iter():
        if node.tag == W + "commentRangeStart" and node.get(W + "id") == ident:
            starts += 1
            active = True
        elif node.tag == W + "commentRangeEnd" and node.get(W + "id") == ident:
            ends += 1
            active = False
        elif node.tag == W + "commentReference" and node.get(W + "id") == ident:
            refs += 1
        elif active and node.tag == W + "t":
            chunks.append(node.text or "")
    return "".join(chunks) if (starts, ends, refs) == (1, 1, 1) else None


def active_stories(doc, kind):
    relations = rels(doc)
    roots = []
    for section in doc.root("word/document.xml").iter(W + "sectPr"):
        for ref in section.findall(W + kind + "Reference"):
            relation = relations.get(ref.get(R + "id"))
            if relation is None or not relation.get("Type", "").endswith("/" + kind):
                raise ValueError(f"unresolved {kind} reference")
            target = posixpath.normpath(
                posixpath.join("word", relation.get("Target", ""))
            )
            roots.append(doc.root(target))
    return roots


def formatted_chars(node):
    return [
        (char, signature(run.find(W + "rPr")))
        for run in node.iter(W + "r")
        for char in visible(run)
    ]


def unchanged_paragraphs(source, output, excluded):
    before = [p for p in paragraphs(source) if visible(p) not in excluded]
    after = [p for p in paragraphs(output) if visible(p) not in excluded]
    return [signature(p) for p in before] == [signature(p) for p in after]


def numbering_format(doc: Docx, item: ET.Element) -> tuple[str, str] | None:
    ident = num_id(item)
    level = item.find("./" + W + "pPr/" + W + "numPr/" + W + "ilvl")
    level_value = "0" if level is None else level.get(W + "val", "0")
    if ident is None:
        return None
    root = doc.root("word/numbering.xml")
    num = next(
        (
            node
            for node in root.findall("./" + W + "num")
            if node.get(W + "numId") == ident
        ),
        None,
    )
    if num is None:
        return None
    ref = num.find("./" + W + "abstractNumId")
    abstract_id = None if ref is None else ref.get(W + "val")
    abstract = next(
        (
            node
            for node in root.findall("./" + W + "abstractNum")
            if node.get(W + "abstractNumId") == abstract_id
        ),
        None,
    )
    if abstract is None:
        return None
    level_node = next(
        (
            node
            for node in abstract.findall("./" + W + "lvl")
            if node.get(W + "ilvl") == level_value
        ),
        None,
    )
    fmt = None if level_node is None else level_node.find("./" + W + "numFmt")
    return (level_value, "" if fmt is None else fmt.get(W + "val", ""))


class Card:
    def __init__(self, task: str):
        self.task = task
        self.checks: list[dict] = []
        self.hard_failures: list[str] = []

    def check(self, condition: bool, message: str) -> None:
        self.checks.append({"name": message, "passed": bool(condition), "weight": 1.0})

    @property
    def score(self) -> float:
        return float(
            not self.hard_failures
            and bool(self.checks)
            and all(item["passed"] for item in self.checks)
        )

    def payload(self) -> dict:
        return {
            "schema": "docx-harbor-grade",
            "version": 2,
            "task": self.task,
            "score": round(self.score, 6),
            "full_pass": self.score == 1.0,
            "hard_failures": self.hard_failures,
            "checks": self.checks,
        }


def verify_numbering_restart(app: Path, output: Docx, card: Card) -> None:
    source = Docx(app / "eval_fixtures/docx/eval_fixture_03_structured_surfaces.docx")
    root = output.root("word/numbering.xml")
    counters, values = {}, {}
    for p in paragraphs(output):
        ident = num_id(p)
        if ident is None:
            continue
        definition = next(
            n for n in root.findall(W + "num") if n.get(W + "numId") == ident
        )
        abstract_id = definition.find(W + "abstractNumId").get(W + "val")
        abstract = next(
            n
            for n in root.findall(W + "abstractNum")
            if n.get(W + "abstractNumId") == abstract_id
        )
        level = next(n for n in abstract.findall(W + "lvl") if n.get(W + "ilvl") == "0")
        start = level.find(W + "start")
        override = definition.find("./" + W + "lvlOverride/" + W + "startOverride")
        initial = (
            int(override.get(W + "val"))
            if override is not None
            else int(start.get(W + "val", "1"))
            if start is not None
            else 1
        )
        counters[ident] = counters.get(ident, initial - 1) + 1
        values[visible(p)] = counters[ident]
    for label, expected in (
        ("Record the request.", 1),
        ("Review the budget.", 2),
        ("Approve the supplier.", 3),
        ("Confirm supplier details.", 1),
    ):
        card.check(
            values.get(label) == expected,
            f"{label} displays automatic number {expected}",
        )
        card.check(
            numbering_format(output, paragraph(output, label)) == ("0", "decimal"),
            f"{label} retains its decimal-list format",
        )
        item = paragraph(output, label)
        definition = next(
            n for n in root.findall(W + "num") if n.get(W + "numId") == num_id(item)
        )
        abstract_id = definition.find(W + "abstractNumId").get(W + "val")
        abstract = next(
            n
            for n in root.findall(W + "abstractNum")
            if n.get(W + "abstractNumId") == abstract_id
        )
        effective = definition.find(W + "lvlOverride/" + W + "lvl")
        if effective is None:
            effective = next(
                n for n in abstract.findall(W + "lvl") if n.get(W + "ilvl") == "0"
            )
        pattern = effective.find(W + "lvlText")
        card.check(
            pattern is not None and pattern.get(W + "val") == "%1.",
            f"{label} uses the original dynamic decimal label",
        )
    normalized = []
    for doc in (source, output):
        node = copy.deepcopy(doc.root("word/document.xml"))
        for parent in node.iter():
            for child in list(parent):
                if child.tag == W + "numPr":
                    parent.remove(child)
        normalized.append(signature(node))
    card.check(
        normalized[0] == normalized[1],
        "restart preserves paragraph text, formatting and structure",
    )
    card.check(
        changed_parts(source, output)
        <= {"word/document.xml", "word/numbering.xml", "docProps/core.xml"},
        "numbering edit has bounded package changes",
    )


def verify_rich_block_insertion(app: Path, output: Docx, card: Card) -> None:
    source = Docx(app / "eval_fixtures/docx/eval_fixture_03_structured_surfaces.docx")
    blocks = body(output)
    anchor = next(
        i for i, item in enumerate(blocks) if "Implementation plan" in text(item)
    )
    inserted = blocks[anchor + 1 : anchor + 7]
    card.check(
        [text(item) for item in inserted]
        == [
            "Implementation overview",
            "Confirm the scope",
            "Assign the owners",
            "Legal review",
            "Finance review",
            "MetricValueCoverageComplete",
        ],
        "report additions appear in the requested order",
    )
    heading = inserted[0]
    table = inserted[-1]
    card.check(
        table.tag == W + "tbl"
        and [
            [visible(c) for c in row.findall(W + "tc")]
            for row in table.findall(W + "tr")
        ]
        == [["Metric", "Value"], ["Coverage", "Complete"]],
        "metrics are a native two-row two-column table",
    )
    style = heading.find("./" + W + "pPr/" + W + "pStyle")
    card.check(
        style is not None and style.get(W + "val") == "Heading3",
        "rich paragraph uses Heading 3",
    )
    runs = list(heading.iter(W + "r"))
    bold_text = "".join(
        visible(run)
        for run in runs
        if run.find("./" + W + "rPr/" + W + "b") is not None
    )
    italic_text = "".join(
        visible(run)
        for run in runs
        if run.find("./" + W + "rPr/" + W + "i") is not None
    )
    card.check("Implementation" in bold_text, "Implementation is bold")
    card.check("overview" in italic_text, "overview is italic")
    card.check(
        all(
            numbering_format(output, item) == ("0", "decimal") for item in inserted[1:3]
        ),
        "decimal list uses real numbering",
    )
    card.check(
        all(
            numbering_format(output, item) == ("1", "bullet") for item in inserted[3:5]
        ),
        "nested list uses real bullet numbering",
    )
    card.check(
        num_id(paragraph(source, "Existing checklist item"))
        == num_id(paragraph(output, "Existing checklist item")),
        "existing numbering is unchanged",
    )
    card.check(
        [signature(item) for item in blocks[: anchor + 1] + blocks[anchor + 7 :]]
        == [signature(item) for item in body(source)],
        "existing report blocks are preserved",
    )


def verify_normalized_context_targeting(app: Path, output: Docx, card: Card) -> None:
    source = Docx(app / "eval_fixtures/docx/eval_fixture_01_perception_gauntlet.docx")
    target = paragraph(output, "Northstar Consulting UK")
    card.check(
        visible(target) == "Supplier: Northstar Consulting UK",
        "current supplier name is updated",
    )
    blocks = body(output)
    heading = next(i for i, p in enumerate(blocks) if visible(p) == "Current supplier")
    card.check(
        blocks[heading + 1] is target,
        "replacement belongs to the current supplier section",
    )
    card.check(
        unchanged_paragraphs(
            source,
            output,
            {
                "Supplier: NORTHSTAR\u00a0SERVICES\u2014UK",
                "Supplier: Northstar Consulting UK",
            },
        ),
        "archived and US supplier records and non-target content are preserved",
    )
    card.check(
        changed_parts(source, output) <= {"word/document.xml", "docProps/core.xml"},
        "supplier update preserves other package parts",
    )


def verify_data_bound_control_sync(app: Path, output: Docx, card: Card) -> None:
    from preservation import (
        binding_address,
        bound_store,
        equivalent_part,
        remapped_parts,
        without,
    )
    from preservation import changed_parts as semantic_changes

    source = Docx(app / "eval_fixtures/docx/bound-control-source.docx")
    card.check(
        signature(without(source.root("word/document.xml"), {"sdt"}))
        == signature(without(output.root("word/document.xml"), {"sdt"})),
        "bound form preserves unrelated content and formatting",
    )
    control = next(
        (
            node
            for node in output.root("word/document.xml").iter(W + "sdt")
            if any(tag.get(W + "val") == "company-name" for tag in node.iter(W + "tag"))
        ),
        None,
    )
    card.check(
        control is not None and "Northstar Holdings" in text(control),
        "bound control display is updated",
    )
    binding = (
        None
        if control is None
        else control.find("./" + W + "sdtPr/" + W + "dataBinding")
    )
    original_control = next(
        node
        for node in source.root("word/document.xml").iter(W + "sdt")
        if any(tag.get(W + "val") == "company-name" for tag in node.iter(W + "tag"))
    )
    original_binding = original_control.find("./" + W + "sdtPr/" + W + "dataBinding")
    card.check(
        binding is not None
        and original_binding is not None
        and binding_address(binding) == binding_address(original_binding),
        "binding continues to address the original store and field",
    )
    store_name, store, props_name = bound_store(
        output.parts, original_binding.get(W + "storeItemID")
    )
    source_name, original_store, original_props = bound_store(
        source.parts, original_binding.get(W + "storeItemID")
    )
    expected_store = copy.deepcopy(original_store)
    ns = "{http://example.com/form}"
    expected_store.find(ns + "name").text = "Northstar Holdings"
    expected_store.find(ns + "client").set("name", "Northstar Holdings")
    card.check(
        signature(store) == signature(expected_store),
        "linked store retains all unrequested data",
    )
    card.check(
        "Northstar Holdings" in "".join(store.itertext()),
        "custom XML store is synchronized",
    )
    ns = "{http://example.com/form}"
    for tag, expected in (
        ("company-name", "Northstar Holdings"),
        ("company-legal-name", "Northstar Holdings"),
        ("account-code", "NW-204"),
    ):
        original = next(
            n
            for n in source.root("word/document.xml").iter(W + "sdt")
            if any(t.get(W + "val") == tag for t in n.iter(W + "tag"))
        )
        candidates = [
            n
            for n in output.root("word/document.xml").iter(W + "sdt")
            if any(t.get(W + "val") == tag for t in n.iter(W + "tag"))
        ]
        card.check(
            len(candidates) == 1 and visible(candidates[0]) == expected,
            f"{tag} has the correct display value",
        )
        if len(candidates) != 1:
            continue
        binding = candidates[0].find("./" + W + "sdtPr/" + W + "dataBinding")
        card.check(
            binding is not None
            and binding_address(binding)
            == binding_address(original.find("./" + W + "sdtPr/" + W + "dataBinding")),
            f"{tag} retains its namespace, store identity and XPath",
        )
        value = (
            store.find(ns + "name").text
            if tag == "company-name"
            else store.find(ns + "client").get(
                "name" if tag == "company-legal-name" else "code"
            )
        )
        card.check(
            value == expected, f"{tag} resolves to the correct linked data value"
        )
    card.check(
        equivalent_part(source.parts[original_props], output.parts.get(props_name)),
        "custom XML store identity is unchanged",
    )
    card.check(
        semantic_changes(
            remapped_parts(source.parts, {}),
            remapped_parts(
                output.parts, {store_name: source_name, props_name: original_props}
            ),
        )
        <= {
            "word/document.xml",
            source_name,
            "docProps/core.xml",
        },
        "bound-control update has bounded package changes",
    )


def verify_isolated_picture_replacement(app: Path, output: Docx, card: Card) -> None:
    from preservation import without

    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    card.check(
        signature(without(source.root("word/document.xml"), {"drawing"}))
        == signature(without(output.root("word/document.xml"), {"drawing"})),
        "picture replacement preserves surrounding body content and formatting",
    )
    replacement = (app / "eval_fixtures/docx/replacement.jpeg").read_bytes()
    source_p = paragraph(source, "Current supplier badge:")
    output_p = paragraph(output, "Current supplier badge:")
    source_drawing = next(source_p.iter(W + "drawing"))
    output_drawing = next(output_p.iter(W + "drawing"))
    source_extent = source_drawing.find(".//" + WP + "extent")
    output_extent = output_drawing.find(".//" + WP + "extent")
    card.check(
        source_extent is not None
        and output_extent is not None
        and source_extent.attrib == output_extent.attrib,
        "picture display size is preserved",
    )
    blip = output_drawing.find(".//" + A + "blip")
    relation = None if blip is None else rels(output).get(blip.get(R + "embed", ""))
    target = (
        None
        if relation is None
        else posixpath.normpath(posixpath.join("word", relation.get("Target", "")))
    )
    card.check(
        target in output.parts and output.parts[target] == replacement,
        "target drawing uses the replacement image bytes",
    )
    for doc in (source, output):
        archived = paragraph(doc, "Archived supplier badge:")
        archive_blip = next(archived.iter(A + "blip"))
        archive_rel = rels(doc)[archive_blip.get(R + "embed")]
        archive_path = posixpath.normpath(
            posixpath.join("word", archive_rel.get("Target"))
        )
        if doc is source:
            archive_bytes = doc.parts[archive_path]
            archive_shape = signature(next(archived.iter(W + "drawing")))
        else:
            card.check(
                doc.parts[archive_path] == archive_bytes,
                "archived drawing still resolves to the original image",
            )
            card.check(
                signature(next(archived.iter(W + "drawing"))) == archive_shape,
                "archived drawing placement and identity are preserved",
            )
    source_geometry, output_geometry = (
        copy.deepcopy(source_drawing),
        copy.deepcopy(output_drawing),
    )
    for drawing in (source_geometry, output_geometry):
        for node in drawing.iter(A + "blip"):
            node.attrib.pop(R + "embed", None)
    card.check(
        signature(source_geometry) == signature(output_geometry),
        "target drawing geometry and placement are preserved",
    )
    for name in source.parts:
        if name.startswith("word/media/"):
            card.check(
                output.parts.get(name) == source.parts[name],
                f"existing media remains unchanged: {name}",
            )
    allowed = {"word/document.xml", "word/_rels/document.xml.rels", "docProps/core.xml"}
    card.check(
        all(
            name in allowed or name.startswith("word/media/")
            for name in changed_parts(source, output)
        ),
        "picture replacement has bounded package changes",
    )


def verify_hyperlink_lifecycle(app: Path, output: Docx, card: Card) -> None:
    from preservation import paragraph_properties

    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    card.check(
        paragraph_properties(source.root("word/document.xml"))
        == paragraph_properties(output.root("word/document.xml")),
        "link edit preserves paragraph formatting",
    )
    for label, url in (
        ("current travel policy", "https://example.org/policies/travel-2026"),
        ("archived travel policy", "https://example.org/policies/travel-2025"),
    ):
        links = [
            n
            for n in paragraph(output, label).iter(W + "hyperlink")
            if visible(n) == label
        ]
        relation = rels(output).get(links[0].get(R + "id")) if len(links) == 1 else None
        card.check(
            relation is not None
            and relation.get("TargetMode") == "External"
            and relation.get("Target") == url,
            f"{label} resolves to its requested URL",
        )
    card.check(
        [visible(p) for p in paragraphs(source)]
        == [visible(p) for p in paragraphs(output)],
        "all handbook labels and prose are unchanged",
    )
    card.check(
        formatted_chars(source.root("word/document.xml"))
        == formatted_chars(output.root("word/document.xml")),
        "link retargeting preserves text formatting",
    )


def verify_footnote_endnote_authoring(app: Path, output: Docx, card: Card) -> None:
    from preservation import reference_positions

    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    for tag in ("bookmarkStart", "bookmarkEnd"):
        card.check(
            reference_positions(source.root("word/document.xml"), tag)
            == reference_positions(output.root("word/document.xml"), tag),
            "existing bookmark anchors are preserved",
        )
    foot_refs = list(
        paragraph(output, "renewal approval").iter(W + "footnoteReference")
    )
    end_refs = list(
        paragraph(output, "notice requirements").iter(W + "endnoteReference")
    )
    card.check(
        len(foot_refs) == 1 and foot_refs[0].get(W + "id") not in {None, "-1", "0"},
        "footnote reference is anchored to the requested text",
    )
    card.check(
        len(end_refs) == 1 and end_refs[0].get(W + "id") not in {None, "-1", "0"},
        "endnote reference is anchored to the requested text",
    )
    card.check(
        "Finance approved the renewal on July 10, 2026."
        in text(output.root("word/footnotes.xml")),
        "footnote body is authored",
    )
    card.check(
        "Source: the signed services agreement."
        in text(output.root("word/endnotes.xml")),
        "endnote body is authored",
    )
    card.check(
        note_at(
            output,
            "renewal approval",
            "footnote",
            "Finance approved the renewal on July 10, 2026.",
        ),
        "footnote resolves at the requested phrase",
    )
    card.check(
        note_at(
            output,
            "notice requirements",
            "endnote",
            "Source: the signed services agreement.",
        ),
        "endnote resolves at the requested phrase",
    )
    card.check(
        note_at(
            output,
            "January 2025",
            "footnote",
            "The original agreement remains on file.",
        ),
        "existing footnote retains its anchor and text",
    )
    card.check(
        note_at(
            output, "Original agreement", "endnote", "Source: signed agreement archive."
        ),
        "existing endnote retains its anchor and text",
    )
    allowed = {
        "[Content_Types].xml",
        "word/document.xml",
        "word/_rels/document.xml.rels",
        "word/footnotes.xml",
        "word/endnotes.xml",
        "docProps/core.xml",
    }
    card.check(
        changed_parts(source, output) <= allowed,
        "note authoring has bounded package changes",
    )


def verify_caption_cross_reference(app: Path, output: Docx, card: Card) -> None:
    from preservation import field_records

    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    caption = paragraph(output, "Supplier badge")
    style = caption.find("./" + W + "pPr/" + W + "pStyle")
    instructions = field_instructions(caption)
    card.check(
        any(tokens[:2] == ("SEQ", "Figure") for tokens, _ in field_records(caption)),
        "caption has a complete live SEQ field",
    )
    card.check(
        style is not None and style.get(W + "val") == "Caption",
        "caption uses the Caption style",
    )
    card.check(
        "SEQ Figure" in instructions and "Supplier badge" in text(caption),
        "caption contains a live Figure sequence field",
    )
    blocks = body(output)
    picture = paragraph(output, "Current supplier badge:")
    card.check(
        blocks.index(caption) == blocks.index(picture) + 1,
        "caption immediately follows the figure",
    )
    for node in (caption, paragraph(output, "See dependent term:")):
        stack = []
        valid = True
        for field in node.iter(W + "fldChar"):
            kind = field.get(W + "fldCharType")
            if kind == "begin":
                stack.append(False)
            elif kind == "separate":
                valid &= bool(stack) and not stack[-1]
                if stack:
                    stack[-1] = True
            elif kind == "end":
                valid &= bool(stack)
                if stack:
                    stack.pop()
        card.check(valid and not stack, "authored field boundaries are balanced")
    reference = paragraph(output, "See dependent term:")
    card.check(
        any(
            tokens[:2] == ("REF", "EvalDependent")
            for tokens, _ in field_records(reference)
        ),
        "reference has a complete live REF field",
    )
    ref_instruction = field_instructions(reference)
    card.check(
        "REF EvalDependent" in ref_instruction,
        "cross-reference contains a live REF field",
    )
    card.check(
        any(
            node.get(W + "name") == "EvalDependent"
            for node in output.root("word/document.xml").iter(W + "bookmarkStart")
        ),
        "referenced bookmark is preserved",
    )
    update = output.root("word/settings.xml").find(".//" + W + "updateFields")
    card.check(
        update is not None and update.get(W + "val", "true") in {"1", "true", "on"},
        "fields are marked for update on open",
    )
    card.check(
        changed_parts(source, output)
        <= {"word/document.xml", "word/settings.xml", "docProps/core.xml"},
        "field authoring has bounded package changes",
    )


def verify_comment_thread_deletion(app: Path, output: Docx, card: Card) -> None:
    from preservation import modern_comment_graph, without

    source = Docx(app / "eval_fixtures/docx/eval_fixture_02_edit_surgery.docx")
    markers = {"commentRangeStart", "commentRangeEnd", "commentReference"}
    card.check(
        signature(without(source.root("word/document.xml"), markers))
        == signature(without(output.root("word/document.xml"), markers)),
        "comment deletion preserves body formatting and structure",
    )
    original = source.root("word/comments.xml").findall(W + "comment")
    keep = [
        n
        for n in original
        if visible(n)
        in {"Legal review is still pending.", "Please keep the existing cap."}
    ]
    removed = [n for n in original if n not in keep]
    comments = output.root("word/comments.xml").findall(W + "comment")
    card.check(
        [signature(n) for n in comments] == [signature(n) for n in keep],
        "pricing thread is deleted and the legal thread remains intact",
    )
    ids = {n.get(W + "id") for n in removed}
    markers = {W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference"}
    card.check(
        not any(
            n.tag in markers and n.get(W + "id") in ids
            for n in output.root("word/document.xml").iter()
        ),
        "pricing thread anchors are removed",
    )
    card.check(
        all(
            comment_anchor(output, n.get(W + "id"))
            == comment_anchor(source, n.get(W + "id"))
            for n in keep
        ),
        "legal comments retain their exact anchors",
    )
    card.check(
        [visible(p) for p in paragraphs(source)]
        == [visible(p) for p in paragraphs(output)],
        "agreement prose is unchanged",
    )
    removed_paras = {p.get(W14 + "paraId") for n in removed for p in n.iter(W + "p")}
    entries = output.root("word/commentsExtended.xml").findall(W15 + "commentEx")
    card.check(
        not any(
            n.get(W15 + "paraId") in removed_paras
            or n.get(W15 + "paraIdParent") in removed_paras
            for n in entries
        ),
        "deleted thread has no extension rows or parent links",
    )
    expected_entries = [
        n
        for n in source.root("word/commentsExtended.xml").findall(W15 + "commentEx")
        if n.get(W15 + "paraId") not in removed_paras
    ]
    card.check(
        [signature(n) for n in entries] == [signature(n) for n in expected_entries],
        "legal reply links and resolution state are preserved",
    )
    card.check(
        modern_comment_graph(output.parts),
        "modern comment identities contain only the surviving comments",
    )
    allowed = {
        "word/document.xml",
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/people.xml",
        "word/commentsIds.xml",
        "word/commentsExtensible.xml",
        "docProps/core.xml",
    }
    card.check(
        changed_parts(source, output) <= allowed,
        "comment deletion has bounded package changes",
    )


def verify_restrict_editing_authoring(app: Path, output: Docx, card: Card) -> None:
    source = Docx(app / "eval_fixtures/docx/eval_fixture_01_perception_gauntlet.docx")
    protection = output.root("word/settings.xml").find(".//" + W + "documentProtection")
    card.check(
        protection is not None and protection.get(W + "edit") == "trackedChanges",
        "Restrict Editing uses trackedChanges mode",
    )
    card.check(
        protection is not None
        and protection.get(W + "enforcement") in {"1", "true", "on"},
        "Restrict Editing is enforced",
    )
    card.check(
        changed_parts(source, output) == {"word/settings.xml"},
        "protection authoring changes only settings.xml",
    )


def verify_append_source_letterhead(app: Path, output: Docx, card: Card) -> None:
    source = Docx(app / "eval_fixtures/docx/letterhead-source.docx")
    destination = Docx(app / "eval_fixtures/docx/letterhead-destination.docx")
    body_text = text(output.root("word/document.xml"))
    card.check(
        "DESTINATION OPENING" in body_text
        and "SOURCE APPEND BODY" in body_text
        and body_text.index("DESTINATION OPENING")
        < body_text.index("SOURCE APPEND BODY"),
        "source body is appended after destination content",
    )
    headers = "\n".join(text(root) for root in active_stories(output, "header"))
    footers = "\n".join(text(root) for root in active_stories(output, "footer"))
    card.check(
        "NORTHSTAR LEGAL | CONFIDENTIAL" in headers
        and "OLD DESTINATION HEADER" not in headers,
        "source header replaces destination letterhead",
    )
    card.check("Northstar Legal" in footers, "source footer is imported")
    card.check(
        source.all_text().count("SOURCE APPEND BODY") == 1
        and destination.all_text().count("DESTINATION OPENING") == 1,
        "input fixtures remain semantically intact",
    )
    blocks = body(output)
    start = next(i for i, p in enumerate(blocks) if "SOURCE APPEND BODY" in text(p))
    end = next(i for i, p in enumerate(blocks) if "DESTINATION OPENING" in text(p))
    boundary = blocks[end : start + 1]
    new_page = any(
        n.tag == W + "br"
        and n.get(W + "type") == "page"
        or n.tag == W + "pageBreakBefore"
        for p in boundary
        for n in p.iter()
    )
    for p in blocks[end:start]:
        for section in p.iter(W + "sectPr"):
            kind = section.find(W + "type")
            new_page |= kind is None or kind.get(W + "val") in {
                "nextPage",
                "oddPage",
                "evenPage",
            }
    card.check(new_page, "source body begins at a page boundary")
    source_paras = [visible(p) for p in paragraphs(source)]
    output_paras = [visible(p) for p in paragraphs(output)]
    offset = output_paras.index(source_paras[0])
    card.check(
        output_paras[offset:] == source_paras, "source body is complete and ordered"
    )
    card.check(
        any(
            "PAGE" in field_instructions(root)
            for root in active_stories(output, "footer")
        ),
        "active source footer retains its page field",
    )
    card.check(
        any(
            list(root.iter(W + "drawing")) for root in active_stories(output, "header")
        ),
        "active source header retains its logo",
    )
    card.check(
        "Legal office" in headers,
        "active source header retains the legal-office link label",
    )
    sections = []
    inherited = {}
    for original in output.root("word/document.xml").iter(W + "sectPr"):
        section = copy.deepcopy(original)
        for kind in ("header", "footer"):
            for ref in section.findall(W + kind + "Reference"):
                inherited[(kind, ref.get(W + "type"))] = copy.deepcopy(ref)
            for variant in ("first", "default", "even"):
                if (
                    not any(
                        n.get(W + "type") == variant
                        for n in section.findall(W + kind + "Reference")
                    )
                    and (kind, variant) in inherited
                ):
                    section.insert(0, copy.deepcopy(inherited[(kind, variant)]))
        sections.append(section)
    section = sections[0]

    def enabled(node):
        return node is not None and node.get(W + "val", "1") not in {
            "0",
            "false",
            "off",
        }

    card.check(enabled(section.find(W + "titlePg")), "first-page letterhead is enabled")
    card.check(
        enabled(output.root("word/settings.xml").find(W + "evenAndOddHeaders")),
        "even-page running headers are enabled",
    )
    for section in sections:
        for kind in ("header", "footer"):
            source_section = list(source.root("word/document.xml").iter(W + "sectPr"))[
                -1
            ]
            for variant in ("first", "default", "even"):
                if variant == "first" and not enabled(section.find(W + "titlePg")):
                    continue

                def story(doc, sect, kind=kind, variant=variant):
                    reference = next(
                        n
                        for n in sect.findall(W + kind + "Reference")
                        if n.get(W + "type") == variant
                    )
                    relationship = rels(doc, "word/_rels/document.xml.rels")[
                        reference.get(R + "id")
                    ]
                    name = posixpath.normpath(
                        posixpath.join("word", relationship.get("Target"))
                    ).lstrip("/")
                    return name, doc.root(name)

                source_name, source_story = story(source, source_section)
                output_name, output_story = story(output, section)
                from preservation import field_records

                card.check(
                    field_records(output_story) == field_records(source_story),
                    f"{variant} {kind} has complete live fields",
                )
                card.check(
                    formatted_chars(output_story) == formatted_chars(source_story),
                    f"{variant} {kind} retains its text and direct formatting",
                )
                card.check(
                    field_instructions(output_story)
                    == field_instructions(source_story),
                    f"{variant} {kind} retains its fields",
                )
                for tag, attr in (("drawing", R + "embed"), ("hyperlink", R + "id")):

                    def dependencies(doc, name, root, tag=tag, attr=attr):
                        rel_name = posixpath.join(
                            posixpath.dirname(name),
                            "_rels",
                            posixpath.basename(name) + ".rels",
                        )
                        relationships = (
                            rels(doc, rel_name) if rel_name in doc.parts else {}
                        )
                        values = []
                        for node in root.iter(W + tag):
                            for child in node.iter():
                                ident = child.get(attr)
                                if ident:
                                    rel = relationships[ident]
                                    target = rel.get("Target")
                                    values.append(
                                        target
                                        if rel.get("TargetMode") == "External"
                                        else doc.parts[
                                            posixpath.normpath(
                                                posixpath.join(
                                                    posixpath.dirname(name), target
                                                )
                                            ).lstrip("/")
                                        ]
                                    )
                        return values

                    card.check(
                        dependencies(output, output_name, output_story)
                        == dependencies(source, source_name, source_story),
                        f"{variant} {kind} retains {tag} dependencies",
                    )
                if kind == "header" and variant != "first":
                    ident = style(next(output_story.iter(W + "p")))
                    definition = next(
                        (
                            s
                            for s in output.root("word/styles.xml").iter(W + "style")
                            if s.get(W + "styleId") == ident
                        ),
                        None,
                    )
                    size = (
                        None
                        if definition is None
                        else definition.find(W + "rPr/" + W + "sz")
                    )
                    card.check(
                        size is not None and size.get(W + "val") == "18",
                        f"{variant} header retains its 9-point custom style",
                    )


def revision_projection(root, decisions):
    """Project only selected review decisions without importing either treatment."""
    root = copy.deepcopy(root)

    def visit(parent):
        for child in list(parent):
            if child.tag == W + "p":
                mark = child.find(W + "pPr/" + W + "rPr/" + W + "del")
                if (
                    mark is not None
                    and decisions.get(mark.get(W + "author")) == "accept"
                ):
                    parent.remove(child)
                    continue
            decision = decisions.get(child.get(W + "author"))
            if child.tag in {W + "ins", W + "del"} and decision:
                keep = (child.tag == W + "ins") == (decision == "accept")
                offset = list(parent).index(child)
                parent.remove(child)
                if keep:
                    visit(child)
                    for item in list(child):
                        for node in item.iter(W + "delText"):
                            node.tag = W + "t"
                        parent.insert(offset, item)
                        offset += 1
            else:
                visit(child)

    visit(root)
    return root


def current_paragraphs(root):
    from preservation import live_text

    return [live_text(p) for p in root.iter(W + "p") if live_text(p)]


def verify_review_decisions_with_dependent_content(
    app: Path, output: Docx, card: Card
) -> None:
    from preservation import modern_comment_graph, reference_positions

    source = Docx(app / "eval_fixtures/docx/review-decisions.docx")
    expected = revision_projection(
        source.root("word/document.xml"),
        {"Morgan Finance": "reject", "Alice Editor": "accept"},
    )
    actual = output.root("word/document.xml")
    card.check(
        reference_positions(actual, "footnoteReference")
        == reference_positions(expected, "footnoteReference"),
        "surviving footnotes retain exact text anchors",
    )
    card.check(
        current_paragraphs(actual) == current_paragraphs(expected),
        "review decisions retain exactly the intended agreement text",
    )
    revisions = lambda root: [n for n in root.iter() if n.tag in {W + "ins", W + "del"}]
    card.check(
        all(n.get(W + "author") == "Bob Reviewer" for n in revisions(actual)),
        "requested decisions are resolved without new revisions",
    )
    card.check(
        [signature(n) for n in revisions(actual)]
        == [signature(n) for n in revisions(expected)],
        "Bob's pending revision text, formatting and metadata are intact",
    )
    card.check(
        formatted_chars(actual) == formatted_chars(expected),
        "remaining body characters retain their formatting",
    )
    old_comments = source.root("word/comments.xml").findall(W + "comment")
    retained = [c for c in old_comments if "cap" in text(c)]
    current = output.root("word/comments.xml").findall(W + "comment")
    card.check(
        {c.get(W + "id"): signature(c) for c in current}
        == {c.get(W + "id"): signature(c) for c in retained},
        "only the pilot discussion is removed; legal comments retain their metadata",
    )
    card.check(
        modern_comment_graph(output.parts),
        "retained comments have a complete modern identity graph",
    )
    for comment in retained:
        ident = comment.get(W + "id")
        card.check(
            comment_anchor(output, ident) == comment_anchor(source, ident),
            "legal comment keeps its exact highlight",
        )
    retained_ids = {c.get(W + "id") for c in retained}
    anchors = [
        n.get(W + "id")
        for n in actual.iter()
        if n.tag
        in {W + "commentRangeStart", W + "commentRangeEnd", W + "commentReference"}
    ]
    card.check(
        set(anchors) == retained_ids
        and all(anchors.count(i) == 3 for i in retained_ids),
        "discarded comment anchors are removed",
    )
    w14 = W14 + "paraId"
    ids = {list(c.iter(W + "p"))[-1].get(w14) for c in retained}
    entries = output.root("word/commentsExtended.xml").findall(W15 + "commentEx")
    source_entries = source.root("word/commentsExtended.xml").findall(W15 + "commentEx")
    card.check(
        {e.get(W15 + "paraId"): signature(e) for e in entries}
        == {
            e.get(W15 + "paraId"): signature(e)
            for e in source_entries
            if e.get(W15 + "paraId") in ids
        },
        "legal reply links and state survive; discarded thread records are removed",
    )
    notes = [
        n
        for n in output.root("word/footnotes.xml").findall(W + "footnote")
        if int(n.get(W + "id")) > 0
    ]
    source_notes = [
        n
        for n in source.root("word/footnotes.xml").findall(W + "footnote")
        if "quarterly invoice schedule" in text(n)
    ]
    card.check(
        [signature(n) for n in notes] == [signature(n) for n in source_notes],
        "surcharge footnote is removed and payment footnote is preserved",
    )
    refs = [n.get(W + "id") for n in actual.iter(W + "footnoteReference")]
    card.check(
        refs == [n.get(W + "id") for n in source_notes],
        "surviving payment note remains referenced",
    )
    allowed = {
        "word/document.xml",
        "word/footnotes.xml",
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/commentsExtensible.xml",
        "word/_rels/document.xml.rels",
        "[Content_Types].xml",
        "docProps/core.xml",
    }
    card.check(
        changed_parts(source, output) <= allowed,
        "unrelated package parts are preserved",
    )


def verify_tracked_terminology_update(app: Path, output: Docx, card: Card) -> None:
    from preservation import formatted_chars as chars
    from preservation import preserves_formatting

    source = Docx(app / "eval_fixtures/docx/portal-agreement.docx")
    stories = [
        name
        for name in source.parts
        if name == "word/document.xml"
        or name.startswith("word/header")
        and name.endswith(".xml")
    ]
    for name in stories:
        before, after = source.root(name), output.root(name)
        accepted = revision_projection(after, {"Casey Approver": "accept"})
        rejected = revision_projection(after, {"Casey Approver": "reject"})
        card.check(
            current_paragraphs(accepted)
            == [
                s.replace("Customer Portal", "Client Portal")
                for s in current_paragraphs(before)
            ],
            f"{name}: accepting the new changes updates all exact current names",
        )
        card.check(
            current_paragraphs(rejected) == current_paragraphs(before),
            f"{name}: rejecting the new changes restores the original current text",
        )
        card.check(
            chars(rejected) == chars(before),
            f"{name}: rejecting restores formatting and historical text",
        )
        card.check(
            preserves_formatting(before, accepted),
            f"{name}: unchanged characters retain their formatting",
        )
        prior = lambda root: {
            (n.tag, n.get(W + "id"), n.get(W + "author"), n.get(W + "date")): chars(n)
            for n in root.iter()
            if n.tag in {W + "ins", W + "del"}
            and n.get(W + "author") != "Casey Approver"
        }
        card.check(
            prior(rejected) == prior(before),
            f"{name}: earlier revisions remain pending with original metadata",
        )
        new = [
            n
            for n in after.iter()
            if n.tag in {W + "ins", W + "del"}
            and n.get(W + "author") == "Casey Approver"
        ]
        card.check(
            bool(new) and all(n.get(W + "date") == "2026-07-13T09:00:00Z" for n in new),
            f"{name}: new changes have the requested reviewer and date",
        )

        def links(root):
            return [
                (n.get(R + "id"), current_paragraphs(n) or [visible(n)])
                for n in root.iter(W + "hyperlink")
            ]

        card.check(
            links(rejected) == links(before),
            f"{name}: hyperlink labels and relationships survive rejection",
        )
        card.check(
            [signature(n.find(W + "tblPr")) for n in after.iter(W + "tbl")]
            == [signature(n.find(W + "tblPr")) for n in before.iter(W + "tbl")],
            "table properties are preserved",
        )
    card.check(
        changed_parts(source, output) <= set(stories) | {"docProps/core.xml"},
        "link destinations and unrelated package parts are unchanged",
    )


def verify_multi_edit_contract_amendment(app: Path, output: Docx, card: Card) -> None:
    source = Docx(app / "eval_fixtures/docx/services-agreement.docx")
    expected = [visible(p) for p in paragraphs(source)]
    obsolete = "The first 60 days are a pilot period and either party may end the pilot without charge."
    renewal = (
        "The agreement renews for one year unless either party gives written notice."
    )
    expected.remove(obsolete)
    anchor = expected.index("Notice: Send renewal notices to the account manager.")
    expected.insert(anchor + 1, renewal)
    section = expected.index("9. Service changes")
    expected[section + 1] = "Either party must give 45 days of written notice."
    actual = [visible(p) for p in paragraphs(output)]
    card.check(obsolete not in actual, "obsolete pilot paragraph is removed")
    card.check(
        renewal in actual
        and actual.index(renewal) > 0
        and actual[actual.index(renewal) - 1] == expected[anchor],
        "renewal paragraph follows the notice clause",
    )
    card.check(
        actual == expected,
        "all amendments are in the correct sections and other content is preserved",
    )
    inserted = paragraph(output, renewal)
    neighbor = paragraph(source, "The annual fee remains unchanged")
    card.check(
        signature(inserted.find(W + "pPr")) == signature(neighbor.find(W + "pPr")),
        "new paragraph preserves local paragraph formatting",
    )
    expected_props = {props for _, props in formatted_chars(neighbor)}
    card.check(
        {props for _, props in formatted_chars(inserted)} == expected_props,
        "new paragraph preserves local run formatting",
    )
    card.check(
        style(inserted) == style(neighbor)
        and all(n.find(W + "b") is None for n in inserted.iter(W + "rPr")),
        "renewal paragraph matches local body style without copying the bold notice label",
    )
    excluded = {
        obsolete,
        renewal,
        "Either party must give 30 days of written notice.",
        "Either party must give 45 days of written notice.",
    }
    card.check(
        unchanged_paragraphs(source, output, excluded),
        "unaffected paragraphs retain their formatting and structure",
    )
    card.check(
        changed_parts(source, output) <= {"word/document.xml", "docProps/core.xml"},
        "amendments preserve other package parts",
    )


def verify_create_review_ready_document(app: Path, output: Docx, card: Card) -> None:
    from preservation import field_records

    card.check(
        "Quarterly Review Memo" in output.all_text()
        and "Decision" in output.all_text(),
        "memo title and section are present",
    )
    card.check(
        style(paragraph(output, "Quarterly Review Memo")) == "Title"
        and style(paragraph(output, "Decision")) == "Heading1",
        "title and heading use semantic styles",
    )
    decision = paragraph(output, "Approve the renewal on revised terms.")
    card.check(
        "".join(
            visible(run)
            for run in decision.iter(W + "r")
            if (bold := run.find(W + "rPr/" + W + "b")) is not None
            and bold.get(W + "val", "1") not in {"0", "false", "off"}
        ).startswith("Approve"),
        "decision verb is bold",
    )
    table_text = [
        text(table) for table in output.root("word/document.xml").iter(W + "tbl")
    ]
    card.check(
        any("OwnerStatusLegalPending" in value for value in table_text),
        "review table is present",
    )
    tables = [
        [
            [visible(cell) for cell in row.findall(W + "tc")]
            for row in table.findall(W + "tr")
        ]
        for table in output.root("word/document.xml").iter(W + "tbl")
    ]
    card.check(
        [["Owner", "Status"], ["Legal", "Pending"]] in tables,
        "review table has two rows and two columns",
    )
    link = next(
        (
            node
            for node in output.root("word/document.xml").iter(W + "hyperlink")
            if "Policy portal" in text(node)
        ),
        None,
    )
    relation = None if link is None else rels(output).get(link.get(R + "id", ""))
    card.check(
        relation is not None
        and relation.get("Target") == "https://example.org/policies",
        "policy portal is a live external hyperlink",
    )
    card.check(
        "Renewal subject to final approval." in text(output.root("word/footnotes.xml")),
        "renewal footnote is present",
    )
    card.check(
        note_at(output, "renewal", "footnote", "Renewal subject to final approval."),
        "renewal footnote resolves at its anchor",
    )
    comments = output.root("word/comments.xml")
    review = [
        node
        for node in comments.findall("./" + W + "comment")
        if "Confirm pricing before signature." in text(node)
    ]
    card.check(
        len(review) == 1
        and review[0].get(W + "author") == "Review Lead"
        and review[0].get(W + "initials") == "RL",
        "review comment has the requested identity",
    )
    card.check(
        len(review) == 1
        and review[0].get(W + "date") == "2026-07-13T09:00:00Z"
        and comment_anchor(output, review[0].get(W + "id")) == "revised terms",
        "review comment has the requested date and exact anchor",
    )
    footer_fields = [
        tokens
        for root in active_stories(output, "footer")
        for tokens, _ in field_records(root)
    ]
    card.check(
        any(tokens[0] == "PAGE" for tokens in footer_fields),
        "footer contains a live page-number field",
    )
    protection = output.root("word/settings.xml").find(".//" + W + "documentProtection")
    card.check(
        protection is not None
        and protection.get(W + "edit") == "comments"
        and protection.get(W + "enforcement") == "1",
        "document is protected for comments-only review",
    )


VERIFY = {
    "numbering-restart": verify_numbering_restart,
    "rich-block-insertion": verify_rich_block_insertion,
    "normalized-context-targeting": verify_normalized_context_targeting,
    "data-bound-control-sync": verify_data_bound_control_sync,
    "isolated-picture-replacement": verify_isolated_picture_replacement,
    "hyperlink-lifecycle": verify_hyperlink_lifecycle,
    "footnote-endnote-authoring": verify_footnote_endnote_authoring,
    "caption-cross-reference": verify_caption_cross_reference,
    "comment-thread-deletion": verify_comment_thread_deletion,
    "restrict-editing-authoring": verify_restrict_editing_authoring,
    "append-source-letterhead": verify_append_source_letterhead,
    "create-review-ready-document": verify_create_review_ready_document,
    "multi-edit-contract-amendment": verify_multi_edit_contract_amendment,
    "review-decisions-with-dependent-content": verify_review_decisions_with_dependent_content,
    "tracked-terminology-update": verify_tracked_terminology_update,
}


def main() -> int:
    if len(sys.argv) != 2 or sys.argv[1] not in TASKS:
        print(f"usage: grader.py {'|'.join(TASKS)}", file=sys.stderr)
        return 2
    task = sys.argv[1]
    app = Path(os.environ.get("HARBOR_WORKDIR", "/app"))
    if not app.exists():
        app = Path.cwd()
    card = Card(task)
    spec = TASKS[task]
    try:
        for filename, expected in spec["inputs"].items():
            source = app / "eval_fixtures/docx" / filename
            if not source.is_file() or sha256(source) != expected:
                raise ValueError(f"immutable input is missing or changed: {filename}")
        output_path = app / spec["output"]
        output = Docx(output_path)
        VERIFY[task](app, output, card)
        for filename, expected in spec["inputs"].items():
            if sha256(app / "eval_fixtures/docx" / filename) != expected:
                raise ValueError(f"immutable input changed during grading: {filename}")
    except Exception as exc:  # noqa: BLE001 - artifact failures must emit a zero-score envelope.
        card.hard_failures.append(f"{type(exc).__name__}: {exc}")
    grading = card.payload()
    envelope = {
        "schema": "docx-grader-score-channel",
        "version": 1,
        "score": grading["score"],
        "grading": grading,
    }
    print(json.dumps(envelope, sort_keys=True, separators=(",", ":")))
    if card.score == 1.0 and not card.hard_failures:
        print(f"PASS {task} score=1.000000", file=sys.stderr)
        return 0
    print(f"FAIL {task} score={card.score:.6f}", file=sys.stderr)
    for item in card.checks:
        if not item["passed"]:
            print(f"- {item['name']}", file=sys.stderr)
    for failure in card.hard_failures:
        print(f"- {failure}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
