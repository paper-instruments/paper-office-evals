"""Independent, bounded outcome grading for generated review/edit scenarios."""
from __future__ import annotations

from collections import Counter
import hashlib
import json
import math
import os
from pathlib import Path
import posixpath
import sys
from xml.etree import ElementTree as ET

from pptx_runtime import _publish_result, run_powerpoint_open_gate, validate_pptx_package

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def read_json(path):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result: raise ValueError("duplicate JSON key: " + key)
            result[key] = value
        return result
    def reject(value):
        raise ValueError("non-finite JSON number: " + value)
    return json.loads(path.read_text(), object_pairs_hook=unique, parse_constant=reject)


def canonical(value):
    if isinstance(value, float):
        if not math.isfinite(value): raise ValueError("non-finite number")
        rounded = round(value, 3)
        return int(rounded) if rounded.is_integer() else rounded
    if isinstance(value, dict):
        result = {k: canonical(v) for k, v in value.items()}
        for key in ("color_rgb", "fill_rgb", "line_rgb"):
            color = result.get(key)
            if isinstance(color, str):
                color = color.lstrip("#")
                if len(color) == 6 and all(c in "0123456789abcdefABCDEF" for c in color): result[key] = color.upper()
        return result
    if isinstance(value, list): return [canonical(v) for v in value]
    return value


def token(value):
    return json.dumps(canonical(value), sort_keys=True, ensure_ascii=False, separators=(",", ":"))


def fact_score(actual, expected):
    got, wanted = Counter(map(token, actual)), Counter(map(token, expected))
    true = sum((got & wanted).values())
    precision = true / len(actual) if actual else float(not expected)
    recall = true / len(expected) if expected else float(not actual)
    f1 = 2 * precision * recall / (precision + recall) if precision + recall else 0.0
    return f1, got == wanted, precision, recall


def report_checks(payload, expected):
    checks = []
    if expected["kind"] == "audit":
        if not isinstance(payload, dict) or set(payload) != {"paragraphs", "shapes"} or not isinstance(payload["paragraphs"], list) or not isinstance(payload["shapes"], list):
            raise ValueError("report must contain paragraphs and shapes arrays")
        rows = payload["paragraphs"]
        required = {"slide", "object", "paragraph", "font", "size_pt", "color_rgb", "bullet"}
        if any(not isinstance(row, dict) or set(row) != required for row in rows):
            raise ValueError("invalid paragraph report entry")
        score, passed, precision, recall = fact_score(rows, expected["rows"])
        checks = [{"name": "effective typography and text-color inventory", "score": score, "passed": passed,
                   "precision": precision, "recall": recall}]
        shape_rows = payload["shapes"]
        if any(not isinstance(row, dict) or set(row) != {"slide", "object", "fill_rgb", "line_rgb"} for row in shape_rows):
            raise ValueError("invalid shape-color entry")
        score, passed, precision, recall = fact_score(shape_rows, expected["shapes"])
        checks.append({"name": "effective shape-color inventory", "score": score, "passed": passed,
                       "precision": precision, "recall": recall})
        return checks
    if not isinstance(payload, dict) or set(payload) != {"reviews"} or not isinstance(payload["reviews"], list):
        raise ValueError("report must contain a reviews array")
    reviews = {}
    for review in payload["reviews"]:
        if not isinstance(review, dict) or set(review) != {"review", "changes"} or not isinstance(review["changes"], list):
            raise ValueError("invalid review entry")
        key = review["review"]
        if not isinstance(key, str) or key in reviews:
            raise ValueError("duplicate or invalid review identifier")
        for row in review["changes"]:
            if not isinstance(row, dict) or set(row) != {"kind", "slide", "shape_id", "before", "after"}:
                raise ValueError("invalid change entry")
        reviews[key] = review["changes"]
    complete = set(reviews) == set(expected["reviews"])
    checks.append({"name": "all supplied reviews, once each", "score": float(complete), "passed": complete})
    for key in expected["reviews"]:
        wanted = [{k: v for k, v in row.items() if k != "review"} for row in expected["rows"] if row["review"] == key]
        score, passed, precision, recall = fact_score(reviews.get(key, []), wanted)
        passed = passed and key in reviews
        checks.append({"name": "review " + key, "score": score if key in reviews else 0.0,
                       "passed": passed, "precision": precision, "recall": recall})
    return checks


def selected_paragraphs(root, shape_id):
    if shape_id is None:
        return root.findall(".//a:p", NS)
    for shape in root.iter():
        if shape.tag not in {"{%s}sp" % NS["p"], "{%s}graphicFrame" % NS["p"]}:
            continue
        identity = shape.find(".//p:cNvPr", NS)
        if identity is not None and identity.get("id") == str(shape_id):
            return shape.findall(".//a:p", NS)
    return []


def paragraph_text(paragraph):
    return "".join(node.text or "" for node in paragraph.findall(".//a:t", NS))


def expected_edit(root, edit):
    paragraphs = selected_paragraphs(root, edit["shape_id"])
    matches = [p for p in paragraphs if edit["find"] in paragraph_text(p)]
    if len(matches) != 1:
        raise ValueError("verifier baseline edit is ambiguous")
    p = matches[0]
    nodes = p.findall(".//a:t", NS)
    text = paragraph_text(p)
    start = text.index(edit["find"])
    end = start + len(edit["find"])
    offset = 0
    first = True
    for node in nodes:
        value = node.text or ""
        left, right = max(start - offset, 0), min(end - offset, len(value))
        if left < right:
            node.text = value[:left] + (edit["replace"] if first else "") + value[right:]
            first = False
        offset += len(value)


def signature(node):
    children = []
    for child in node:
        current = signature(child)
        # Splitting a run with identical formatting does not change the content.
        if child.tag == "{%s}r" % NS["a"]:
            props = child.find("a:rPr", NS)
            run = ("text-run", signature(props) if props is not None else None,
                   child.findtext("a:t", default="", namespaces=NS))
            if children and children[-1][0] == "text-run" and children[-1][1] == run[1]:
                children[-1] = (run[0], run[1], children[-1][2] + run[2])
            else:
                children.append(run)
        else:
            children.append(current)
    text = node.text or ""
    if node.tag != "{%s}t" % NS["a"] and not text.strip():
        text = ""
    return node.tag, tuple(sorted(node.attrib.items())), text, tuple(children)


def member_signature(name, data):
    if not name.endswith((".xml", ".rels")) and name != "[Content_Types].xml":
        return data
    root = ET.fromstring(data)
    if name.endswith(".rels"):
        owner = name.replace("/_rels/", "/")[:-5] if name != "_rels/.rels" else ""
        rows = []
        for node in root:
            attrs = dict(node.attrib)
            if attrs.get("TargetMode", "Internal") == "Internal":
                attrs.pop("TargetMode", None)
                target = attrs["Target"]
                attrs["Target"] = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(owner), target))
            rows.append(tuple(sorted(attrs.items())))
        return tuple(sorted(rows))
    if name == "[Content_Types].xml":
        return tuple(sorted(signature(child) for child in root))
    return signature(root)


def edit_checks(output, baseline, expected):
    trees = {name: ET.fromstring(output[name]) for name in {e["member"] for e in expected["edits"]}}
    checks = []
    for i, edit in enumerate(expected["edits"], 1):
        actual = ET.fromstring(output[edit["member"]])
        paragraphs = selected_paragraphs(actual, edit["shape_id"])
        present = sum(edit["replace"] in paragraph_text(p) for p in paragraphs) == 1
        old_absent = all(edit["find"] not in paragraph_text(p) for p in paragraphs)
        passed = present and old_absent
        checks.append({"name": "approved correction %d" % i, "passed": passed, "score": float(passed)})
        if present:
            reverse = dict(edit, find=edit["replace"], replace=edit["find"])
            expected_edit(trees[edit["member"]], reverse)
    # Undo only the approved wording changes. Missing edits lose effect credit;
    # unauthorized edits or lost formatting independently lose preservation.
    normalized = dict(output)
    normalized.update({name: ET.tostring(tree) for name, tree in trees.items()})
    members_same = set(output) == set(baseline)
    mismatches = [name for name in baseline if name not in normalized or member_signature(name, normalized[name]) != member_signature(name, baseline[name])]
    preserved = members_same and not mismatches
    checks.append({"name": "formatting and all other package content preserved", "passed": preserved,
                   "score": float(preserved), "different_members": mismatches})
    return checks


def grade(root):
    expected = read_json(Path(__file__).with_name("expected.json"))
    files = root / "eval_fixtures/pptx"
    packages = {}
    for name, digest in expected["inputs"].items():
        path = files / name
        if hashlib.sha256(path.read_bytes()).hexdigest() != digest:
            raise ValueError("input was modified: " + name)
        if path.suffix == ".pptx": packages[name] = validate_pptx_package(path)
    output = root / "evals" / expected["task"] / expected["output"]
    gates = []
    if expected["kind"] == "edit":
        members = validate_pptx_package(output)
        open_gate = os.environ.get("PAPER_PPTX_OPEN_GATE_CMD")
        if os.environ.get("PAPER_PPTX_REQUIRE_POWERPOINT") == "1" and not open_gate:
            raise ValueError("PowerPoint gate is required but PAPER_PPTX_OPEN_GATE_CMD is unset")
        if open_gate:
            gates.append(run_powerpoint_open_gate(open_gate, output))
        checks = edit_checks(members, packages[expected["baseline"]], expected)
    else:
        checks = report_checks(read_json(output), expected)
    success = bool(checks) and all(check["passed"] for check in checks)
    reward = float(success)
    return {"schema": "paper-pptx-harbor-grading", "version": 3, "task": expected["task"],
            "reward": round(reward, 4), "task_success": success, "checks": checks, "hard_failures": [],
            "hard_gate_details": {"powerpoint_open_gates": gates}}


def main():
    try:
        result = grade(Path(sys.argv[1]).resolve())
    except Exception as exc:
        result = {"reward": 0.0, "task_success": False, "checks": [],
                  "hard_failures": [type(exc).__name__ + ": " + str(exc)]}
    _publish_result(result)
    print(json.dumps(result))


if __name__ == "__main__":
    main()
