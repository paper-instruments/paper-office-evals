"""Package-neutral checks shared by the standalone DOCX verifiers."""

import io
import posixpath
import shlex
from copy import deepcopy
from difflib import SequenceMatcher
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def signature(node):
    if node is None:
        return None
    if (
        node.tag in {W + "pPr", W + "rPr", W + "trPr"}
        and not len(node)
        and not node.attrib
    ):
        return None
    return (
        node.tag,
        tuple(sorted(node.attrib.items())),
        (node.text or "") if not len(node) or (node.text or "").strip() else "",
        tuple(signature(child) for child in node),
        node.tail if (node.tail or "").strip() else "",
    )


def live_text(root):
    if root.tag in {W + "del", W + "moveFrom"}:
        return ""
    if root.tag == W + "t":
        return root.text or ""
    return "".join(live_text(child) for child in root)


def formatted_chars(root):
    return [
        (char, signature(run.find(W + "rPr")))
        for run in root.iter(W + "r")
        for node in run
        if node.tag in {W + "t", W + "delText"}
        for char in node.text or ""
    ]


def preserves_formatting(before, after):
    old, new = formatted_chars(before), formatted_chars(after)
    matcher = SequenceMatcher(
        None, "".join(c for c, _ in old), "".join(c for c, _ in new), autojunk=False
    )
    return all(
        old[a:b] == new[c:d]
        for op, a, b, c, d in matcher.get_opcodes()
        if op == "equal"
    )


def semantic_xml(payload):
    return run_normalized_signature(ET.fromstring(payload))


def equivalent_part(before, after):
    if before == after:
        return True
    if before is None or after is None:
        return False
    try:
        return semantic_xml(before) == semantic_xml(after)
    except ET.ParseError:
        return False


def changed_parts(before, after):
    return {
        name
        for name in set(before) | set(after)
        if not equivalent_part(before.get(name), after.get(name))
    }


def normalized_tree(node):
    """Coalesce identical adjacent runs/text without crossing semantic boundaries."""
    root = deepcopy(node)
    if root is None:
        return None
    text_tags = {W + "t", W + "delText", W + "instrText"}
    space = "{http://www.w3.org/XML/1998/namespace}space"

    def normalize_space(element, inherited="default"):
        mode = element.get(space, inherited)
        if element.tag in text_tags:
            if mode != "preserve":
                element.text = (element.text or "").strip(" \t\r\n")
            # Keep normalized text stable if a caller normalizes it again.
            element.set(space, "preserve")
        for child in element:
            normalize_space(child, mode)

    # Resolve whitespace at each original text boundary before merging runs.
    normalize_space(root)
    for parent in reversed(list(root.iter())):
        for child in list(parent):
            if (
                child.tag == W + "proofErr"
                or (
                    child.tag in {W + "pPr", W + "rPr", W + "trPr"}
                    and parent.tag in {W + "p", W + "r", W + "tr", W + "pPr"}
                    and not len(child)
                    and not child.attrib
                )
                or (child.tag == W + "r" and all(n.tag == W + "rPr" for n in child))
            ):
                parent.remove(child)
    for parent in root.iter():
        previous = None
        for child in list(parent):
            if child.tag != W + "r":
                previous = None
                continue
            if (
                previous is not None
                and previous.attrib == child.attrib
                and signature(previous.find(W + "rPr"))
                == signature(child.find(W + "rPr"))
            ):
                for item in list(child):
                    if item.tag != W + "rPr":
                        previous.append(item)
                parent.remove(child)
            else:
                previous = child
    for parent in root.iter():
        previous = None
        for child in list(parent):
            if (
                child.tag in text_tags
                and previous is not None
                and previous.tag == child.tag
                and previous.attrib == child.attrib
            ):
                previous.text = (previous.text or "") + (child.text or "")
                parent.remove(child)
                continue
            previous = child
    return root


def run_normalized_signature(node):
    return signature(normalized_tree(node))


def without(root, tags):
    root = deepcopy(root)
    for parent in root.iter():
        for child in list(parent):
            if child.tag in {W + tag for tag in tags}:
                parent.remove(child)
    return root


def paragraph_properties(root):
    return [signature(p.find(W + "pPr")) for p in root.iter(W + "p")]


def reference_positions(root, tag):
    """Record note/bookmark positions in paragraph text, including empty paragraphs."""
    result = []
    for index, paragraph in enumerate(root.iter(W + "p")):
        offset = 0
        for node in paragraph.iter():
            if node.tag == W + tag:
                result.append((index, offset, tuple(sorted(node.attrib.items()))))
            elif node.tag in {W + "t", W + "delText"}:
                offset += len(node.text or "")
    return result


def field_records(root):
    """Parse complete simple/complex fields; reject orphan instructions and bad nesting."""
    records, stack = [], []

    def walk(node):
        if node.tag == W + "fldSimple":
            records.append(
                (
                    tuple(shlex.split(node.get(W + "instr", ""), posix=False)),
                    live_text(node),
                )
            )
            return
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                stack.append(["", "", False])
            elif kind == "separate" and stack and not stack[-1][2]:
                stack[-1][2] = True
            elif kind == "end" and stack:
                instruction, result, _ = stack.pop()
                if not instruction.strip():
                    raise ValueError("field has no instruction")
                records.append((tuple(shlex.split(instruction, posix=False)), result))
            else:
                raise ValueError("unbalanced field boundary")
        elif node.tag == W + "instrText":
            if not stack or stack[-1][2]:
                raise ValueError("field instruction outside instruction region")
            stack[-1][0] += node.text or ""
        elif node.tag == W + "t" and stack and stack[-1][2]:
            stack[-1][1] += node.text or ""
        for child in node:
            walk(child)

    walk(root)
    if stack:
        raise ValueError("unclosed field")
    if any(not instruction for instruction, _ in records):
        raise ValueError("empty field instruction")
    return records


def binding_address(binding):
    # Collect all namespace declarations, including multiple aliases.
    namespaces = dict(
        value
        for _, value in ET.iterparse(
            io.StringIO("<root " + binding.get(W + "prefixMappings", "") + "/>"),
            events=("start-ns",),
        )
    )
    path = []
    for component in binding.get(W + "xpath", "").split("/"):
        if not component:
            continue
        attr = component.startswith("@")
        name = component.lstrip("@").removesuffix("[1]")
        if ":" in name:
            prefix, local = name.split(":", 1)
            name = "{" + namespaces[prefix] + "}" + local
        path.append((attr, name))
    return binding.get(W + "storeItemID", "").upper(), tuple(path)


def bound_store(parts, ident):
    ds = "{http://schemas.openxmlformats.org/officeDocument/2006/customXml}"
    matches = []
    relationships = ET.fromstring(parts["word/_rels/document.xml.rels"])
    reachable = {
        posixpath.normpath(posixpath.join("word", n.get("Target", ""))).lstrip("/")
        for n in relationships
        if n.get("Type", "").endswith("/customXml")
    }
    for name in reachable:
        rel_path = posixpath.join(
            posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels"
        )
        for rel in ET.fromstring(parts[rel_path]):
            if not rel.get("Type", "").endswith("/customXmlProps"):
                continue
            props = posixpath.normpath(
                posixpath.join(posixpath.dirname(name), rel.get("Target"))
            ).lstrip("/")
            root = ET.fromstring(parts[props])
            if root.get(ds + "itemID", "").upper() == ident.upper():
                matches.append((name, ET.fromstring(parts[name]), props))
    if len(matches) != 1:
        raise ValueError("binding must resolve to exactly one connected store")
    return matches[0]


def remapped_parts(parts, names):
    """Compare a known part rename without exempting relationship or type edits."""
    names = dict(names)
    for old, new in list(names.items()):
        old_rels = posixpath.join(
            posixpath.dirname(old), "_rels", posixpath.basename(old) + ".rels"
        )
        if old_rels in parts:
            names[old_rels] = posixpath.join(
                posixpath.dirname(new), "_rels", posixpath.basename(new) + ".rels"
            )
    result = {}
    for old, payload in parts.items():
        new = names.get(old, old)
        if new in result:
            raise ValueError("part rename collides with an unrelated part")
        if old.endswith(".rels"):
            root = ET.fromstring(payload)
            base = posixpath.dirname(posixpath.dirname(old))
            for rel in root:
                if rel.get("TargetMode") == "External":
                    continue
                target = posixpath.normpath(
                    posixpath.join(base, rel.get("Target", ""))
                ).lstrip("/")
                rel.set("Target", "/" + names.get(target, target))
            payload = ET.tostring(root)
        elif old == "[Content_Types].xml":
            root = ET.fromstring(payload)
            for entry in root:
                target = entry.get("PartName", "").lstrip("/")
                if target in names:
                    entry.set("PartName", "/" + names[target])
            payload = ET.tostring(root)
        result[new] = payload
    return result


def project_revisions(root, decisions):
    """Resolve selected insertion/deletion IDs while preserving other XML content."""
    root = deepcopy(root)

    def visit(parent):
        for child in list(parent):
            if child.tag in {W + "p", W + "tr"}:
                path = W + "pPr/" + W + "rPr/" if child.tag == W + "p" else W + "trPr/"
                marks = [(kind, child.find(path + W + kind)) for kind in ("ins", "del")]
                if any(
                    mark is not None and decisions.get(mark.get(W + "id")) == decision
                    for kind, mark in marks
                    for decision in ["reject" if kind == "ins" else "accept"]
                ):
                    parent.remove(child)
                    continue
            decision = decisions.get(child.get(W + "id"))
            if (
                child.tag
                in {
                    W + "moveFromRangeStart",
                    W + "moveFromRangeEnd",
                    W + "moveToRangeStart",
                    W + "moveToRangeEnd",
                }
                and decision
            ):
                parent.remove(child)
            elif (
                child.tag in {W + "ins", W + "del", W + "moveFrom", W + "moveTo"}
                and decision
            ):
                offset = list(parent).index(child)
                parent.remove(child)
                if (child.tag in {W + "ins", W + "moveTo"}) == (decision == "accept"):
                    visit(child)
                    for item in list(child):
                        if child.tag == W + "del":
                            for text in item.iter(W + "delText"):
                                text.tag = W + "t"
                        parent.insert(offset, item)
                        offset += 1
            else:
                visit(child)

    visit(root)
    # Resolving paragraph marks can leave empty property containers.
    for parent in reversed(list(root.iter())):
        for child in list(parent):
            if (
                child.tag in {W + "rPr", W + "pPr", W + "trPr"}
                and parent.tag in {W + "r", W + "p", W + "tr", W + "pPr"}
                and not len(child)
                and not child.attrib
                or child.tag == W + "tbl"
                and not child.findall(W + "tr")
            ):
                parent.remove(child)
    return root


def comment_anchor(root, ident):
    active, chunks, counts = False, [], [0, 0, 0]
    for node in root.iter():
        if node.tag == W + "commentRangeStart" and node.get(W + "id") == ident:
            counts[0] += 1
            active = True
        elif node.tag == W + "commentRangeEnd" and node.get(W + "id") == ident:
            counts[1] += 1
            active = False
        elif node.tag == W + "commentReference" and node.get(W + "id") == ident:
            counts[2] += 1
        elif active and node.tag == W + "t":
            chunks.append(node.text or "")
    return "".join(chunks) if counts == [1, 1, 1] else None


def modern_comment_graph(parts):
    w14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
    cid = "{http://schemas.microsoft.com/office/word/2016/wordml/cid}"
    cex = "{http://schemas.microsoft.com/office/word/2018/wordml/cex}"
    try:
        comments = ET.fromstring(parts["word/comments.xml"]).findall(W + "comment")
        paragraphs = [list(c.iter(W + "p"))[-1].get(w14 + "paraId") for c in comments]
        ids = ET.fromstring(parts["word/commentsIds.xml"]).findall(cid + "commentId")
        durable = [n.get(cid + "durableId") for n in ids]
        extended = ET.fromstring(parts["word/commentsExtensible.xml"]).findall(
            cex + "commentExtensible"
        )
        return (
            None not in paragraphs
            and len(paragraphs) == len(set(paragraphs))
            and sorted(n.get(cid + "paraId", "") for n in ids) == sorted(paragraphs)
            and None not in durable
            and len(durable) == len(set(durable))
            and sorted(durable)
            == sorted(n.get(cex + "durableId", "") for n in extended)
        )
    except (KeyError, IndexError, ET.ParseError):
        return False
