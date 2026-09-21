"""Grade editable, template-independent handoff without importing either pptx library."""

from __future__ import annotations

import copy
import hashlib
import json
import os
from pathlib import Path
import sys
from xml.etree import ElementTree as ET

import pptx_runtime as rt
from pptx_runtime import _publish_result, run_powerpoint_open_gate, validate_pptx_package

TASK = "appearance-preserving-slide-handoff"
NS, A, P, R = rt.NS, rt.A, rt.P, rt.R


def xml_signature(node):
    if node is None:
        return None
    node = copy.deepcopy(node)
    for paragraph in node.iter(rt.qn(A, "p")):
        previous = None
        for child in list(paragraph):
            if child.tag != rt.qn(A, "r"):
                previous = None
                continue
            props, text = child.find("./a:rPr", NS), child.find("./a:t", NS)
            style = rt.node_signature(props) if props is not None and (props.attrib or len(props)) else None
            if text is not None and previous is not None and previous[0] == style:
                previous[1].text = (previous[1].text or "") + (text.text or "")
                paragraph.remove(child)
            elif text is not None:
                previous = style, text
    return rt.node_signature(node)


def related(members, owner, kind):
    return [target for reltype, target, mode in rt.relationships(members, owner).values() if reltype == R + "/" + kind and mode != "External"]


def shape_chain(members, slide, shape):
    layouts = related(members, slide["member"], "slideLayout")
    layout = layouts[0] if len(layouts) == 1 else None
    masters = related(members, layout, "slideMaster") if layout else []
    master = masters[0] if len(masters) == 1 else None
    themes = related(members, master, "theme") if master else []
    theme = rt.xml(members, themes[0]) if len(themes) == 1 else None
    placeholder = rt.shape_placeholder(shape)
    owners = []
    if placeholder is not None and layout:
        candidates = [s for s in rt.iter_shapes(rt.xml(members, layout)) if rt.shape_placeholder(s) is not None and rt.shape_placeholder(s).get("idx", "0") == placeholder.get("idx", "0")]
        layout_shape = candidates[0] if len(candidates) == 1 else None
        kind = placeholder.get("type", "obj")
        if layout_shape is not None:
            kind = rt.shape_placeholder(layout_shape).get("type", "obj")
        family = {"title", "ctrTitle"} if kind in {"title", "ctrTitle"} else {"body", "obj"} if kind in {"body", "obj"} else {kind}
        if master:
            candidates = [s for s in rt.iter_shapes(rt.xml(members, master)) if rt.shape_placeholder(s) is not None and rt.shape_placeholder(s).get("type", "obj") in family]
            if len(candidates) == 1:
                owners.append(candidates[0])
        if layout_shape is not None:
            owners.append(layout_shape)
    return owners + [shape], master, theme


def geometry(members, slide, shape):
    owners, _, _ = shape_chain(members, slide, shape)
    result = {"rot": 0, "flipH": False, "flipV": False}
    for owner in owners:
        transform = rt.xfrm(owner)
        if transform is None:
            continue
        for name in ("rot", "flipH", "flipV"):
            if name in transform.attrib:
                value = transform.get(name)
                result[name] = int(value) if name == "rot" else value in {"true", "1"}
        for child in transform:
            if rt.local(child.tag) in {"off", "ext"}:
                result.update({key: int(value) for key, value in child.attrib.items()})
    return tuple(result.get(key) for key in ("x", "y", "cx", "cy", "rot", "flipH", "flipV"))


def resolved_font(token, theme):
    if token not in {"+mj-lt", "+mn-lt"}:
        return token
    node = theme.find(".//a:%sFont/a:latin" % ("major" if token == "+mj-lt" else "minor"), NS) if theme is not None else None
    return node.get("typeface") if node is not None else None


def resolved_color(props, theme, master):
    fill = props.find("./a:solidFill", NS)
    if fill is None:
        return "non-solid text fill" if any(props.find("./a:" + name, NS) is not None for name in ("noFill", "gradFill", "blipFill", "pattFill", "grpFill")) else None
    rgb = fill.find("./a:srgbClr", NS)
    if rgb is not None:
        return rgb.get("val", "").upper() if not len(rgb) else "unresolved transformed color"
    scheme = fill.find("./a:schemeClr", NS)
    if scheme is not None and theme is not None and not len(scheme):
        token = scheme.get("val")
        mapping = master.find("./p:clrMap", NS) if master is not None else None
        token = mapping.get(token, token) if mapping is not None else token
        node = theme.find(".//a:clrScheme/a:" + token, NS)
        if node is not None and len(node) == 1:
            return node[0].get("val") if rt.local(node[0].tag) == "srgbClr" else node[0].get("lastClr")
    return "unresolved color"


def paragraph_appearance(members, slide, shape, paragraph):
    owners, master_member, theme = shape_chain(members, slide, shape)
    master = rt.xml(members, master_member) if master_member else None
    ppr = paragraph.find("./a:pPr", NS)
    level = int(ppr.get("lvl", "0")) + 1 if ppr is not None else 1
    ph = rt.shape_placeholder(shape)
    category = "titleStyle" if ph is not None and ph.get("type") in {"title", "ctrTitle"} else "bodyStyle" if ph is not None else "otherStyle"
    sources = [rt.xml(members, "ppt/presentation.xml").find("./p:defaultTextStyle/a:lvl%dpPr" % level, NS)]
    if master is not None:
        sources.append(master.find("./p:txStyles/p:%s/a:lvl%dpPr" % (category, level), NS))
    for owner in owners:
        sources.append(owner.find("./p:txBody/a:lstStyle/a:lvl%dpPr" % level, NS))
        if owner is not shape:
            sources.append(owner.find("./p:txBody/a:p/a:pPr", NS))
    sources.append(ppr)
    layout = {"algn": "l", "marL": "0", "indent": "0", "bullet": "none"}
    for source in sources:
        if source is None:
            continue
        for key in ("algn", "marL", "indent"):
            if key in source.attrib:
                layout[key] = source.get(key)
        for tag in ("lnSpc", "spcBef", "spcAft"):
            node = source.find("./a:" + tag, NS)
            if node is not None:
                layout[tag] = xml_signature(node)
        for child in source:
            if rt.local(child.tag) in {"buNone", "buChar", "buAutoNum", "buBlip"}:
                layout["bullet"] = "none" if rt.local(child.tag) == "buNone" else xml_signature(child)
    font_sources = [source.find("./a:defRPr", NS) if source is not None else None for source in sources]
    chars = []
    for child in paragraph:
        if child.tag == rt.qn(A, "br"):
            chars.append(("\n", None))
            continue
        if child.tag != rt.qn(A, "r"):
            continue
        font = {"name": None, "size": None, "color": "000000", "bold": False, "italic": False, "underline": "none", "strike": "noStrike", "cap": "none", "baseline": "0", "spc": "0"}
        for props in font_sources + [child.find("./a:rPr", NS)]:
            if props is None:
                continue
            for name in ("strike", "cap", "baseline", "spc"):
                if name in props.attrib:
                    font[name] = props.get(name)
            for attr, key in (("sz", "size"), ("b", "bold"), ("i", "italic"), ("u", "underline")):
                if attr in props.attrib:
                    value = props.get(attr)
                    font[key] = int(value) if attr == "sz" else value in {"1", "true"} if attr in {"b", "i"} else value
            latin = props.find("./a:latin", NS)
            if latin is not None:
                font["name"] = resolved_font(latin.get("typeface"), theme)
            color = resolved_color(props, theme, master)
            if color is not None:
                font["color"] = color
        text = child.findtext("./a:t", default="", namespaces=NS)
        chars.extend((character, tuple(sorted(font.items()))) for character in text)
    return tuple(sorted(layout.items())), tuple(chars)


def text_shapes(members, slide):
    result = {}
    for shape in rt.iter_shapes(slide["root"]):
        paragraphs = shape.findall("./p:txBody/a:p", NS)
        key = tuple(rt.shape_text(p) for p in paragraphs)
        if paragraphs and any(key):
            if key in result:
                raise ValueError("duplicate editable source content")
            result[key] = (geometry(members, slide, shape), text_frame(members, slide, shape), surface(members, slide, shape), tuple(paragraph_appearance(members, slide, shape, p) for p in paragraphs))
    return result


def text_frame(members, slide, shape):
    attrs = {"rot": "0", "vert": "horz", "wrap": "square", "lIns": "91440", "tIns": "45720", "rIns": "91440", "bIns": "45720", "rtlCol": "0", "anchor": "t", "anchorCtr": "0", "numCol": "1", "spcCol": "0", "upright": "0", "compatLnSpc": "0", "forceAA": "0", "fromWordArt": "0"}
    children = {}
    for owner in shape_chain(members, slide, shape)[0]:
        props = owner.find("./p:txBody/a:bodyPr", NS)
        if props is None:
            continue
        attrs.update(props.attrib)
        for child in props:
            tag = rt.local(child.tag)
            if tag in {"normAutofit", "noAutofit", "spAutoFit"}:
                children.pop("autofit", None)
                if tag == "normAutofit":
                    children["autofit"] = (tag, child.get("fontScale", "100000"), child.get("lnSpcReduction", "0"))
                elif tag != "noAutofit":
                    children["autofit"] = (tag,)
            else:
                children[tag] = xml_signature(child)
    for key in ("rtlCol", "anchorCtr", "upright", "compatLnSpc", "forceAA", "fromWordArt"):
        attrs[key] = attrs[key] in {"true", "1"}
    return tuple(sorted(attrs.items())), tuple(sorted(children.items()))


def surface(members, slide, shape):
    properties = {}
    for owner in shape_chain(members, slide, shape)[0]:
        props = owner.find("./p:spPr", NS)
        if props is not None:
            for child in props:
                tag = rt.local(child.tag)
                if tag in {"xfrm", "prstGeom"}:
                    if tag == "prstGeom" and (child.get("prst") != "rect" or any(len(c) for c in child)):
                        properties["geometry"] = xml_signature(child)
                    continue
                if tag == "noFill" or (tag == "ln" and child.find("./a:noFill", NS) is not None and len(child) == 1):
                    properties.pop("fill" if tag == "noFill" else "ln", None)
                elif tag in {"solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}:
                    properties["fill"] = xml_signature(child)
                elif tag == "effectLst" and not len(child):
                    properties.pop(tag, None)
                else:
                    properties[tag] = xml_signature(child)
        style = owner.find("./p:style", NS)
        if style is not None:
            properties["style"] = xml_signature(style)
    return tuple(sorted(properties.items()))


def image_shapes(members, slide):
    result = []
    for shape in rt.iter_shapes(slide["root"]):
        if shape.tag != rt.qn(P, "pic"):
            continue
        blip = shape.find("./p:blipFill/a:blip", NS)
        target = rt.relationships(members, slide["member"])[blip.get(rt.qn(R, "embed"))][1]
        fill = copy.deepcopy(shape.find("./p:blipFill", NS))
        fill.find("./a:blip", NS).attrib.pop(rt.qn(R, "embed"), None)
        result.append((hashlib.sha256(members[target]).hexdigest(), geometry(members, slide, shape), xml_signature(fill), xml_signature(shape.find("./p:spPr", NS))))
    return result


def templates(members):
    return {name: data for name, data in members.items() if name.startswith(("ppt/slideMasters/", "ppt/slideLayouts/", "ppt/theme/", "ppt/notesMasters/"))}


def notes_signature(node):
    node = copy.deepcopy(node)
    for identity in node.findall(".//p:cNvPr", NS):
        identity.attrib.pop("id", None)
        identity.attrib.pop("name", None)
    return xml_signature(node)


def background(members, slide):
    layout = related(members, slide["member"], "slideLayout")[0]
    master = related(members, layout, "slideMaster")[0]
    master_root = rt.xml(members, master)
    theme = rt.xml(members, related(members, master, "theme")[0])
    for owner in (slide["root"], rt.xml(members, layout), master_root):
        bg = owner.find("./p:cSld/p:bg", NS)
        if bg is None:
            continue
        props = bg.find("./p:bgPr", NS)
        if props is not None:
            color = resolved_color(props, theme, master_root)
            extras = [xml_signature(c) for c in props if rt.local(c.tag) != "solidFill" and not (rt.local(c.tag) == "effectLst" and not len(c))]
            return color, tuple(extras)
        ref = bg.find("./p:bgRef", NS)
        if ref is not None:
            styles = theme.findall(".//a:bgFillStyleLst/*", NS)
            index = int(ref.get("idx", "0")) - 1001
            if 0 <= index < len(styles):
                style = copy.deepcopy(styles[index])
                for parent in style.iter():
                    for child in list(parent):
                        if child.tag == rt.qn(A, "schemeClr") and child.get("val") == "phClr" and not len(child):
                            parent.remove(child)
                            parent.append(copy.deepcopy(ref[0]))
                props = ET.Element("props")
                props.append(style)
                return resolved_color(props, theme, master_root), () if style.tag == rt.qn(A, "solidFill") else (xml_signature(style),)
        return xml_signature(bg)
    return "FFFFFF", ()


def preserved(before, after):
    mutable = {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"}
    differences = []
    for name, data in before.items():
        if name in mutable:
            continue
        if name not in after:
            differences.append(name)
        elif name.endswith(".rels"):
            owner = rt.owner_from_rels(name)
            if sorted(rt.relationships(before, owner).values()) != sorted(rt.relationships(after, owner).values()):
                differences.append(name)
        elif name.endswith(".xml"):
            if xml_signature(ET.fromstring(data)) != xml_signature(ET.fromstring(after[name])):
                differences.append(name)
        elif data != after[name]:
            differences.append(name)
    old, new = rt.xml(before, "ppt/presentation.xml"), rt.xml(after, "ppt/presentation.xml")
    if old.find("./p:notesMasterIdLst", NS) is None:
        # Registering the already-present notes master is not a template import.
        note_ids = new.find("./p:notesMasterIdLst", NS)
        if note_ids is not None and all(target in before for target in related(after, "ppt/presentation.xml", "notesMaster")):
            new.remove(note_ids)
    for presentation in (old, new):
        slide_ids = presentation.find("./p:sldIdLst", NS)
        if slide_ids is not None:
            presentation.remove(slide_ids)
    if xml_signature(old) != xml_signature(new):
        differences.append("presentation settings")
    return differences


def changed_themes(members):
    result = dict(members)
    for name in list(result):
        if not name.startswith("ppt/theme/") or not name.endswith(".xml"):
            continue
        theme = ET.fromstring(result[name])
        for token in ("majorFont", "minorFont"):
            font = theme.find(".//a:" + token + "/a:latin", NS)
            if font is not None:
                font.set("typeface", "Courier New")
        for child in theme.findall(".//a:clrScheme/*", NS):
            for color in list(child):
                child.remove(color)
            ET.SubElement(child, rt.qn(A, "srgbClr"), {"val": "CC3300"})
        result[name] = ET.tostring(theme)
    return result


def grade(root):
    manifest = json.loads(Path(__file__).with_name("fixture_hashes.json").read_text())
    for filename, digest in manifest.items():
        if hashlib.sha256((root / "eval_fixtures/pptx" / filename).read_bytes()).hexdigest() != digest:
            raise ValueError("input file changed: " + filename)
    source = validate_pptx_package(root / "eval_fixtures/pptx/partner-launch.pptx")
    before = validate_pptx_package(root / "eval_fixtures/pptx/leadership-review.pptx")
    output_path = root / "evals" / TASK / "output.pptx"
    output = validate_pptx_package(output_path)
    open_gate = os.environ.get("PAPER_PPTX_OPEN_GATE_CMD")
    if os.environ.get("PAPER_PPTX_REQUIRE_POWERPOINT") == "1" and not open_gate:
        raise ValueError("PowerPoint gate is required but PAPER_PPTX_OPEN_GATE_CMD is unset")
    if open_gate:
        run_powerpoint_open_gate(open_gate, output_path)
    baseline, records = rt.slide_records(before), rt.slide_records(output)
    effects, preservation = [], []
    def check(group, passed, name):
        group.append({"name": name, "passed": bool(passed)})
    check(effects, len(records) == len(baseline) + 1, "one partner slide appended")
    if len(records) != len(baseline) + 1:
        return result(effects, preservation)
    expected_slide, actual_slide = rt.slide_records(source)[0], records[-1]
    check(effects, actual_slide["root"].get("show", "1") in {"1", "true"}, "handed-off slide is visible")
    check(effects, all(node.get("hidden", "0") in {"0", "false"} for node in actual_slide["root"].findall(".//p:cNvPr", NS)), "handed-off objects are visible")
    expected_text = text_shapes(source, expected_slide)
    check(effects, text_shapes(output, actual_slide) == expected_text, "editable source text retains effective font, color, emphasis, paragraph spacing and geometry")
    perturbed = changed_themes(output)
    perturbed_slide = rt.slide_records(perturbed)[-1]
    check(effects, text_shapes(perturbed, perturbed_slide) == expected_text, "handed-off text survives a destination theme change")
    check(effects, image_shapes(source, expected_slide) == image_shapes(output, actual_slide), "wordmark pixels, crop and placement preserved")
    check(effects, background(source, expected_slide) == background(output, actual_slide) == background(perturbed, perturbed_slide), "slide background appearance remains independent")
    check(effects, len(list(rt.iter_shapes(actual_slide["root"]))) == len(list(rt.iter_shapes(expected_slide["root"]))), "no added or rasterized slide objects")
    check(effects, set(templates(output)) == set(templates(before)), "no source layout, master, theme or notes-master parts added")
    left_notes, right_notes = rt.notes_part(source, expected_slide), rt.notes_part(output, actual_slide)
    check(preservation, left_notes is not None and right_notes is not None and notes_signature(rt.xml(source, left_notes)) == notes_signature(rt.xml(output, right_notes)), "speaker notes and their formatting preserved")
    check(preservation, all(left["id"] == right["id"] and xml_signature(left["root"]) == xml_signature(right["root"]) for left, right in zip(baseline, records)), "existing slides and ordering preserved")
    differences = preserved(before, output)
    check(preservation, not differences, "existing template, chart, media, notes and metadata preserved: " + ", ".join(differences))
    return result(effects, preservation)


def result(effects, preservation):
    effect_ok = bool(effects) and all(item["passed"] for item in effects)
    preservation_ok = bool(preservation) and all(item["passed"] for item in preservation)
    return {"schema": "paper-pptx-harbor-grading", "version": 3, "task": TASK, "reward": float(effect_ok and preservation_ok), "task_success": effect_ok and preservation_ok, "checks": effects + preservation, "hard_failures": []}


def main():
    try:
        grading = grade(Path(sys.argv[1]).resolve())
    except Exception as error:
        grading = {"reward": 0.0, "task_success": False, "checks": [], "hard_failures": [type(error).__name__ + ": " + str(error)]}
    _publish_result(grading)
    print(json.dumps(grading))


if __name__ == "__main__":
    main()
