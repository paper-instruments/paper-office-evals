"""Outcome checks for the regenerated editing fixtures; uses no pptx package APIs."""

import copy
import io
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile


def one_shape(root, name):
    found = [shape for shape in iter_shapes(root) if shape_name(shape) == name]
    check(len(found) == 1, "effect: expected one shape named %s" % name)
    return found[0]


def slide_by_title(members, title):
    found = [item for item in slide_records(members) if item["title"] == title]
    check(len(found) == 1, "effect: expected one slide titled %s" % title)
    return found[0]


def xfrm(shape):
    for path in ("./p:spPr/a:xfrm", "./p:xfrm", "./p:grpSpPr/a:xfrm"):
        result = shape.find(path, NS)
        if result is not None:
            return result
    return None


def geometry_signature(shape):
    return signature(xfrm(shape))


def slide_relationship_target(members, slide, kind):
    return [target for reltype, target, mode in relationships(members, slide["member"]).values() if reltype == kind and mode != "External"]


def chart_part(members, slide, name):
    ref = one_shape(slide["root"], name).find(".//c:chart", NS)
    check(ref is not None, "effect: %s is not an editable chart" % name)
    return relationships(members, slide["member"])[ref.get(qn(R, "id"))][1]


def notes_part(members, slide):
    found = slide_relationship_target(members, slide, R + "/notesSlide")
    return found[0] if found else None


def signature(node):
    if node is None:
        return None
    node = copy.deepcopy(node)
    for paragraph in node.iter(qn(NS["a"], "p")):
        previous = None
        for run in list(paragraph):
            if run.tag != qn(NS["a"], "r"):
                previous = None
                continue
            props, text = run.find("./a:rPr", NS), run.find("./a:t", NS)
            style = node_signature(props) if props is not None and (props.attrib or len(props)) else None
            if text is not None and previous is not None and previous[0] == style:
                previous[1].text = (previous[1].text or "") + (text.text or "")
                paragraph.remove(run)
            elif text is not None:
                previous = style, text
    return node_signature(node)


def template_chain(members, slide, shape):
    """Return matching layout/master placeholders for these fixture shapes."""
    ph = shape_placeholder(shape)
    layouts = slide_relationship_target(members, slide, R + "/slideLayout")
    if ph is None or not layouts:
        return [], None
    layout_member = layouts[0]
    candidates = [s for s in iter_shapes(xml(members, layout_member)) if shape_placeholder(s) is not None and shape_placeholder(s).get("idx", "0") == ph.get("idx", "0")]
    layout_shape = candidates[0] if len(candidates) == 1 else None
    master_members = [target for kind, target, mode in relationships(members, layout_member).values() if kind == R + "/slideMaster" and mode != "External"]
    master_member = master_members[0] if master_members else None
    master_shape = None
    if master_member:
        wanted_type = shape_placeholder(layout_shape).get("type", "obj") if layout_shape is not None else ph.get("type", "obj")
        family = {"title", "ctrTitle"} if wanted_type in {"title", "ctrTitle"} else {"body", "obj"} if wanted_type in {"body", "obj"} else {wanted_type}
        candidates = [s for s in iter_shapes(xml(members, master_member)) if shape_placeholder(s) is not None and shape_placeholder(s).get("type", "obj") in family]
        master_shape = candidates[0] if len(candidates) == 1 else None
    return [s for s in (master_shape, layout_shape) if s is not None], master_member


def resolved_geometry(members, slide, shape):
    parents, _ = template_chain(members, slide, shape)
    values = {"rot": 0}
    for owner in parents + [shape]:
        transform = xfrm(owner)
        if transform is None:
            continue
        if "rot" in transform.attrib:
            values["rot"] = int(transform.get("rot"))
        for tag, fields in (("off", ("x", "y")), ("ext", ("cx", "cy"))):
            node = transform.find("./a:" + tag, NS)
            if node is not None:
                values.update({key: int(node.get(key)) for key in fields})
    return tuple(values.get(key) for key in ("x", "y", "cx", "cy", "rot"))


def effective_run_styles(members, slide, shape, paragraph):
    """Resolve Latin text styling independently of explicit versus inherited storage."""
    ppr = paragraph.find("./a:pPr", NS)
    level = int(ppr.get("lvl", "0")) + 1 if ppr is not None else 1
    parents, master_member = template_chain(members, slide, shape)
    layout_members = slide_relationship_target(members, slide, R + "/slideLayout")
    if master_member is None and layout_members:
        masters = [target for kind, target, mode in relationships(members, layout_members[0]).values() if kind == R + "/slideMaster" and mode != "External"]
        master_member = masters[0] if masters else None
    theme = None
    if master_member:
        themes = [target for kind, target, mode in relationships(members, master_member).values() if kind == R + "/theme" and mode != "External"]
        theme = xml(members, themes[0]) if themes else None
    sources = [xml(members, "ppt/presentation.xml").find("./p:defaultTextStyle/a:lvl%dpPr/a:defRPr" % level, NS)]
    ph = shape_placeholder(shape)
    category = "titleStyle" if ph is not None and ph.get("type") in {"title", "ctrTitle"} else "bodyStyle" if ph is not None else "otherStyle"
    if master_member:
        sources.append(xml(members, master_member).find("./p:txStyles/p:%s/a:lvl%dpPr/a:defRPr" % (category, level), NS))
    for owner in parents + [shape]:
        tx = owner.find("./p:txBody", NS)
        if tx is None:
            tx = owner.find("./a:txBody", NS)
        if tx is not None:
            sources.append(tx.find("./a:lstStyle/a:lvl%dpPr/a:defRPr" % level, NS))
            if owner is not shape:
                sources.append(tx.find("./a:p/a:pPr/a:defRPr", NS))
    sources.append(paragraph.find("./a:pPr/a:defRPr", NS))

    def style(run):
        result = {"size": None, "font": None, "bold": False, "italic": False, "underline": "none", "color": "000000"}
        for props in sources + [run.find("./a:rPr", NS)]:
            if props is None:
                continue
            for attr, key in (("sz", "size"), ("b", "bold"), ("i", "italic"), ("u", "underline")):
                if attr in props.attrib:
                    value = props.get(attr)
                    result[key] = int(value) if attr == "sz" else value in {"1", "true"} if attr in {"b", "i"} else value
            latin = props.find("./a:latin", NS)
            if latin is not None:
                font = latin.get("typeface")
                if font in {"+mj-lt", "+mn-lt"} and theme is not None:
                    node = theme.find(".//a:%sFont/a:latin" % ("major" if font == "+mj-lt" else "minor"), NS)
                    font = node.get("typeface") if node is not None else None
                result["font"] = font
            rgb = props.find("./a:solidFill/a:srgbClr", NS)
            if rgb is not None:
                result["color"] = rgb.get("val").upper()
        return tuple(sorted(result.items()))
    return {style(run) for run in paragraph.findall("./a:r", NS) if shape_text(run)}


def paragraphs(shape):
    return shape.findall("./p:txBody/a:p", NS)


def character_styles(paragraph):
    defaults = paragraph.find("./a:pPr/a:defRPr", NS)
    result = []
    for run in paragraph.findall("./a:r", NS):
        attrs, children = {}, {}
        for props in (defaults, run.find("./a:rPr", NS)):
            if props is not None:
                attrs.update(props.attrib)
                children.update({child.tag: signature(child) for child in props})
        style = tuple(sorted(attrs.items())), tuple(sorted(children.items()))
        result.extend((char, style) for char in shape_text(run))
    return result


def preserved_members(before, out, allowed=(), removed=(), target_map=None):
    target_map = target_map or {}
    failures = []
    for name in before:
        if name in allowed or name in removed:
            continue
        output_name = target_map.get(name, name)
        if output_name not in out:
            failures.append(name)
        elif name.endswith((".xml", ".rels")):
            if name.endswith(".rels"):
                expected = sorted((kind, target_map.get(target, target), mode) for kind, target, mode in relationships(before, owner_from_rels(name)).values())
                equal = expected == sorted(relationships(out, owner_from_rels(output_name)).values())
            else:
                equal = signature(xml(before, name)) == signature(xml(out, output_name))
            if not equal:
                failures.append(name)
        elif before[name] != out[output_name]:
            failures.append(name)
    check(not failures, "preservation: unrelated package content changed: %r" % failures)


def preserved_slide(before, after, mutable=()):
    copies = copy.deepcopy(before), copy.deepcopy(after)
    for root in copies:
        for parent in root.iter():
            for child in list(parent):
                if child.tag in SHAPE_TAGS and shape_name(child) in mutable:
                    parent.remove(child)
    check(signature(copies[0]) == signature(copies[1]), "preservation: unrelated slide content changed")


def resolved_bullet(shape, paragraph):
    ppr = paragraph.find("./a:pPr", NS)
    level = int(ppr.get("lvl", "0")) if ppr is not None else 0
    inherited = shape.find("./p:txBody/a:lstStyle/a:lvl%dpPr" % (level + 1), NS)
    result = {}
    for props in (inherited, ppr):
        if props is None:
            continue
        for key in ("marL", "indent"):
            if key in props.attrib:
                result[key] = props.get(key)
        for child in props:
            key = local(child.tag)
            if key in {"buNone", "buChar", "buAutoNum", "buBlip"}:
                result["kind"] = key, dict(child.attrib)
            elif key in {"buFont", "buFontTx"}:
                result["font"] = key, dict(child.attrib)
            elif key in {"buSzPct", "buSzPts", "buSzTx"}:
                result["size"] = key, dict(child.attrib)
    size = result.get("size")
    if size:
        run = paragraph.find("./a:r/a:rPr", NS)
        font_size = next((props.get("sz") for props in (
            run, ppr.find("./a:defRPr", NS) if ppr is not None else None,
            inherited.find("./a:defRPr", NS) if inherited is not None else None)
            if props is not None and props.get("sz") is not None), None)
        if size[0] == "buSzPts":
            result["size"] = float(size[1]["val"])
        elif font_size is not None:
            factor = float(size[1]["val"]) / 100000 if size[0] == "buSzPct" else 1
            result["size"] = round(float(font_size) * factor, 6)
    return result


def verify_bullets(before, out):
    original = slide_by_title(before, "Account renewal actions")
    edited = slide_by_title(out, "Account renewal actions")
    reference = one_shape(slide_by_title(before, "Approved list style")["root"], "Approved action list")
    source_shape = one_shape(original["root"], "Renewal actions")
    target_shape = one_shape(edited["root"], "Renewal actions")
    old, new, guide = paragraphs(source_shape), paragraphs(target_shape), paragraphs(reference)
    check(len(new) == 3, "effect: three action-list paragraphs required")
    if len(new) != 3:
        return
    check([shape_text(p) for p in new] == ["Renew the Northstar agreement", "Expand the partner channel", "Review progress on Friday."], "effect: remove pasted prefixes and preserve the action wording")
    for index, paragraph in enumerate(new):
        actual, expected = resolved_bullet(target_shape, paragraph), resolved_bullet(reference, guide[index])
        check(actual.get("kind") == expected.get("kind"), "effect: paragraph %d has the wrong real list marker" % (index + 1))
        if index < 2:
            check(all(actual.get(key) == expected.get(key) for key in ("font", "size", "marL", "indent")), "effect: paragraph %d list appearance does not match the approved style" % (index + 1))
        check(character_styles(paragraph) == character_styles(old[index])[2 if index < 2 else 0:], "preservation: paragraph %d lost wording or emphasis" % (index + 1))
        unrelated_props = []
        for p in (old[index], paragraph):
            props = copy.deepcopy(p.find("./a:pPr", NS))
            if props is not None:
                for child in list(props):
                    if local(child.tag).startswith("bu"):
                        props.remove(child)
                if index < 2:
                    props.attrib.pop("marL", None)
                    props.attrib.pop("indent", None)
            unrelated_props.append(signature(props) if props is not None and (props.attrib or len(props)) else None)
        check(unrelated_props[0] == unrelated_props[1], "preservation: action-list paragraph spacing or unrelated properties changed")
    source_copy, target_copy = copy.deepcopy(source_shape), copy.deepcopy(target_shape)
    for shape in (source_copy, target_copy):
        txbody = shape.find("./p:txBody", NS)
        for paragraph in txbody.findall("./a:p", NS):
            txbody.remove(paragraph)
    check(signature(source_copy) == signature(target_copy), "preservation: action-list geometry or inherited style changed")
    preserved_slide(original["root"], edited["root"], {"Renewal actions"})
    preserved_members(before, out, {original["member"]})


def cell_properties(cell):
    return signature(cell.find("./a:tcPr", NS))


def paragraph_layout(paragraph):
    props = copy.deepcopy(paragraph.find("./a:pPr", NS))
    if props is None:
        return None
    for child in list(props):
        if local(child.tag) == "defRPr":
            props.remove(child)
    for name, default in (("algn", "l"), ("lvl", "0"), ("marL", "0"), ("indent", "0"), ("rtl", "0")):
        if props.get(name) == default:
            props.attrib.pop(name)
    return signature(props) if props.attrib or len(props) else None


def verify_table(before, out):
    original = slide_by_title(before, "Regional forecast")
    edited = slide_by_title(out, "Regional forecast")
    old_shape = one_shape(original["root"], "Regional forecast table")
    new_shape = one_shape(edited["root"], "Regional forecast table")
    old_table, new_table = old_shape.find(".//a:tbl", NS), new_shape.find(".//a:tbl", NS)
    old = [row.findall("./a:tc", NS) for row in old_table.findall("./a:tr", NS)]
    new = [row.findall("./a:tc", NS) for row in new_table.findall("./a:tr", NS)]
    check(len(new) == 7 and all(len(row) == 6 for row in new), "effect: forecast must have seven rows and six columns")
    if len(new) != 7 or any(len(row) != 6 for row in new):
        return
    expected = [["Rolling regional forecast", "", "", "", "", ""], ["Region", "FY26 Q1", "FY26 Q2", "FY26 Q3", "FY26 Q4", "FY27 Q1"], ["North", "12", "14", "16", "18", "20"], ["", "10", "12", "13", "15", "17"], ["South", "9", "11", "12", "14", "16"], ["Central", "7", "8", "9", "10", "11"], ["West", "8", "10", "11", "13", "15"]]
    check([[shape_text(cell) for cell in row] for row in new] == expected, "effect: forecast values or heading wording differ")
    merge_state = lambda cell: (int(cell.get("gridSpan", "1")), int(cell.get("rowSpan", "1")), cell.get("hMerge", "0") in ("1", "true"), cell.get("vMerge", "0") in ("1", "true"))
    wanted = {(0, 0): (6, 1, False, False), (2, 0): (1, 2, False, False), (3, 0): (1, 1, False, True)}
    wanted.update({(0, col): (1, 1, True, False) for col in range(1, 6)})
    check(all(merge_state(cell) == wanted.get((row, col), (1, 1, False, False)) for row, cells in enumerate(new) for col, cell in enumerate(cells)), "effect: merged heading or North label is incorrect")
    old_widths = [int(col.get("w")) for col in old_table.findall("./a:tblGrid/a:gridCol", NS)]
    widths = [int(col.get("w")) for col in new_table.findall("./a:tblGrid/a:gridCol", NS)]
    old_heights = [int(row.get("h")) for row in old_table.findall("./a:tr", NS)]
    heights = [int(row.get("h")) for row in new_table.findall("./a:tr", NS)]
    check(widths == old_widths + [old_widths[-1]] and heights == old_heights[:5] + [old_heights[4]] + old_heights[5:], "effect: row heights or column widths changed")
    transform = xfrm(new_shape)
    off, ext = transform.find("./a:off", NS), transform.find("./a:ext", NS)
    size = xml(out, "ppt/presentation.xml").find("./p:sldSz", NS)
    check(int(off.get("x")) >= 0 and int(off.get("y")) >= 0 and int(off.get("x")) + int(ext.get("cx")) <= int(size.get("cx")) and int(off.get("y")) + int(ext.get("cy")) <= int(size.get("cy")) and int(ext.get("cx")) == sum(widths) and int(ext.get("cy")) == sum(heights), "effect: expanded table is clipped or has incoherent frame dimensions")
    retained_styles, retained_text, new_styles, new_typography = [], [], [], []
    for row, cells in enumerate(new):
        old_row = 4 if row == 5 else row - (row > 5)
        for col, cell in enumerate(cells):
            template = old[old_row][min(col, 4)]
            if row == 5 or col == 5:
                new_styles.append(cell_properties(cell) == cell_properties(template))
                if shape_text(cell):
                    old_paragraphs, new_paragraphs = template.findall("./a:txBody/a:p", NS), cell.findall("./a:txBody/a:p", NS)
                    desired = set().union(*(effective_run_styles(before, original, template, p) for p in old_paragraphs))
                    actual = set().union(*(effective_run_styles(out, edited, cell, p) for p in new_paragraphs))
                    layouts_match = [paragraph_layout(p) for p in old_paragraphs] == [paragraph_layout(p) for p in new_paragraphs]
                    new_typography.append(bool(desired) and actual == desired and layouts_match)
            else:
                retained_styles.append(cell_properties(cell) == cell_properties(template))
                retained_text.append(signature(cell.find("./a:txBody", NS)) == signature(template.find("./a:txBody", NS)))
    check(all(new_styles), "effect: new row or column lost the template cell formatting")
    check(all(new_typography), "effect: new row or column text does not match the template typography")
    check(all(retained_styles) and all(retained_text), "preservation: existing cells lost text or direct formatting")
    check(signature(new_table.find("./a:tblPr", NS)) == signature(old_table.find("./a:tblPr", NS)), "preservation: table style changed")
    preserved_slide(original["root"], edited["root"], {"Regional forecast table"})
    preserved_members(before, out, {original["member"]})


def section_records(members):
    return [(section.get("id"), section.get("name"), [int(item.get("id")) for item in section.findall("./p14:sldIdLst/p14:sldId", NS)]) for section in xml(members, "ppt/presentation.xml").findall(".//p14:section", NS)]


def custom_shows(members):
    by_rid = {item["rid"]: item["id"] for item in slide_records(members)}
    return {show.get("name"): [by_rid.get(item.get(qn(R, "id"))) for item in show.findall("./p:sldLst/p:sld", NS)] for show in xml(members, "ppt/presentation.xml").findall("./p:custShowLst/p:custShow", NS)}


def verify_lifecycle(before, out):
    original, records = slide_records(before), slide_records(out)
    titles = ["Northstar leadership review", "Financial summary", "Regional outlook", "Decisions and next steps", "Appendix: planning assumptions"]
    check([item["title"] for item in records] == titles, "effect: leadership-review slide order or appendix is incorrect")
    if len(records) != 5:
        return
    originals = {item["title"]: item for item in original}
    deleted = originals["Superseded risk assessment"]
    survivors = [item for item in original if item is not deleted]
    current = {item["title"]: item for item in records}
    unchanged = all(item["title"] in current and current[item["title"]]["id"] == item["id"] and signature(item["root"]) == signature(current[item["title"]]["root"]) for item in survivors)
    check(unchanged, "preservation: a surviving slide or its identity changed")
    appendix = records[-1]
    texts = [shape_text(shape) for shape in iter_shapes(appendix["root"])]
    check("Forecast uses signed contracts and qualified pipeline.Currency rates are held at the September planning baseline." in texts or all(any(text in value for value in texts) for text in ("Forecast uses signed contracts and qualified pipeline.", "Currency rates are held at the September planning baseline.")), "effect: appendix assumptions are missing")
    check(appendix["id"] not in {item["id"] for item in original}, "effect: appendix reuses an existing slide identity")
    cover = originals["Northstar leadership review"]
    cover_shapes = list(iter_shapes(cover["root"]))
    title_template = next(s for s in cover_shapes if shape_text(s) == cover["title"])
    body_template = one_shape(cover["root"], "Review summary")
    title_styles = set().union(*(effective_run_styles(before, cover, title_template, p) for p in paragraphs(title_template)))
    body_styles = set().union(*(effective_run_styles(before, cover, body_template, p) for p in paragraphs(body_template)))
    style_matches = []
    for shape in iter_shapes(appendix["root"]):
        for paragraph in paragraphs(shape):
            text = shape_text(paragraph)
            if not text:
                continue
            wanted = title_styles if text == appendix["title"] else body_styles
            style_matches.append(bool(wanted) and effective_run_styles(out, appendix, shape, paragraph) == wanted)
    check(bool(style_matches) and all(style_matches), "effect: appendix typography does not match the deck title-and-body style")
    before_sections = section_records(before)
    groups = {"Opening": [records[0]["id"]], "Review": [records[1]["id"], records[2]["id"]],
              "Close": [records[3]["id"], appendix["id"]]}
    expected = [(ident, name, groups[name]) for ident, name, values in before_sections]
    check(section_records(out) == expected, "effect: section memberships do not match the requested agenda")
    expected_shows = {name: [value for value in values if value != deleted["id"]] for name, values in custom_shows(before).items()}
    check(custom_shows(out) == expected_shows, "effect: custom-show memberships changed incorrectly")
    removed = {deleted["member"], rels_member(deleted["member"])}
    for kind, target, mode in relationships(before, deleted["member"]).values():
        if kind == R + "/notesSlide":
            removed.update((target, rels_member(target)))
    check(not any(item["title"] == deleted["title"] or item["root"].get("show", "1") in {"0", "false"} for item in records), "effect: obsolete hidden material remains")
    target_map = {}
    for item in survivors:
        if item["title"] in current:
            destination = current[item["title"]]["member"]
            target_map[item["member"]] = destination
            target_map[rels_member(item["member"])] = rels_member(destination)
    preserved_members(before, out, {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"}, removed, target_map)
    # Check every visible appendix shape fits; no particular new shape/part ID is required.
    size = xml(out, "ppt/presentation.xml").find("./p:sldSz", NS)
    fits = []
    for shape in iter_shapes(appendix["root"]):
        transform = xfrm(shape)
        if transform is not None:
            off, ext = transform.find("./a:off", NS), transform.find("./a:ext", NS)
            if off is not None and ext is not None:
                fits.append(int(off.get("x")) >= 0 and int(off.get("y")) >= 0 and int(off.get("x")) + int(ext.get("cx")) <= int(size.get("cx")) and int(off.get("y")) + int(ext.get("cy")) <= int(size.get("cy")))
    check(all(fits), "effect: appendix content lies outside the slide")


def workbook_signature(payload):
    with ZipFile(io.BytesIO(payload)) as archive:
        return {name: signature(ET.fromstring(archive.read(name))) if name.endswith(".xml") else archive.read(name) for name in archive.namelist() if name != "docProps/core.xml"}


def verify_import(source, before, out):
    original, records = slide_records(before), slide_records(out)
    check(len(records) == len(original) + 1 and records[-1]["title"] == "Partner growth plan", "effect: partner plan was not appended exactly once")
    if len(records) != len(original) + 1:
        return
    check(all(left["id"] == right["id"] and signature(left["root"]) == signature(right["root"]) for left, right in zip(original, records[:-1])), "preservation: existing destination slides changed")
    imported = records[-1]
    check(imported["root"].get("show", "1") not in {"0", "false"}, "effect: required imported slide is hidden")
    source_slide = slide_by_title(source, "Partner growth plan")
    financial = slide_by_title(before, "Financial summary")
    expected_layout = slide_relationship_target(before, financial, R + "/slideLayout")
    check(slide_relationship_target(out, imported, R + "/slideLayout") == expected_layout, "effect: imported slide uses a different house layout than Financial summary")
    placeholders = {int(shape_placeholder(shape).get("idx", "0")): shape for shape in iter_shapes(imported["root"]) if shape_placeholder(shape) is not None}
    check(0 in placeholders and 4 in placeholders and shape_text(placeholders[4]) == "Expand the partner channelLaunch the joint account plan" and (8 not in placeholders or not shape_text(placeholders[8])), "effect: growth actions are not in the wide narrative area with an empty sidebar")
    if 4 in placeholders:
        source_body = one_shape(source_slide["root"], "Growth actions")
        check([character_styles(p) for p in paragraphs(placeholders[4])] == [character_styles(p) for p in paragraphs(source_body)], "preservation: imported growth actions lost their emphasis, colors or text size")
        expected_area = resolved_geometry(before, financial, one_shape(financial["root"], "Main narrative"))
        actual_area = resolved_geometry(out, imported, placeholders[4])
        check(None not in expected_area and actual_area == expected_area, "effect: growth actions do not occupy the house narrative area")
    source_title = next(shape for shape in iter_shapes(source_slide["root"]) if shape_placeholder(shape) is not None and shape_placeholder(shape).get("type") in ("title", "ctrTitle"))
    if 0 in placeholders:
        check([character_styles(p) for p in paragraphs(placeholders[0])] == [character_styles(p) for p in paragraphs(source_title)], "preservation: imported title formatting changed")
    sections_before, sections_after = section_records(before), section_records(out)
    selected = next(ident for ident, name, values in sections_before if financial["id"] in values)
    expected_sections = [(ident, name, values + ([imported["id"]] if ident == selected else [])) for ident, name, values in sections_before]
    check(sections_after == expected_sections, "effect: imported slide is in the wrong closing section")
    for name in ("Regional pipeline", "Partner wordmark"):
        left, right = one_shape(source_slide["root"], name), one_shape(imported["root"], name)
        check(geometry_signature(left) == geometry_signature(right), "preservation: imported %s moved or resized" % name)
        if name == "Partner wordmark":
            ignored = {"id", "name", qn(R, "embed"), qn(R, "link"), qn(R, "id")}
            check(node_signature(left, ignore_attributes=ignored) == node_signature(right, ignore_attributes=ignored), "preservation: imported partner wordmark crop or picture style changed")
    left_chart, right_chart = chart_part(source, source_slide, "Regional pipeline"), chart_part(out, imported, "Regional pipeline")
    check(signature(xml(source, left_chart)) == signature(xml(out, right_chart)), "effect: imported editable chart data or formatting changed")
    left_book, right_book = workbook_part(source, left_chart), workbook_part(out, right_chart)
    check(left_book is not None and right_book is not None and workbook_signature(source[left_book]) == workbook_signature(out[right_book]), "effect: imported chart workbook changed or disappeared")
    left_image = one_shape(source_slide["root"], "Partner wordmark").find(".//a:blip", NS)
    right_image = one_shape(imported["root"], "Partner wordmark").find(".//a:blip", NS)
    if left_image is not None and right_image is not None:
        source_media = relationships(source, source_slide["member"])[left_image.get(qn(R, "embed"))][1]
        output_media = relationships(out, imported["member"])[right_image.get(qn(R, "embed"))][1]
        check(source[source_media] == out[output_media], "effect: imported partner wordmark changed")
    else:
        check(False, "effect: imported partner wordmark is missing")
    left_notes, right_notes = notes_part(source, source_slide), notes_part(out, imported)
    check(left_notes is not None and right_notes is not None and signature(xml(source, left_notes)) == signature(xml(out, right_notes)), "preservation: imported speaker notes or notes formatting changed")
    preserved_members(before, out, {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "[Content_Types].xml"})


def verify(root, task, namespace):
    for name in ("check", "package", "xml", "node_signature", "local", "shape_text", "slide_records", "iter_shapes", "shape_name", "shape_placeholder", "relationships", "owner_from_rels", "rels_member", "workbook_part", "qn", "R", "NS", "SHAPE_TAGS"):
        globals()[name] = namespace[name]
    names = {"semantic-bullet-surgery": "eval_fixture_02_text.pptx", "merge-safe-table-expansion": "eval_fixture_03_objects.pptx", "section-aware-slide-lifecycle": "eval_fixture_04_structure.pptx", "cross-deck-slide-import": "eval_fixture_pair_destination.pptx"}
    hashes = json.loads(Path(__file__).with_name("fixture_hashes.json").read_text())
    for name, expected in hashes.items():
        check(namespace["file_hash"](root / "eval_fixtures/pptx" / name) == expected, "preservation: input fixture changed")
    before = package(root / "eval_fixtures/pptx" / names[task])
    out = package(root / "evals" / task / "output.pptx")
    if task == "cross-deck-slide-import":
        verify_import(package(root / "eval_fixtures/pptx/eval_fixture_pair_source.pptx"), before, out)
    else:
        {"semantic-bullet-surgery": verify_bullets, "merge-safe-table-expansion": verify_table, "section-aware-slide-lifecycle": verify_lifecycle}[task](before, out)
