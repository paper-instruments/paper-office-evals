"""Reference and damaged-output probes for a regenerated editing task."""

import importlib.util
import copy
import shutil
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


HERE = Path(__file__).resolve().parent
TASK = HERE.parent.name
SPEC = importlib.util.spec_from_file_location("task_grader_" + TASK.replace("-", "_"), HERE / "grader.py")
GRADER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = GRADER
SPEC.loader.exec_module(GRADER)


class RevisedTaskTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        fixtures = self.root / "eval_fixtures/pptx"
        fixtures.mkdir(parents=True)
        for path in (HERE.parent / "environment/filesystem").glob("*.pptx"):
            shutil.copyfile(path, fixtures / path.name)
        self.output = self.root / "evals" / TASK / "output.pptx"
        self.output.parent.mkdir(parents=True)
        shutil.copyfile(HERE.parent / "solution/output.pptx", self.output)

    def mutate(self, mutation):
        with ZipFile(self.output) as archive:
            members = {name: archive.read(name) for name in archive.namelist()}
        mutation(members)
        with ZipFile(self.output, "w", ZIP_DEFLATED) as archive:
            for name, value in members.items():
                archive.writestr(name, value)

    def must_fail(self):
        with self.assertRaises((AssertionError, KeyError, IndexError, ValueError)):
            GRADER.verify_revised(self.root, TASK)
        scorecard = GRADER.Scorecard()
        failures = []
        try:
            with GRADER.collect_checks(scorecard, "task"):
                GRADER.verify_revised(self.root, TASK)
        except (AssertionError, KeyError, IndexError, ValueError) as error:
            failures.append(str(error))
        score = GRADER.score_results(GRADER.TASK, scorecard, failures)[0]
        self.assertEqual(score, 0.0)

    def test_reference(self):
        GRADER.validate_pptx_package(self.output)
        GRADER.verify_revised(self.root, TASK)
        scorecard = GRADER.Scorecard()
        with GRADER.collect_checks(scorecard, "task"):
            GRADER.verify_revised(self.root, TASK)
        self.assertEqual(GRADER.score_results(GRADER.TASK, scorecard, [])[0], 1.0)

    def test_secondary_effect_damage(self):
        def mutate(members):
            if TASK == "semantic-bullet-surgery":
                member = "ppt/slides/slide2.xml"
                root = ET.fromstring(members[member])
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Renewal actions")
                size = shape.find("./p:txBody/a:lstStyle/a:lvl1pPr/a:buSzPct", GRADER.NS)
                size.set("val", "120000")
            elif TASK == "merge-safe-table-expansion":
                member = "ppt/slides/slide1.xml"
                root = ET.fromstring(members[member])
                transform = root.find(".//p:graphicFrame/p:xfrm/a:off", GRADER.NS)
                transform.set("x", "2500000")
            elif TASK == "section-aware-slide-lifecycle":
                member = "ppt/presentation.xml"
                root = ET.fromstring(members[member])
                show = root.find("./p:custShowLst/p:custShow/p:sldLst", GRADER.NS)
                show.remove(show[-1])
            else:
                records = GRADER.slide_records(members)
                member = GRADER.rels_member(records[-1]["member"])
                root = ET.fromstring(members[member])
                rel = next(item for item in root if item.get("Type").endswith("/slideLayout"))
                rel.set("Target", "../slideLayouts/slideLayout2.xml")
            members[member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        self.mutate(mutate)
        self.must_fail()

    def test_reviewed_style_or_placement_regression(self):
        def mutate(members):
            record = GRADER.slide_records(members)[-1] if TASK in {"section-aware-slide-lifecycle", "cross-deck-slide-import"} else GRADER.slide_records(members)[1 if TASK == "semantic-bullet-surgery" else 0]
            member = record["member"]
            root = ET.fromstring(members[member])
            if TASK == "semantic-bullet-surgery":
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Renewal actions")
                ppr = shape.find("./p:txBody/a:p/a:pPr", GRADER.NS)
                spacing = ET.Element(GRADER.qn(GRADER.A, "spcBef"))
                ET.SubElement(spacing, GRADER.qn(GRADER.A, "spcPts"), {"val": "24000"})
                ppr.insert(0, spacing)
            elif TASK == "merge-safe-table-expansion":
                cell = root.findall(".//a:tbl/a:tr", GRADER.NS)[5].findall("./a:tc", GRADER.NS)[0]
                cell.find("./a:txBody/a:p/a:pPr/a:defRPr", GRADER.NS).set("sz", "100")
            elif TASK == "section-aware-slide-lifecycle":
                for props in root.findall(".//a:rPr", GRADER.NS):
                    props.set("sz", "100")
            else:
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Growth actions")
                shape.find("./p:spPr/a:xfrm/a:off", GRADER.NS).set("x", "15000000")
            members[member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        self.mutate(mutate)
        self.must_fail()

    def test_import_cannot_crop_wordmark(self):
        if TASK != "cross-deck-slide-import":
            self.skipTest("picture import only")
        def mutate(members):
            member = GRADER.slide_records(members)[-1]["member"]
            root = ET.fromstring(members[member])
            shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Partner wordmark")
            fill = shape.find("./p:blipFill", GRADER.NS)
            rect = fill.find("./a:srcRect", GRADER.NS)
            if rect is None:
                rect = ET.Element(GRADER.qn(GRADER.A, "srcRect"))
                fill.insert(1, rect)
            rect.set("l", "90000")
            members[member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        self.mutate(mutate)
        self.must_fail()

    def test_equivalent_inheritance_or_run_splitting_passes(self):
        def mutate(members):
            record = GRADER.slide_records(members)[-1] if TASK in {"section-aware-slide-lifecycle", "cross-deck-slide-import"} else GRADER.slide_records(members)[1 if TASK == "semantic-bullet-surgery" else 0]
            member = record["member"]
            root = ET.fromstring(members[member])
            if TASK == "cross-deck-slide-import":
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Growth actions")
                props = shape.find("./p:spPr", GRADER.NS)
                props.remove(props.find("./a:xfrm", GRADER.NS))
            elif TASK == "section-aware-slide-lifecycle":
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Review summary")
                tx = shape.find("./p:txBody", GRADER.NS)
                style = tx.find("./a:lstStyle", GRADER.NS)
                level = ET.SubElement(style, GRADER.qn(GRADER.A, "lvl1pPr"))
                original = tx.find("./a:p/a:r/a:rPr", GRADER.NS)
                inherited = copy.deepcopy(original)
                inherited.tag = GRADER.qn(GRADER.A, "defRPr")
                level.append(inherited)
                for run in tx.findall("./a:p/a:r", GRADER.NS):
                    props = run.find("./a:rPr", GRADER.NS)
                    if props is not None:
                        run.remove(props)
            elif TASK == "merge-safe-table-expansion":
                cell = root.findall(".//a:tbl/a:tr", GRADER.NS)[5].findall("./a:tc", GRADER.NS)[0]
                paragraph = cell.find("./a:txBody/a:p", GRADER.NS)
                props = paragraph.find("./a:pPr/a:defRPr", GRADER.NS)
                inherited = copy.deepcopy(props)
                inherited.tag = GRADER.qn(GRADER.A, "rPr")
                paragraph.find("./a:pPr", GRADER.NS).remove(props)
                paragraph.find("./a:r", GRADER.NS).insert(0, inherited)
            else:
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Renewal actions")
                paragraph = shape.find("./p:txBody/a:p", GRADER.NS)
                run = paragraph.find("./a:r", GRADER.NS)
                following = copy.deepcopy(run)
                text = run.find("./a:t", GRADER.NS)
                following.find("./a:t", GRADER.NS).text = text.text[8:]
                text.text = text.text[:8]
                paragraph.insert(list(paragraph).index(run) + 1, following)
            ET.indent(root)
            members[member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        self.mutate(mutate)
        GRADER.validate_pptx_package(self.output)
        GRADER.verify_revised(self.root, TASK)

    def test_unchanged_or_unimported_input_fails(self):
        source_name = GRADER.TASK.inputs[-1].split("/")[-1]
        shutil.copyfile(self.root / "eval_fixtures/pptx" / source_name, self.output)
        self.must_fail()

    def test_old_failure(self):
        def mutate(members):
            if TASK == "semantic-bullet-surgery":
                root = ET.fromstring(members["ppt/slides/slide2.xml"])
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Renewal actions")
                text = shape.find("./p:txBody/a:p/a:r/a:t", GRADER.NS)
                text.text = "\u2022 " + text.text
                member = "ppt/slides/slide2.xml"
            elif TASK == "merge-safe-table-expansion":
                member = "ppt/slides/slide1.xml"
                root = ET.fromstring(members[member])
                cell = root.findall(".//a:tbl/a:tr", GRADER.NS)[2].findall("./a:tc", GRADER.NS)[5]
                fill = cell.find("./a:tcPr/a:solidFill/a:srgbClr", GRADER.NS)
                fill.set("val", "FFFFFF")
            elif TASK == "section-aware-slide-lifecycle":
                member = "ppt/slides/slide5.xml"
                root = ET.fromstring(members[member])
                root.find(".//a:t", GRADER.NS).text = "Appendix: planning assumptions"
            else:
                member = "ppt/presentation.xml"
                root = ET.fromstring(members[member])
                sections = root.findall(".//p14:section", GRADER.NS)
                selected = sections[-1].find("./p14:sldIdLst", GRADER.NS)
                item = selected[-1]
                selected.remove(item)
                sections[1].find("./p14:sldIdLst", GRADER.NS).append(item)
            members[member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        self.mutate(mutate)
        self.must_fail()

    def test_preservation_damage(self):
        def mutate(members):
            if TASK == "semantic-bullet-surgery":
                member = "ppt/slides/slide2.xml"
                root = ET.fromstring(members[member])
                shape = next(s for s in GRADER.iter_shapes(root) if GRADER.shape_name(s) == "Renewal actions")
                shape.findall("./p:txBody/a:p", GRADER.NS)[1].find("./a:r/a:rPr", GRADER.NS).set("i", "0")
            elif TASK == "merge-safe-table-expansion":
                member = "ppt/slides/slide1.xml"
                root = ET.fromstring(members[member])
                run = root.find(".//a:tbl/a:tr/a:tc/a:txBody/a:p/a:r/a:rPr", GRADER.NS)
                run.set("b", "0")
            else:
                member = "ppt/notesSlides/notesSlide1.xml"
                root = ET.fromstring(members[member])
                node = root.find(".//a:t", GRADER.NS)
                node.text = "Unrelated notes were replaced."
            members[member] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        self.mutate(mutate)
        self.must_fail()


if __name__ == "__main__":
    unittest.main()
