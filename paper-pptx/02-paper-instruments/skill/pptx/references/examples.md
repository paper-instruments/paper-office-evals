# Examples

These patterns are intentionally small. Adapt targeting, content, layout, and validation to the task rather than copying every step into every workflow.

## Contents

- Create from a template
- Make a preservation-sensitive text edit
- Update a table and chart
- Clone and reorder safely
- Import with an explicit design policy
- Report effective formatting honestly

## 1. Create from a template

Use the template's layouts and theme instead of rebuilding its design. Search all masters when the template has more than one.

```python
from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

source = "company-template.pptx"
output = "operating-review.pptx"
prs = Presentation(source)

layouts = [layout for master in prs.slide_masters for layout in master.slide_layouts]
layout = next(layout for layout in layouts if layout.name == "Title and Content")
slide = prs.slides.add_slide(layout)
slide.shapes.title.text = "Retention improved, but expansion remains below plan"

body = next(
    placeholder
    for placeholder in slide.placeholders
    if placeholder.placeholder_format.type
    in {PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT}
)
body.name = "body_retention_drivers"
tf = body.text_frame
tf.clear()
for index, text in enumerate(
    [
        "Enterprise renewal rate increased to 94%.",
        "Expansion slowed in accounts without executive sponsorship.",
        "Assign sponsors to the top 20 expansion opportunities this month.",
    ]
):
    paragraph = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
    paragraph.text = text
    paragraph.level = 0

prs.save(output)
Presentation(output)  # fresh reopen; add render review for a created deck
```

This example inherits the template's body formatting and geometry. If no compatible placeholder exists, inspect the intended content region before adding a text box. Use direct formatting only when the design calls for it.

## 2. Make one preservation-sensitive text edit

Target an inspected block, retain the full baseline, and use narrow save only when package churn matters.

```python
from pptx import Presentation
from pptx.edit import replace_text_at
from pptx.inspect import inspect_text
from pptx.package import patch_save

source = "quarterly-review.pptx"
output = "quarterly-review-updated.pptx"
prs = Presentation(source)

inspection = inspect_text(prs.slides[0])
baseline = "Q2 2027 Business Review"
matches = [block for block in inspection.blocks if block.text == baseline]
if len(matches) != 1:
    raise RuntimeError(f"expected one exact title, found {len(matches)}")

result = replace_text_at(
    prs,
    matches[0].anchor,
    baseline,
    "Q3 2027 Business Review",
)
assert result.replacements == 1
save_report = patch_save(source, prs, output)

reopened = Presentation(output)
assert any(
    block.text == "Q3 2027 Business Review"
    for block in inspect_text(reopened.slides[0]).blocks
)
print(save_report.to_dict())
```

If normalized serialization is acceptable, use `prs.save(output)` instead of `patch_save()`.

## 3. Update a table and chart by semantic name

Confirm names are unique, use guarded object operations, and validate the reopened objects.

```python
from pptx import Presentation

source = "board-pack.pptx"
output = "board-pack-updated.pptx"
prs = Presentation(source)
slide = prs.slides[3]

table = slide.shapes.table_by_name("table_regional_revenue")
row = table.insert_row(len(table.rows) - 1, copy_format_from=len(table.rows) - 1)
for cell, value in zip(row.cells, ("Central", "$18.2M", "+7%")):
    cell.text = value

chart = slide.shapes.chart_by_name("chart_revenue_by_region")
chart.replace_data_safe(
    ["North", "South", "Central"],
    [("FY27", (24.0, 20.5, 18.2)), ("Plan", (23.0, 21.0, 19.0))],
    number_format="0.0",
)

prs.save(output)
fresh = Presentation(output)
fresh_slide = fresh.slides[3]
assert fresh_slide.shapes.table_by_name("table_regional_revenue").rows[-1].cells[0].text == "Central"
fresh_chart = fresh_slide.shapes.chart_by_name("chart_revenue_by_region")
assert [series.name for series in fresh_chart.series] == ["FY27", "Plan"]
```

Also verify categories, values, formulas/caches/workbook agreement, neighboring charts, and rendering when they are part of the acceptance criteria.

## 4. Clone and reorder safely

Use the slide lifecycle API so charts, notes, media, and relationships follow the declared policy.

```python
from pptx import Presentation
from pptx.slide import SlideClonePolicy

prs = Presentation("operating-plan.pptx")
source = prs.slides[2]
clone = prs.slides.clone(
    source,
    after=source,
    policy=SlideClonePolicy(
        deep_copy_charts=True,
        deep_copy_notes=True,
        share_media=True,
    ),
)
clone.shapes.title.text = "Downside scenario requires two additional actions"
prs.slides.move(clone, 3)
prs.save("operating-plan-with-downside.pptx")
```

After reopen, verify slide order, unique shape IDs, independent chart/workbook mutation, intended media sharing, notes, internal links, and rendered appearance.

## 5. Import a slide with an explicit design policy

Do not paste raw slide XML. Choose whether destination design or source appearance should win.

```python
from pptx import Presentation

destination = Presentation("company-template.pptx")
source = Presentation("research-deck.pptx")
if (destination.slide_width, destination.slide_height) != (
    source.slide_width,
    source.slide_height,
):
    raise ValueError("slide dimensions differ; rescaling requires a design decision")

report = destination.import_slide(
    source,
    source.slides[4],
    mode="keep_appearance",
    notes=True,
)
destination.save("combined-deck.pptx")
print(report.to_dict())
```

Use `adopt_theme` when the imported content should join the destination design system. Use `bake` only after accepting reported appearance localization and unsupported boundaries.

## 6. Produce an honest effective-format report

Keep local and effective values separate. Never substitute a fallback for an unresolved value.

```python
from pptx import Presentation
from pptx.inspect import effective_font

prs = Presentation("incoming.pptx")
shape = prs.slides[1].shapes.shape_by_name("body_key_message")
run = shape.text_frame.paragraphs[0].runs[0]

report = {
    "text": run.text,
    "local": {
        "size_pt": run.font.size.pt if run.font.size is not None else None,
        "latin_typeface": run.font.name,
        "bold": run.font.bold,
        "italic": run.font.italic,
        "underline": run.font.underline,
    },
    "effective": effective_font(run).to_dict(),
}
print(report)
```

If the task requires a particular JSON contract, follow its exact property names and units. API truth and artifact-schema compliance are separate requirements.
