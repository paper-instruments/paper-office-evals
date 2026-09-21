# Paper API routing and contracts

Use this reference to select and call a public Paper PPTX operation. The contracts target current `paper-pptx` 0.2.x, imported through `pptx`. Ordinary upstream authoring does not require this reference; consult the relevant contract before a Paper-specific call when the signature, return value, or refusal boundary is uncertain.

## Contents

- Inspect and target
- Text and effective formatting
- Notes and footers
- Slide and shape lifecycle
- Tables, images, and charts
- Layouts and composition
- Save and compare
- Result objects
- Errors and refusals

## Inspect and target

| Need | Public operation | Important boundary |
|---|---|---|
| Survey a deck | `inspect_deck(prs)` | Explicit geometry and structural facts; inherited values may be `None` |
| Survey slide text | `inspect_text(slide)` | Includes groups and table cells; chart text and some effective formatting remain blind |
| Resolve run appearance | `effective_font(run)` | Supported slide shapes only; returns value, resolution state, and provenance |
| Resolve paragraph appearance | `effective_paragraph_format(paragraph)` | Supported slide-shape paragraphs only |
| Resolve shape appearance | `effective_shape_format(shape)` | Some theme format-scheme effects remain honestly unresolved |
| Target nested content by name | `shape_by_name`, `picture_by_name`, `table_by_name`, `chart_by_name` | Requires uniqueness; raises rather than selecting the first duplicate |

Import inspectors from `pptx.inspect`. Inspection payloads have deterministic `.to_dict()` forms.

```text
inspect_deck(prs) -> DeckManifest
inspect_text(slide) -> TextInspection
effective_font(run) -> EffectiveFont
effective_paragraph_format(paragraph) -> EffectiveParagraphFormat
effective_shape_format(shape) -> EffectiveShapeFormat
```

`inspect_deck()` does not currently expose complete chart-part/workbook ownership. When that graph is essential, use the operation's own safe API if possible. Do not infer ownership from shape names.

## Text and effective formatting

| Intent | Operation | Boundary |
|---|---|---|
| Replace one inspected block | `replace_text_at(prs, anchor, find, replace)` | Refuses stale, absent, or boundary-crossing targets |
| Replace literal text deck-wide | `replace_text(prs, find, replace, include_notes=False)` | Case-sensitive; does not cross paragraphs, line breaks, or fields |
| Recover a shifted anchor | `refind(prs, anchor)` | Current anchors require a unique fingerprint within the original structural container; legacy anchors require uniqueness within the part |
| Set real bullets | `paragraph.bullet.set_character`, `.set_numbered`, `.set_none` | Setting a new kind replaces an existing picture bullet |
| Freeze supported autofit | `text_frame.normalize_autofit(min_font_size=None, resolve=False)` | Unsupported containers or unresolved spacing can refuse |
| Add live fields | `paragraph.add_slide_number_field()`, `.add_datetime_field(format_code)` | PowerPoint calculates the displayed value; Paper stores cached text |

`ReplaceResult.replacements` contains the occurrence count. Post-edit anchors are in `ReplaceResult.blocks`; do not reuse the baseline anchor after mutation.

Import `replace_text`, `replace_text_at`, and `refind` from `pptx.edit`:

```text
replace_text(prs, find, replace, *, include_notes=False) -> ReplaceResult
replace_text_at(prs, anchor, find, replace) -> ReplaceResult
refind(prs, anchor) -> BlockAnchor
```

Zero matches from `replace_text()` are a normal result. Current anchors locate a shape or table cell before matching a full paragraph fingerprint. Keep the original anchor object and use `refind()` after paragraph insertion within that container.

Use `with prs.batch():` for a group of edits that should validate and roll back together. Save after the block exits. An exception rolls back the whole block.

`effective_paragraph_format()` resolves bullet kind, `bullet_font` and `bullet_size` independently through the template chain. Font and size can follow the text; preserve unresolved values in inspection reports.

Never fill an unresolved effective value with a generic “professional default.” Preserve the unresolved token, transform, or script boundary in the report.

## Notes and footers

| Intent | Operation | Boundary |
|---|---|---|
| Read existing notes | `slide.read_notes_text()` | Does not create a notes part |
| Replace existing notes | `slide.replace_notes_text(text)` | Refuses when no supported existing notes body exists; rebuilds the body with limited formatting preservation |
| Apply footer/date/slide number | `prs.apply_footers(...)` or `slide.apply_footers(...)` | Each call sets the complete footer state; missing or disabled layout furniture can refuse |

Avoid accessing notes through APIs that create a notes slide when the task is read-only or “existing notes only.” Notes replacement is validated, but it is not wrapped in the same package transaction used by every guarded structural operation.

Footer arguments are keyword-only, and footer calls return `None`:

```text
Presentation.apply_footers(*, footer=None, slide_number=False, date_format=None,
                           fixed_date=None, skip_title_slides=False, now=None) -> None
Slide.apply_footers(*, footer=None, slide_number=False, date_format=None,
                    fixed_date=None, now=None) -> None
Slide.read_notes_text() -> str
Slide.replace_notes_text(text) -> None
```

## Slide and shape lifecycle

| Intent | Operation | Ownership behavior |
|---|---|---|
| Clone a slide | `prs.slides.clone(source, after=None, policy=None)` | Deep-copies charts/workbooks and notes by default; shares media |
| Delete/move/reorder slides | `delete`, `move`, `reorder` on `prs.slides` | Maintains slide order, sections, custom shows, and reachable parts where supported |
| Copy a same-package shape | `slide.shapes.add_copy(shape)` | Fresh shape IDs; shares media; deep-copies charts/workbooks |
| Delete/move a shape | `slide.shapes.delete`, `.move` | Requires a direct member, not a nested group child |

Unknown or prohibited relationships can make lifecycle operations refuse before mutation. Keep the source object live and from the same package.

```text
Slides.clone(source, *, after=None, policy=None) -> Slide
Slides.delete(slide) -> None
Slides.move(slide, to_index) -> None
Slides.reorder(new_order) -> None
SlideShapes.add_copy(shape) -> Shape
SlideShapes.delete(shape) -> None
SlideShapes.move(shape, to_index) -> None
```

`source`, `after`, `slide`, and `shape` accept the forms documented by the owning collection; do not pass a proxy from another presentation. Clone and copy return new live objects. Reacquire collections and targets after deletion or graph replacement.

## Tables, images, and charts

| Intent | Operation | Boundary |
|---|---|---|
| Insert/delete table rows | `Table.insert_row`, `delete_row` | Maintains a rectangular grid; refuses edits that cut a merge |
| Insert/delete columns | `Table.insert_column`, `delete_column` | `after=-1` prepends; deleting the last row/column is invalid |
| Replace an embedded image | `Picture.replace_image(file, allow_format_change=False)` | Preserves geometry, crop, rotation, and mask; format changes require opt-in |
| Replace category-chart data | `Chart.replace_data_safe(categories, series, number_format=None)` | Refuses shared chart/workbook ownership, combos, unsupported types, invalid data, and stale proxies |

For charts, keep categories, every series, formulas/caches, and embedded workbook state coherent. Use `replace_data_safe()` instead of editing cache XML alone. A workbookless supported chart remains workbookless.

```text
Table.insert_row(after, *, copy_format_from=None) -> row proxy
Table.delete_row(row_idx) -> None
Table.insert_column(after, *, width=None, copy_format_from=None) -> column proxy
Table.delete_column(col_idx) -> None
Cell.extend_merge(other_cell) -> None
Picture.replace_image(image_file, *, allow_format_change=False) -> None
Chart.replace_data_safe(categories, series, *, number_format=None) -> None
```

Table indexes are zero-based. For insertion, `after=-1` prepends. Column `copy_format_from` refers to the pre-insertion column and copies direct cell properties without copying text or merge state. `extend_merge()` grows an existing merged rectangle right or down and retains its content. `series` for `replace_data_safe()` is a sequence of `(series_name, values)` pairs.

## Layouts and composition

| Intent | Operation | Choice to make |
|---|---|---|
| Rebind within a package | `slide.rebind_layout(target_layout, placeholder_map="auto", orphan_policy="refuse")` | Review placeholder mapping, baked orphans, and effective-format shifts |
| Import one slide | `prs.import_slide(source_prs, slide, mode=..., ...)` | Choose `adopt_theme`, `keep_appearance`, or `bake` explicitly |
| Append a deck | `prs.append_deck(source_prs, mode=..., notes=True)` | Validates the complete source before writing; source sections are not copied |

Composition modes:

- `adopt_theme`: bind content to a destination layout and report effective-format shifts.
- `keep_appearance`: copy and deduplicate the required source design chain.
- `bake`: localize safely resolvable appearance, convert remaining placeholders, and accept reported drops or refusals.

Require identical page dimensions before composition. Do not manually scale a package graph as an incidental side effect.

```text
Slide.rebind_layout(target_layout, *, placeholder_map="auto",
                    orphan_policy="refuse") -> RebindReport
Presentation.import_slide(source_prs, slide, *, mode, position=None, notes=True,
                          section=None, section_id=None, target_layout=None,
                          placeholder_map="auto") -> ImportReport
Presentation.append_deck(source_prs, *, mode, notes=True) -> tuple[ImportReport, ...]
```

`mode` is required and must be `adopt_theme`, `keep_appearance`, or `bake`. A foreign `target_layout` passed to `rebind_layout()` is a caller `ValueError`, not a `PaperRefusal`.

Automatic layout and placeholder selection require unique matches. Resolve competing layouts with `target_layout` and adopt-theme placeholder choices with a partial `placeholder_map`. A mapping value of `None` deliberately orphans and bakes that placeholder. Choose a section by unique exact name or by its stored `section_id`; use only one selector. Whole-deck append uses automatic selection.

## Save and compare

| Intent | Operation | Boundary |
|---|---|---|
| Compare package members | `diff_package(a, b)` | XML-semantic and binary-byte comparison; not a visual diff |
| Preserve unchanged original XML | `patch_save(original, prs, output)` | Retains bytes of semantically unchanged members, including equivalent relationship and content-type registries |
| Compare presentation semantics | `diff_decks(a, b, detail="structure"|"text"|"full")` | Permanent slide IDs are reliable matching keys only for lineage-related decks |
| Compare XML fragments | `xml_equivalent(a, b)` | Generic XML semantics, not all OPC relationship/content-type semantics |

For cleanup, identify the requested categories and use the owning notes, comments, core-properties, slide or layout operations.

```text
xml_equivalent(a, b) -> bool
diff_package(path_a, path_b) -> PackageDiff
patch_save(original_path, document, out_path) -> PackageDiff
diff_decks(path_a, path_b, *, detail="structure") -> DeckDiff
```

Import package operations from `pptx.package` and `diff_decks` from `pptx.diff`. `detail` must be `structure`, `text`, or `full`.

Diff schema v5 uses stable shape identity and exact text snapshots. It reports one changed middle region after equal prefixes and suffixes; it does not infer paragraph edit history. Full detail includes bullet-only changes and conservative formatting shifts for uniquely aligned unchanged paragraphs. Table dimension changes are factual row and column counts.

## Result objects

Use Python attributes while operating and `.to_dict()` for deterministic serialization. Do not guess `.count`, `.changes`, or generic report fields.

| Result | Exact public attributes that commonly matter |
|---|---|
| `ReplaceResult` | `replacements`, `blocks` |
| `RebindReport` | `source_layout`, `source_layout_name`, `target_layout`, `target_layout_name`, `placeholder_map_used`, `baked_orphans`, `run_shifts` |
| `ImportReport` | `mode`, `source_slide`, `dest_slide`, `dest_slide_id`, `position`, `layout_binding`, `layout_binding_method`, `placeholder_map_used`, `parts_added`, `parts_reused`, `notes_copied`, `comments_dropped`, `section`, `section_id`, `baked_shapes`, `dropped_placeholders`, `run_shifts` |
| `PackageDiff` | `deltas`, `is_empty`; each `PartDelta` has `partname`, `kind`, `change`, `detail` |
| `DeckDiff` | `detail`, `slides_added`, `slides_removed`, `slides_moved`, `slide_changes`, `package_changes`, `is_empty` |

Serialized reports add their own `schema` and `version`.

## Errors and refusals

Catch `PaperRefusal` only when the workflow needs to record a conservative stop:

```python
from pptx.errors import PaperRefusal

try:
    # supported Paper operation
    ...
except PaperRefusal as exc:
    evidence = {"exception_class": type(exc).__name__, "message": str(exc)}
    raise
```

Public subclasses are `PackageLimitError`, `AmbiguousTargetError`, `TargetNotFoundError`, `StaleAnchorError`, `UnsupportedStructureError`, `BoundaryViolationError`, and `RelationshipPolicyError`, all imported from `pptx.errors`. Catch the base class only when the workflow needs one conservative-stop path; catch a subclass when recovery differs by cause. Invalid values, ranges, types, dimensions, or permutations generally raise `ValueError` or `TypeError`; fix the caller rather than presenting them as document-safety refusals.

If rollback itself raises unexpectedly, stop using that in-memory presentation. Restoration is not established.
