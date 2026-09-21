# Editing and structural safety

Use this reference for brownfield edits, ownership-sensitive changes, layout work, composition, and exceptional OOXML.

## Contents

- Preserve scope and target from evidence
- Text and effective formatting
- Tables, charts, and images
- Slide lifecycle and composition
- Custody surfaces
- Exceptional OOXML

## Preserve the user's scope

Before mutation, define the expected delta:

- target slides and objects;
- content and properties allowed to change;
- content and features required to remain unchanged;
- whether normalized serialization is acceptable;
- source and destination paths; and
- any supported refusal expected from adversarial or unsupported input.

Do not confuse “minimal visible change” with “minimal package change.” Establish which one matters.

## Target from evidence

Prefer this order:

1. a unique semantic shape name confirmed by inspection;
2. an inspected text block plus `BlockAnchor` and retained full baseline text;
3. an explicit slide/object index whose stability is part of the task; or
4. a carefully validated structural predicate.

Avoid coordinates, generated shape names, substring-only matching, or the first item returned when a deck can contain duplicates.

After a structural mutation, reacquire affected proxies and anchors. Do not reuse stale objects after deletion, cross-package import, or graph replacement.

## Text edits

Use anchored replacement when formatting, fields, and run boundaries must survive. Use deck-wide replacement only when the requested scope is genuinely deck-wide.

Text can live in ordinary shapes, nested groups, table cells, fields, notes, chart parts, SmartArt, and unsupported alternate content. `inspect_text()` reports blind regions; it does not imply that every visible string is editable.

Preserve:

- paragraph boundaries and levels;
- live fields and their types;
- untouched runs and local formatting;
- language, hyperlinks, and meaningful whitespace; and
- existing notes-part presence when the task forbids creating notes.

Re-inspect after replacement and assert the intended replacement count. Zero matches is a result to interpret, not automatically an exception.

## Effective formatting

Local `run.font` values can be `None` because PowerPoint inherits appearance from paragraph defaults, list styles, placeholders, layouts, masters, theme fonts, and color maps.

Use Paper's effective inspectors when displayed appearance determines the edit. Keep local and effective values separate. For each effective value, retain:

- value;
- resolved state; and
- ordered provenance.

Do not flatten a theme token, gradient, transformed scheme color, or East Asian/complex-script selection into a guessed RGB or Latin font. If an operation requires a value that remains unresolved, refuse, preserve, or obtain explicit user direction.

## Tables

Use table row/column APIs so the grid, geometry, and merge topology remain coherent. Before insertion or deletion:

- inspect row/column count and merge regions;
- identify the source row/column for formatting, if any;
- decide where formulas or semantic totals belong; and
- retain numeric alignment, number formats, and header/body hierarchy.

Afterward, verify the rectangular grid, merge map, geometry, cell text, and formatting of neighboring cells. Rendering is appropriate when added content can change wrapping or row height.

## Charts and workbooks

A chart is not only the visible frame. It can include a chart part, caches and formulas, an embedded workbook, style/color parts, and relationships. Ownership matters.

Use `replace_data_safe()` for supported category charts. Never update only the displayed cache or only the workbook. Refuse shared editable ownership and unsupported graph types before mutation.

After a supported update, verify:

- chart type and formatting remained stable;
- categories and every series have the expected values and lengths;
- formulas, caches, and workbook content agree;
- a workbookless chart remains workbookless; and
- neighboring or similarly named charts are unchanged.

Do not rebuild a chart as an image unless the user requests flattening and accepts lost editability.

## Images and shared media

Use `replace_image()` to preserve crop, rotation, mask, and geometry. Determine whether multiple pictures share one image part before low-level replacement. A local replacement should isolate the target when required, not silently change every consumer.

Preserve aspect ratio unless deliberate cropping is part of the design. Do not stretch a logo, photograph, or screenshot to fill a box.

## Slide lifecycle and composition

Use Paper lifecycle APIs rather than copying or deleting XML parts directly. Slide operations can affect:

- presentation order and permanent IDs;
- sections and custom shows;
- layouts, masters, and themes;
- notes and comments;
- charts, workbooks, media, and external links; and
- internal slide hyperlinks.

For cloning, retain the operation's policy and verify independence of charts/workbooks and the intended sharing of media. For deletion, verify that inbound references and reachability remain valid.

For cross-deck imports, compare slide dimensions and choose composition mode based on intent:

- preserve destination design with `adopt_theme`;
- preserve source appearance with `keep_appearance`; or
- localize supported appearance with `bake` after accepting its limits.

Review every returned import or rebind report. A completed call can still report format shifts, reused parts, or deliberate drops that require acceptance.

## Custody surfaces

Inventory features that may need preservation even when Paper cannot author them:

- macros and digital signatures;
- OLE, ActiveX, controls, and embedded objects;
- animations, transitions, audio, and video;
- SmartArt and alternate content;
- embedded fonts;
- comments and external links; and
- unknown extensions or vendor-specific parts.

Do not mutate or publish a digitally signed presentation as though the signature remained valid. Do not use another office application to resave the deliverable unless conversion is explicitly requested.

## Exceptional OOXML

Treat low-level OOXML as a narrow engineering operation, not the default escape hatch.

Before editing:

1. Work on a candidate copy or in-memory candidate, never the source.
2. Record source hashes and the expected member budget.
3. Identify the owning part, relationship file, content type, namespace, and schema-defined child position.
4. Enumerate inbound and outbound relationships and decide what is shared versus independently owned.
5. Define the exact rollback/refusal condition before the first write.

After editing:

1. Parse every changed XML member.
2. Check internal relationship targets and content types.
3. Fresh-reopen with Paper.
4. Compare expected and actual package members.
5. Verify semantic objects and unaffected consumers.
6. Render every visually affected slide.
7. Use Microsoft PowerPoint without accepting Repair when compatibility risk is material.

Never continue a refused mutation through raw XML on the same live graph. Obtain a clean export, narrow the request, flatten an explicitly approved object, report the boundary, or begin a separately scoped package-repair candidate with authoritative invariants and full low-level validation.
