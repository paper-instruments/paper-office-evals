# Professional workbook quality

Use this reference for new workbooks, substantial model construction, redesign, and deliverables another professional must maintain. Let the decision, workflow, source workbook, and user instructions determine the design. Do not force a finance template, fixed color convention, or dashboard onto every workbook.

## Begin with the work

Identify the reader, recurring workflow, source data, expected outputs, and update cadence before choosing sheets or styling. For an existing workbook, infer its local grammar first: sheet roles, repeated row patterns, defined names, input styles, units, dates, signs, and chart language. Preserve that grammar unless redesign is requested.

## Make the workbook navigable

Use short stable sheet names and workflow order. Make the first relevant area communicate purpose, period, units, and update status when those are not obvious. Use freeze panes, filters, tables, and defined names only when they help repeated review. Avoid hidden dependencies that a reviewer cannot trace.

## Separate inputs, calculations, and outputs

Make user-controlled assumptions, imported observations, calculations, and decision outputs distinguishable through labels, position, styles, protection, or names. Do not hardcode calculated results to make a current output look right. Keep consequential assumptions close to their use and record source, period, unit, and retrieval context when available.

## Build formulas for review and reuse

Prove a small representative slice before extending a formula pattern. Verify row, period, unit, sign, precedents, and expected result; investigate isolated exceptions. Guard plausible empty and zero-denominator cases deliberately, without hiding genuine data problems behind broad error suppression. A clean recalculation does not prove the business logic.

## Format for interpretation

Use restrained hierarchy and number formats that communicate meaning: real dates, percentages stored as fractions, currencies with clear scale, and precision consistent with the source. Use alignment, whitespace, modest fills, and borders to make assumptions, exceptions, totals, and outputs easy to find. Avoid merged cells as a layout system and never rely on color alone.

## Use tables, charts, and templates deliberately

Use native tables when sorting, filtering, expansion, or repeated entry warrants them. Verify headers, totals, references, and validation after structural edits. Choose charts for the analytical question and preserve their categories, series, units, ordering, and source ranges.

A reusable template should identify editable cells, required fields, units, valid choices, and update order. Treat validation and protection as guardrails, not security. Do not add example rows, legends, new input colors, or instruction sheets to an existing workbook unless requested.

## Review the finished workbook

After saving and reopening:

1. Confirm sheets, dimensions, and requested content.
2. Recheck representative formulas, patterns, units, signs, and sources.
3. Inspect relevant error tokens, cache state, and calculation exclusions.
4. Compare source and candidate when preservation or structural change matters.
5. Render in Microsoft Excel when charts, images, print layout, pivots, VBA, or visual fidelity matter.

Report only evidence actually obtained. Validation, receipts, semantic differences, oracle measurements, and visual rendering answer different questions; none independently proves the user's intended result.
