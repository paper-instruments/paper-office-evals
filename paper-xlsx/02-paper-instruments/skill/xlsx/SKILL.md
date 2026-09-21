---
name: xlsx
description: Inspect, create, edit, and verify Excel workbooks with the installed paper-xlsx 0.2.1 distribution and its openpyxl-compatible API. Use for preservation-aware work on .xlsx and .xlsm files, especially when formulas, drawings, validations, links, VBA, pivots, or other package content must survive an edit.
---

# Excel workbooks with paper-xlsx 0.2.1

The `paper-xlsx` distribution is already installed and imports as `openpyxl`. Ordinary openpyxl APIs remain the primary editing surface. Do not install, uninstall, upgrade, replace, or bypass either spreadsheet distribution. When provenance matters, confirm `openpyxl.__paper_version__ == "0.2.1"`; the compatible openpyxl import version is 3.1.5.

## Choose the workbook mode

Load an editable OOXML workbook normally to use preserve mode:

```python
from openpyxl import load_workbook

wb = load_workbook(source_path)
```

Preserve mode splices modeled changes into the source package while retaining unrelated content. Use `Workbook()` for a genuinely new workbook. Preserve the source extension and workbook type unless conversion is requested. Use `preserve=False` only when the task intentionally accepts stock openpyxl regeneration; never use it to evade a `PaperRefusal`.

A preserve-mode workbook loaded with `data_only=True` contains cached results in place of formulas. Saving it refuses unless formula loss is explicitly allowed, and even then only edited formula cells lose their formulas. Prefer a normal formula-preserving load for edits.

## Route by risk

- Use bounded ranges, sheet metadata, defined names, and ordinary collections for direct inspection. Use `wb.search()` when the target is not already bounded and `ws.allowed_values()` when list validation matters.
- Use standard openpyxl cell, formula, style, table, chart, image, name, validation, and protection APIs for ordinary work. Use `copy_format()` to extend a local style pattern without inventing a new one.
- Use guarded row, column, range, and sheet operations when addresses move. Retain the returned `AddressRemap` when later checks or dependent references need the before-to-after mapping.
- Use `chart.repoint()` for an existing chart series, `ws.append_table_row()` for atomic table expansion, and `ws.replace_image()` for a targeted image replacement.
- Use `wb.set_pivot_refresh_on_load()` only when Excel should refresh existing pivot caches on open.
- Use `wb.validate()`, a save receipt, `scan_errors()`, or `diff_workbooks()` only when each answers a concrete risk. These are evidence sources, not substitutes for checking the requested result.
- Use the optional oracle only when current calculated values are required and LibreOffice is available. It measures a candidate; it does not mutate the workbook or prove Excel or business fidelity.

## Preserve formula custody

Value or formula edits invalidate affected retained formula caches and request recalculation so stale results are not shipped as current. Style-only edits preserve caches. A later `data_only=True` reopen may therefore return `None` for an invalidated formula until Excel or an oracle recalculates it; absence of a cache is not itself a calculation failure.

## Handle refusals and protection

A public operation that raises `PaperRefusal` is atomic: the refused mutation is not applied. Follow the structured remedy, narrow the operation, or report the blocker. Do not switch to stock mode, another package, or raw OOXML to force the same unsafe structural change.

Writes to locked cells on protected worksheets are advisory by default and can emit `ProtectedWriteWarning`. Set `wb.strict_protection = True` before editing when violating protection must refuse. Worksheet protection is a workbook guardrail, not access-control security.

## Finish with proportionate evidence

Save to the required output path and reopen the candidate. Assert the requested values, formulas, object state, and relevant dependencies. A normal save is valid; request a version 2 `EditReceipt` with `wb.save(output_path, receipt=True)` when changed-part evidence is useful. A receipt describes observed changes since load, not user intent. Validation checks the preserve save plan, not task correctness. A same-library reopen checks serialization, not Microsoft Excel rendering.

For a substantial new workbook or redesign, also read [Professional workbook quality](references/professional-workbook-quality.md). Existing workbook conventions take precedence over generic styling.

## Focused references

- [Perception and ordinary editing](references/perception-and-editing.md) covers bounded discovery, formulas, styles, and cache-aware edits.
- [Structural safety and refusals](references/structural-safety.md) covers address remaps, atomic helpers, relationships, and remedies.
- [Computation and delivery](references/computation-and-delivery.md) covers measurement-only calculation, diagnostics, receipts, and pivot delivery.
- [Professional workbook quality](references/professional-workbook-quality.md) covers workbook architecture, presentation, and review.
