# Computation and delivery

paper-xlsx does not calculate formulas inside the workbook model. It preserves honest cache state and offers an optional, profile-isolated LibreOffice oracle for measurement when current calculated values are required.

## Formula custody

Preserve saves invalidate affected caches after value or formula edits and request recalculation. Style-only edits preserve caches. A `data_only=True` view may show `None` after invalidation; do not mistake an absent cache for an evaluation error.

A preserve-mode workbook loaded with `data_only=True` refuses a save unless `allow_formula_loss=True` is explicit. When explicitly allowed, only cells edited through that cached-value view lose their formulas; untouched formula cells retain their original bytes. Prefer a formula-preserving load.

## Measurement-only oracle

The supported functions operate on paths rather than a live workbook:

- `openpyxl.oracle.recalc(source, output_path=...)` recalculates a temporary candidate and reports recognized formula and error cells.
- `openpyxl.oracle.certify(source)` compares recalculated results with existing caches and records exclusions. A source without usable caches can be baseline-unverifiable rather than failed.
- `openpyxl.oracle.evaluate(source, set=..., read=...)` measures specified outputs under one temporary set of inputs.
- `openpyxl.oracle.evaluate_many(source, cases, read, ...)` batches those measurements through isolated profiles.

These operations do not write results back into the source workbook. Save and reopen the candidate before measuring unsaved changes. Run them on copies and retain the paper-xlsx-produced deliverable rather than replacing it with a generic LibreOffice round trip.

Oracle status is calculation evidence only. It does not prove business logic, Microsoft Excel fidelity, chart appearance, pivot behavior, VBA behavior, external-link availability, or user intent. Review exclusions as unmeasured scope, not as passes or failures. Do not rewrite a correct Excel formula solely to satisfy an optional LibreOffice measurement.

## Diagnostics and save evidence

Use checks in proportion to the risk:

- `wb.validate()` exercises the current preserve-save plan without writing.
- `wb.save(path, receipt=True)` writes once and returns a version 2 cumulative `EditReceipt` for changes since load.
- `openpyxl.preserve.scan_errors(wb)` reports visible error tokens without recalculation.
- `openpyxl.preserve.diff_workbooks(before, after, remaps=())` classifies semantic workbook differences.

Validation proves only the plan it checked. A receipt proves what the library recorded, not that the task was completed correctly. A same-library reopen proves serialization consistency, not Excel compatibility.

When an edit affects a recognized local pivot source or its dependencies, validation and save refuse unless the affected pivot is explicitly selected for refresh. `wb.set_pivot_refresh_on_load()` accepts that its saved cache can remain stale until Excel refreshes it; the receipt reports that requirement. This flag is not recalculation evidence, and a headless reader can still observe the old cached result before Excel opens the file. Existing VBA and pivots are custody surfaces; use native authoring tools when asked to create new VBA or pivot structures. Render with Microsoft Excel when visual fidelity is material.
