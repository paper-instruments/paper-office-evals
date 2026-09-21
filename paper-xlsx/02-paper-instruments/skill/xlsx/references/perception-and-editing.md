# Perception and ordinary editing

Choose the smallest inspection surface that can identify a unique target. Whole-workbook dumps waste context and can conceal repeated labels or scenario-specific regions.

## Targeted inspection

Start with `wb.sheetnames`, `ws.calculate_dimension()`, bounded cell iteration, `wb.defined_names`, and relevant worksheet collections. Use `wb.search()` only when the location is unknown; it returns matches from materialized cells and can include values and formulas. Confirm sheet, coordinate, nearby labels, units, and repeated occurrences before editing.

Use `ws.allowed_values(cell)` to inspect a list validation that covers a cell. It returns the allowed values, or `None` when no list validation covers that cell. Unsupported or ambiguous validation sources refuse instead of returning a partial list. This is useful evidence for selecting or preserving an input, not permission to change a validation rule.

Use `openpyxl.preserve.scan_errors(wb)` to find visible Excel error tokens without recalculating. It does not establish formula logic, cache freshness, or coverage of values that an external engine could not produce.

## Ordinary edits and formatting

Use normal openpyxl APIs for cells, formulas, styles, comments, validation, names, tables, charts, images, and worksheets. Preserve mode retains unrelated source-package content around those modeled changes.

To extend an established style pattern, use `openpyxl.preserve.copy_format(ws, source, destination)`. It applies the source cell's finite formatting to the destination range, expands normal range bounds, and refuses before mutation for merged-cell interiors or strict-protection violations. Inspect the nearby pattern first; format copying should preserve local grammar, not create a new visual system.

## Formula work

Check formula repairs at the task level:

1. Compare neighboring formulas, labels, units, periods, and signs.
2. Repair a few representative cells and verify their precedents and intended result.
3. Extend only the proven pattern and investigate exceptions.
4. Save and reopen the candidate.
5. Recalculate only when current values are part of the requested evidence.

Value or formula edits invalidate affected retained formula caches and request recalculation. Style-only edits preserve formula caches. Consequently, an invalidated formula may appear as `None` in a later `data_only=True` view until a calculation engine runs. That missing cache is honest state, not proof that the formula failed.

Never edit and save a formula workbook from a `data_only=True` load unless the user explicitly accepts formula loss. Preserve mode refuses such saves by default; `allow_formula_loss=True` is an explicit custody decision, not a routine workaround.

An error-free recalculation can still use the wrong row, sign, period, or assumption. Check task-specific invariants and representative outputs separately.
