# Structural safety and refusals

Use public paper-xlsx operations for edits that change addresses, sheets, relationships, drawings, charts, tables, or media. Raw ZIP/XML mutation bypasses dependency rewrites and preservation checks and is not a remedy for a refused public operation.

## Guarded address changes

Supported row and column insertion or deletion, range movement, and sheet lifecycle operations update recognized formulas, defined names, ranges, tables, chart references, and drawing anchors. Address-changing methods return an `openpyxl.preserve.AddressRemap` that maps pre-edit references into their post-edit locations. Keep the remap when verifying dependent references or passing it to `diff_workbooks()`.

If paper-xlsx cannot prove a structural rewrite safe, it raises a typed `PaperRefusal`. A public mutation that refuses is atomic: it leaves the operation unapplied. A save refusal leaves the destination untouched.

## Focused structural helpers

- `chart.repoint(series_index, new_range)` changes one existing series formula without rebuilding the chart. The new range must be valid for the series and workbook.
- `ws.append_table_row(table_name, values)` appends a row and expands the table plus supported dependent ranges as one guarded operation. It refuses before mutation when inputs or dependencies are unsafe.
- `ws.replace_image(target, replacement, name=None)` replaces one identified drawing image while preserving the surrounding drawing structure. Confirm the target is unique.
- `wb.set_pivot_refresh_on_load(pivots=..., all=False)` marks selected existing pivot caches for refresh. It does not create pivots or calculate their results.

## Respond to a refusal

Catch `openpyxl.errors.PaperRefusal` only when you can use its `kind`, `anchor`, and `options` to choose a documented safe remedy. Otherwise let the refusal surface with the blocker intact. Useful subclasses distinguish ambiguous or missing targets, unsupported structures, boundary violations, relationship policy, and unavailable or timed-out oracle operations.

Do not respond by disabling preserve mode, installing another spreadsheet library, or editing package XML. Narrow the target, use a supported public operation, preserve the workbook unchanged, or report what could not be completed.

Writes to protected cells emit `ProtectedWriteWarning` by default because Excel worksheet protection is advisory. Set `wb.strict_protection = True` before the write when the workflow requires an atomic refusal instead.

Use `openpyxl.preserve.diff_workbooks(before, after, remaps=(remap,))` when a semantic before/after comparison materially helps validate a structural change. A clean diff still does not prove that the chosen transformation matches user intent.
