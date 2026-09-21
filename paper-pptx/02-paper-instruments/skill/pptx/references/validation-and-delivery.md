# Validation and delivery

Use the least expensive validation that can detect the plausible failures of the operation. Add layers when risk increases; do not repeat every check by ritual.

## Contents

- Evidence layers and baseline checks
- Risk tiers
- Visual review
- Comparison and minimal churn
- Refusal and fault evidence
- Delivery language

## What each layer proves

| Layer | Establishes | Does not establish |
|---|---|---|
| Source/output hashes | File identity and source immutability | Semantic correctness |
| ZIP/XML checks | Readable package members and well-formed XML | OOXML schema validity or PowerPoint acceptance |
| Relationship checks | No detected dangling internal target | Correct ownership or consumer behavior |
| Fresh Paper reopen | Paper can parse the saved package | Visual fidelity or PowerPoint compatibility |
| Targeted semantic assertions | Requested objects contain expected state | Unchecked objects are unchanged |
| Paper package/deck diff | Recorded member or supported semantic changes | Visual equivalence |
| Rendering | Visible output in that renderer | Notes, metadata, formulas, animations, or PowerPoint behavior |
| Microsoft PowerPoint open | That tested PowerPoint environment accepts the file | Every other platform/version is identical |

State results at this granularity.

## Baseline for any produced deck

1. Write to the requested output. For brownfield work, default to a distinct candidate unless in-place replacement was explicitly authorized.
2. Confirm the output exists and is nonempty. Hash inputs when immutability or custody must be evidenced; do not require a source hash for greenfield or authorized in-place work.
3. Fresh-reopen the output from disk.
4. Reacquire targets from the reopened presentation.
5. Assert the requested semantic result and relevant preservation conditions.

Do not repeat Paper's guarded intake as a second generic validator. Use a fresh Paper reopen plus
assertions against the specific objects and invariants the operation could have changed. For
graph-level work, combine the operation report with `compare_pptx.py` and inspect the relevant
relationships or package members only when that evidence is material.

## Risk tiers

### Tier 1 — local content or style

Examples: one text replacement, one local format change, one existing-notes update.

Add:

- exact replacement or property assertions;
- unchanged slide count/order and target identity;
- nearby text/layout review; and
- rendering of affected slides when wrapping, autofit, font, crop, or geometry can change.

Do not require a full package manifest unless minimal churn or custody features are in scope.

### Tier 2 — object/data mutation

Examples: table expansion, image replacement, chart data update, footer application.

Add:

- object-specific semantic checks;
- isolation checks for shared or similarly named objects;
- geometry/crop/merge/chart-format preservation as relevant;
- formula/cache/workbook agreement for charts; and
- rendering of affected slides plus any slides sharing changed resources.

### Tier 3 — graph or inheritance mutation

Examples: clone/delete/reorder, layout rebind, cross-deck import, shape copy, low-level OOXML.

Add:

- slide IDs, order, sections, layouts, masters, notes presence, and shape identity checks;
- operation-report review;
- expected package-member budget and relationship closure;
- source and secondary-input immutability;
- refusal/rollback checks on disposable input when required; and
- rendering of every visually affected slide.

For unrelated source/destination decks, do not treat matching numeric slide IDs as lineage proof.

### Tier 4 — compatibility-sensitive delivery

Examples: macros or complex custody surfaces, prior Repair warnings, low-level package surgery, regulated or high-stakes delivery.

Add actual Microsoft PowerPoint testing in the target environment:

- open a copy without accepting Repair;
- record application, platform, and version;
- check warnings, field refresh, links, media, transitions, and feature-specific behavior; and
- ensure testing did not replace the intended candidate with a resaved or repaired copy.

If PowerPoint is unavailable, say so. Do not promote a library reopen into equivalent evidence.

## Visual review

For a new or redesigned deck, render every slide. Review a contact sheet for sequence, rhythm, and consistency, then inspect dense or visually risky slides at full resolution.

For a narrow brownfield edit, render affected slides and any slide that shares the changed layout, master, theme, or resource. Expand to all slides when the change can propagate globally.

Check:

- clipping, overflow, off-slide objects, and overlap;
- unexpected wrapping, font substitution, bullets, spacing, and autofit;
- alignment, gutters, hierarchy, and repeated-peer consistency;
- chart labels, axes, legends, data, and emphasis;
- table density, merges, numeric alignment, and headers;
- image crop, aspect ratio, resolution, and masks;
- placeholders, fields, footers, page numbers, and sources; and
- whether the deck reads coherently as a sequence.

LibreOffice or another renderer is useful evidence but not the target PowerPoint application. Never deliver an office-suite-resaved PPTX unless conversion is requested.

## Comparison and minimal churn

Use `scripts/compare_pptx.py` or Paper's `diff_package()` and `diff_decks()` when the task includes preservation, lineage, or change-budget requirements.

Interpret deltas, do not merely count them. A changed slide can legitimately require relationship, chart, workbook, notes, content-type, or media changes. Conversely, a small member count does not prove semantic correctness.

`patch_save()` can restore original bytes for XML it judges semantically unchanged. Always inspect the returned residual diff. Equivalent OPC relationships or content-type allocations from different producers can still be serialized differently.

## Refusal and fault evidence

For an operation expected to refuse atomically:

1. Hash or snapshot the candidate immediately before the attempt.
2. Capture exception class and message.
3. Verify the full relevant package or in-memory semantic state remains unchanged.
4. Verify related objects remain usable.
5. Do not save a mutated graph after rollback failure.

Syntax-checking a fault-injection script is not execution evidence. Run the probe in a disposable environment when the task requires demonstrated behavior.

## Delivery language

Prefer exact statements:

- “Fresh-reopened with Paper PPTX and passed targeted chart-data assertions.”
- “Rendered all 12 slides with LibreOffice; reviewed the contact sheet and slides 4, 7, and 9 at full size.”
- “Microsoft PowerPoint was unavailable, so open-without-Repair compatibility was not tested.”
- “Only the requested slide and its chart/workbook relationship members changed according to the package comparison.”

Avoid unbounded claims such as “fully preserved,” “pixel perfect,” “works everywhere,” or “safe.”
