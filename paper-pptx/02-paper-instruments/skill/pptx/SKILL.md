---
name: pptx
description: Create, edit, inspect, compose, repair, and validate professional PowerPoint `.pptx` presentations with Paper PPTX. Use for slide creation, template-based deck work, text and formatting edits, tables, charts, images, notes, footers, layouts, cross-deck imports, package-preserving changes, presentation audits, and PowerPoint output QA.
---

# PowerPoint with Paper PPTX

Produce a presentation that is useful to its audience, editable in PowerPoint, faithful to its source material, and no more changed than the task requires. Treat visual quality, content quality, and package integrity as separate responsibilities.

Paper PPTX is installed as the `paper-pptx` distribution but imports through `pptx`. Its public surface includes upstream `python-pptx` APIs plus guarded inspection, targeting, composition, and package operations.

## Start from the outcome

Before editing, establish the facts that affect the work:

- audience, decision, and presentation setting;
- supplied content, factual constraints, and source requirements;
- source deck or template, including its visual grammar and slide size;
- whether the output must remain natively editable;
- requested fidelity, compatibility, privacy, and delivery evidence; and
- exact output path.

Infer these from the request, source deck, template, and surrounding materials when possible. Ask only when a missing choice would materially change the scope, design, or compatibility strategy; otherwise state a reasonable assumption and proceed.

Do not invent a new visual system when a usable template exists. Do not preserve a weak layout merely because it exists when the user asked for a redesign. Resolve that distinction from the request.

Preserve supplied inputs by default. For brownfield work, write a distinct candidate; for greenfield work, write the requested output. If the user explicitly authorizes in-place replacement, use a backup or rollback strategy proportionate to the risk and do not claim source immutability.

## Select the work mode

Choose one primary mode, then load only its relevant references.

- **Create:** Build a new presentation, preferably from a supplied template. Read [Professional presentation quality](references/professional-quality.md) and the creation examples in [Examples](references/examples.md).
- **Edit:** Change content or formatting in an existing deck while preserving unrelated behavior. Read [Editing and structural safety](references/editing-and-structure.md).
- **Compose or restructure:** Clone, move, import, rebind, or delete slides and owned objects. Read [Editing and structural safety](references/editing-and-structure.md) and [Paper API routing](references/paper-api-routing.md).
- **Inspect or report:** Open read-only, inspect only the needed surfaces, and do not save. Read [Paper API routing](references/paper-api-routing.md).
- **Validate or compare:** Apply the risk-based ladder in [Validation and delivery](references/validation-and-delivery.md).

For an uncertain runtime, check provenance once before work:

```bash
paper-pptx-doctor
python -c "import pptx; print(pptx.__file__, getattr(pptx, '__paper_version__', None), pptx.__version__)"
```

Do not repeat this check for every file in the same verified environment. Do not install, remove, or replace presentation packages unless the user requests environment changes.

## Inspect proportionally

Inspect enough to choose a safe operation, not everything by reflex.

1. Open with `Presentation(path)`.
2. For an unfamiliar brownfield deck, use `inspect_deck(prs)` or `scripts/inspect_pptx.py` to identify relevant slides, layouts, names, geometry, notes, charts, tables, and custody surfaces.
3. Inspect text or effective formatting only on the relevant slides and objects.
4. Retain the full baseline values used to choose a target.
5. Require a unique target. Prefer semantic shape names or a Paper `BlockAnchor`; never silently choose the first duplicate.

Opening and inspecting are not permission to save. Effective values can be unresolved by design. Preserve and report unresolved evidence instead of substituting a plausible font, color, size, or layout.

## Choose the narrowest supported operation

Prefer a public API that directly expresses the requested intent:

- ordinary upstream-compatible APIs for supported creation and local formatting;
- Paper named-object and anchored-text APIs for precise targeting;
- Paper lifecycle, chart, image, table, footer, layout, and composition APIs when ownership or relationships matter; and
- Paper inspection/diff APIs when effective values or package evidence matter.

Read [Paper API routing](references/paper-api-routing.md) for exact operations and refusal boundaries. Before calling a Paper-specific operation whose signature, transaction scope, or result shape you do not know, consult its contract rather than inferring it from the method name.

Treat `PaperRefusal` as a document-safety result. Report the operation, exception class, relevant structure, and smallest safe alternative. Do not treat a refusal as automatic permission to continue the same mutation through private modules or raw XML. An explicitly requested package-repair workflow is a separate operation with its own baseline, invariants, and validation.

Read-only ZIP/XML inspection is acceptable when it provides necessary evidence and does not alter the package. Mutate raw OOXML only when all of the following are true:

- no public API expresses the requested change;
- low-level mutation is genuinely necessary to the requested outcome, not merely more familiar or convenient;
- the user has explicitly accepted the low-level path when it creates material preservation or compatibility risk;
- the exact schema position, relationship closure, content types, and ownership effects are understood from authoritative evidence and this package; and
- the candidate can be tested separately with a declared change budget.

Keep low-level work narrowly scoped. Do not copy a slide, chart, image, notes part, or shape XML without its complete relationship and ownership graph.

## Build for the presentation, not the library

For creation or redesign, make the content hierarchy clear before polishing shapes:

- Give each slide one purpose.
- Make the title useful: state the topic or takeaway appropriate to the context.
- Select a chart, table, diagram, image, or prose because it communicates the evidence best.
- Keep related facts together and repeated peers geometrically consistent.
- Preserve exact numbers, units, dates, source names, and qualifications.
- Use native charts, tables, text, and shapes when editability matters.
- Use image assets when they add evidence or meaning, not to satisfy a visual quota.
- Preserve the template's type, color, spacing, footer, and layout system unless redesign is requested.

Do not impose a consulting style, a fixed font, rounded or square cards, a palette, or takeaway titles on every deck. These are contextual choices. Read [Professional presentation quality](references/professional-quality.md) for adaptable standards and examples.

Name important content objects semantically when creating them. Names should describe roles such as `title`, `chart_revenue`, `table_scenarios`, or `source_market_data`, not implementation order.

## Save according to the task

Use ordinary `prs.save(output)` when normalized serialization is acceptable. Use `patch_save(source, prs, output)` only when preserving semantically unchanged original XML bytes is material to the request. Always inspect its returned diff; do not assume a no-op or minimal package delta.

For bulk operations, prefer a supported operation that changes the logical unit in one call. Current Paper setters can carry per-operation transaction cost; avoid needless setter loops and repeated open/save cycles. Do not bypass public APIs solely for speed.

## Validate in proportion to risk

Every produced presentation needs:

1. the requested nonempty output;
2. a fresh reopen from the saved file; and
3. targeted assertions proving the requested result.

For a separate brownfield candidate, verify input immutability when it is required by the task, custody expectations, or the risk of the operation. Do not require a source hash for greenfield work or use an unchanged-source claim for authorized in-place work.

Add validation based on what could have broken:

- **Content-only local edit:** verify changed text/formatting and nearby layout; render affected slides when wrapping or geometry could change.
- **New or redesigned deck:** render every slide and review the deck as a sequence, then inspect risky slides at full size.
- **Chart, image, table, notes, or footer work:** verify the edited object's semantics and the features promised to remain unchanged.
- **Lifecycle, layout, composition, cleanup, or low-level work:** compare slide/object identity, relevant relationships, package members, source immutability, and operation reports.
- **Compatibility-sensitive work:** open a copy in the target Microsoft PowerPoint environment without accepting Repair.

Use the bundled helpers when useful:

```bash
python "$PPTX_SKILL_DIR/scripts/inspect_pptx.py" output.pptx
python "$PPTX_SKILL_DIR/scripts/compare_pptx.py" source.pptx output.pptx
python "$PPTX_SKILL_DIR/scripts/render_pptx.py" output.pptx --output-dir rendered
```

Set `PPTX_SKILL_DIR` to the directory containing this `SKILL.md`, or invoke each helper by its actual absolute path. Do not assume the task repository itself contains a `scripts/` or `tools/` copy.

The scripts report only their named checks. A successful Paper reopen does not prove PowerPoint compatibility. Rendering does not verify notes, metadata, formulas, or relationship ownership. Read [Validation and delivery](references/validation-and-delivery.md) before making broad fidelity claims.

## Deliver with calibrated evidence

Report concisely:

- output path and intended changes;
- whether inputs remained unchanged;
- validation actually performed and its result;
- rendering or PowerPoint application/version used, if any; and
- unresolved values, blind regions, deliberate flattening/drops, consumer-dependent behavior, or untested compatibility risks.

Do not claim “pixel perfect,” “fully preserved,” “PowerPoint compatible,” or “transactional” without naming the evidence that establishes that scope.

## Reference map

- [Paper API routing](references/paper-api-routing.md): capability selection plus exact contracts for high-risk public operations, outputs, and boundaries.
- [Editing and structural safety](references/editing-and-structure.md): brownfield targeting, ownership, notes, layouts, composition, and exceptional OOXML.
- [Professional presentation quality](references/professional-quality.md): narrative, templates, layout, typography, tables, charts, images, and financial/board conventions.
- [Validation and delivery](references/validation-and-delivery.md): risk tiers, package checks, rendering, PowerPoint QA, and evidence language.
- [Examples](references/examples.md): compact end-to-end patterns for common work modes.
