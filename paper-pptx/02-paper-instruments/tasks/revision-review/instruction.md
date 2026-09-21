# Review presentation revisions

Review the before/after presentation pairs listed in `eval_fixtures/pptx/reviews.json` before they go back to the account teams. Report changes to slide order, slide text, tables, speaker notes, and list formatting. Ignore file serialization, object renaming, run splitting, and stacking changes that do not change the visible content. Leave every input unchanged.

Save `evals/revision-review/report.json` for the review tracker. Use a `reviews` array with one entry per supplied pair, each containing its `review` identifier and a `changes` array. Include an empty array when there are no substantive changes.

Each change has `kind`, `slide` (the slide title), `shape_id` (the object's numeric PowerPoint ID), `before`, and `after`. Use these value formats:

- `slide_order`: ordered arrays of slide titles; `slide` and `shape_id` are null.
- `text`: arrays of paragraph text for the affected object, preserving whitespace.
- `table`: arrays of rows, each containing its cell texts.
- `notes`: the complete speaker-notes text; `shape_id` is null.
- `list`: an array, in paragraph order, of `{kind, marker, font, size_pt}`. Kind is `none`, `character`, or `numbered`; marker is the character or first rendered number label, including punctuation. For no list marker, use null for marker, font, and size. Report list changes separately only where the object's paragraph text is unchanged.

Report what changed without guessing why it changed. Do not include unchanged objects.
