# Inventory the account template's typography and colors

Before the design team updates `eval_fixtures/pptx/account-template.pptx`, inventory the typography and text colors used by every paragraph in the `Priorities` and `Template annotation` objects on the `Account priorities` slide. Include formatting inherited from the template and the colors actually selected by the slide's theme mapping. Also inventory the fill and outline colors of `Rollout banner`, `Reference panel`, and `Review highlight` on `Brand color review`.

Save `evals/template-format-audit/report.json` with two arrays:

- `paragraphs`: each entry contains `slide`, `object`, `paragraph` (one-based), `font`, `size_pt`, `color_rgb`, and `bullet`. The bullet object contains `kind` (`none`, `character`, or `numbered`), `marker` (the character or first rendered number label, including punctuation), `font`, and `size_pt`. For an unbulleted paragraph use `none` and null for its other bullet fields. Use actual point sizes for proportional bullets.
- `shapes`: each entry contains `slide`, `object`, `fill_rgb`, and `line_rgb`.

Report colors as six-digit RGB values. Use `none` for an absent fill or outline, and null where there is no single fixed RGB value, such as a gradient or a system-dependent color. Use null for any font the file does not determine. Do not guess a replacement color or font.

Leave the source presentation unchanged.
