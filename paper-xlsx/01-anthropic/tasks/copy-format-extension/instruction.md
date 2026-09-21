# Extend the input format

In `eval_fixtures/xlsx/format-extension.xlsx`, make `Inputs!B3:B5` use the same
formatting as `B2` for publication. Keep their contents and the sheet layout.
Our reporting system reads the workbook without opening Excel, so keep the
existing calculated results available as well as the formulas.

Save to `evals/copy-format-extension/output.xlsx`.
