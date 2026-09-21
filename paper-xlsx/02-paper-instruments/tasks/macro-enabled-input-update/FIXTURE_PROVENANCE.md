# Fixture provenance

## VBA project

The `xl/vbaProject.bin` embedded in this task's XLSM fixtures across all three
treatments is extracted unchanged from openpyxl's
`openpyxl/tests/data/reader/vba+comments.xlsm`, retained in the Paper fork at
commit `96cea952166d203ed11b0f0c67a3e8ad65d06157` (source Git blob
`e8bc9b844862ac63793c30ad2ded784b22829ac0`).

- Source workbook SHA-256: `982cdf0b67d8e4d96e8afd89385c6bc0b978c725ee70113a818372a33dc11935`.
- Extracted project SHA-256: `f17a5a2161e3589276425f08bdd450c9bc99c1c9574d1b5d3738fdc0fbf7186a`.
- License: upstream openpyxl MIT; see `OPENPYXL-LICENSE.txt`.

Static inspection with oletools 0.60.2 found three modules. `Module1` contains
only `Sub Button1_Click(): MsgBox "Hello World": End Sub`. `ThisWorkbook` and
`Sheet1` contain class attributes and no executable statements. There are no
auto-open handlers. No macro was executed during inspection. The fixtures
preserve the project's bytes and matching workbook/worksheet code names.
Prior container tests traversed FAT, MiniFAT and directory streams and validated
the project and module metadata;
this is not evidence of successful macro execution in Excel.

## Threaded review

The threaded review fixture is authored test data, not an Excel export. Its
people, messages, reply links, timestamps and resolved state are retained from
the original benchmark. Its `text` is string content, per
[Microsoft's threaded-comment schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/adb84732-9fc8-48b6-bddc-6b0bcdaad940).
Prior tests validated this supported schema subset, identifiers, graph links and OPC
relationships. They did not establish full-document XSD validation or native Excel
rendering. All people and email addresses are fictional benchmark data.
