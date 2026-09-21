# Third-party notices

The root MIT license applies to the original evaluation tasks, Paper skills,
and repository documentation unless a file states otherwise. It does not
replace the terms of third-party packages or materials.

## Anthropic document skills

The `01-anthropic` treatment uses the DOCX, PPTX, and XLSX skills from
[`anthropics/skills`](https://github.com/anthropics/skills). Anthropic describes
these document skills as source-available rather than open source. Their
included license restricts reproduction and redistribution, so this repository
does not vendor them. Users must obtain the skills from Anthropic and comply
with Anthropic's applicable terms.

## Upstream Office packages

- `python-docx` 1.2.0 is maintained by the python-openxml project and is
  distributed under the MIT License.
- `python-pptx` 1.0.2 is maintained by the python-openxml project and is
  distributed under the MIT License.
- `openpyxl` 3.1.5 is maintained by the openpyxl project and is distributed
  under the MIT License.

The corresponding upstream wheels and license metadata are retained inside the
treatment Docker build contexts.

## Paper packages

`paper-docx`, `paper-pptx`, and `paper-xlsx` are Paper Instruments forks of the
upstream packages above. Their treatment wheels retain the applicable package
metadata and upstream notices.

## Harbor

The tasks use the Harbor task format. Harbor is maintained by the Harbor
Framework Team and distributed under the Apache License 2.0. Harbor itself is
not vendored in this repository.
