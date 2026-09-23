# Paper Office Evals

61 Office evals: 27 DOCX, 15 PPTX, 19 XLSX.

## Treatments

Each format has the same tasks in three runnable Harbor datasets:

| Directory | Package | Skill |
|---|---|---|
| `01-anthropic` | Upstream | From [anthropics/skills](https://github.com/anthropics/skills) |
| `02-paper-instruments` | Paper | Included in `skill/<format>` |
| `03-no-skill` | Upstream | None |

## Run

Install [Harbor](https://harborframework.com/) and Docker. Configure your provider credentials and set `MODEL` to a provider-qualified model ID. From this directory:

```sh
harbor run --path paper-pptx/02-paper-instruments/tasks \
  --skill paper-pptx/02-paper-instruments/skill/pptx \
  --agent opencode --model "$MODEL" --env docker
```

Replace `pptx` with `docx` or `xlsx` for another format.

- **Anthropic:** use `01-anthropic/tasks` and point `--skill` at the matching skill in your Anthropic checkout.
- **No skill:** use `03-no-skill/tasks` and omit `--skill`.

Keep model and run settings the same across treatments. Find results in `jobs/`.

[MIT license](LICENSE) · [Third-party notices](THIRD_PARTY_NOTICES.md)
