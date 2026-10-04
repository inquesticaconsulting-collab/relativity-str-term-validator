# relativity-str-term-validator

[![tests](https://github.com/inquesticaconsulting-collab/relativity-str-term-validator/actions/workflows/tests.yml/badge.svg)](https://github.com/inquesticaconsulting-collab/relativity-str-term-validator/actions/workflows/tests.yml) [![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE) ![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)

Validate, auto-repair and self-verify search terms **before** they go into a Relativity **Search Terms Report (STR)**.

Every line you enter in an STR runs as its own dtSearch query, and dtSearch rarely rejects a broken query — it runs it and quietly counts the wrong documents. An unclosed quote, a comma-separated list on one line, `w/7/`, `custodian:Smith` or a stray `~~` won't raise an error; they just produce wrong hit counts. This tool catches those problems, fixes what can be fixed safely, explains every change, and re-checks its own output so the revised terms are guaranteed to re-validate.

## What it does

- **40+ checks** against dtSearch syntax and STR-specific limits (450-character term limit, no leading wildcards, no fuzzy `%`, duplicates ignored).
- **Auto-repair to a fixpoint** — the revised term is re-validated until it stops changing, so running the tool again on its own output never surfaces new errors.
- **Self-verification** — every revised term gets a *Revised Term Check*: `CLEAN`, `ADVISORY` (only warnings remain; the term runs) or `NEEDS REVIEW` (a human must decide).
- **Metadata extraction** — `custodian:`, `date:`, `ext:`, `subject:`, `folder:`, `project:` are not dtSearch syntax. They're pulled out to an *Extracted Conditions* sheet (for the searchable-set saved search), and their values are checked: impossible dates (`2024/31/02`), misspelled extensions (`xlxs` → `xlsx`), misspelled field names (`auther` → `author`).
- **Labelled assumptions** — when a default has to be chosen (e.g. `w/` with no number → `w/5`), the finding says `ASSUMPTION APPLIED` so the case team confirms it.
- **Optional spelling suggestions** (never auto-applied — deliberate misspellings are common in eDiscovery).

## Quick start

```bash
git clone https://github.com/inquesticaconsulting-collab/relativity-str-term-validator.git
cd relativity-str-term-validator
pip install -r requirements.txt          # openpyxl (+ optional pyspellchecker)

# try it on the bundled examples
python str_term_validator.py examples/sample_terms.txt

# your own list
python str_term_validator.py terms.xlsx -o terms_validation.xlsx --clean terms_clean.txt
```

Input is `.xlsx`, `.csv` or `.txt`, one term per line/row. In spreadsheets a column headed `Term`, `Search Term`, `Keyword` or `Revised Term` is found automatically — so a validation report can be fed straight back in to re-check edited terms.

| Option | Meaning |
|---|---|
| `-o, --output` | Excel report path (default `<input>_validation.xlsx`) |
| `--clean PATH` | Also write an STR-ready `.txt` (NEEDS REVIEW terms excluded, duplicates removed) |
| `--no-spellcheck` | Skip spelling suggestions |
| `--version` | Print version |

## Example

Input (one STR line):

```
(("data breech" AND w/10 "notification" AND NOT stemming()) OR ("cyber security" w/5 "protocol failure") AND custodian:Garcia AND date:2024-14-01 TO 2024-02-40 AND ext:emll AND NOT ("privilege" OR "attorney client") AND fileext:.doc** AND "transaction"~~~ AND "wire transfer" w/4 "approval"))
```

Revised term (`ADVISORY: MANY_PROX_OPS, NESTED_PROX`):

```
(("data breech" W/10 "notification") OR ("cyber security" W/5 "protocol failure") AND NOT ("privilege" OR "attorney client") AND "transaction" AND "wire transfer" W/4 "approval")
```

plus four extracted conditions for the saved search (`custodian:Garcia`, the `date:` range, `ext:emll`, `fileext:.doc**`), with notes that `2024-02-40` is not a real date and `emll` is probably `eml`. Note the misspelling `breech` is kept as written — it's a real English word, and the tool never "corrects" terms on its own.

## Output

| Sheet | Contents |
|---|---|
| **Summary** | Status counts, self-check counts, most frequent issues |
| **Validation Detail** | Original term, status, issue codes, findings with reasoning, revised term, Revised Term Check |
| **Extracted Conditions** | Every `field:value` condition, with validation notes |

Severity: **ERROR** must be resolved before loading · **WARNING** needs a recorded decision · **INFO** is a note. The full catalog is in [docs/RULES.md](docs/RULES.md).

## Repository layout

```
str_term_validator.py      the validator (single file, CLI + importable)
requirements.txt           runtime dependencies
requirements-dev.txt       + pytest
examples/                  sample term lists to try the tool on
tests/                     pytest regression suite
docs/                      SOP, overview slide, team-review matrix, rule catalog
docs/build/                scripts that generate the documents in docs/
```

## Documentation

- [docs/SOP_STR_Term_Validation.docx](docs/SOP_STR_Term_Validation.docx) — standard operating procedure
- [docs/str_term_validator_overview.pptx](docs/str_term_validator_overview.pptx) — one-slide overview
- [docs/STR_Validation_Rules_Team_Review.xlsx](docs/STR_Validation_Rules_Team_Review.xlsx) — use-case matrix for team review (decision/owner/status columns)
- [docs/RULES.md](docs/RULES.md) — rule catalog

The docs are generated by the scripts in `docs/build/` (Python + Node: `docx`, `pptxgenjs`).

## Tests

```bash
pip install -r requirements-dev.txt
pytest -q
```

The same suite runs automatically on GitHub for Python 3.9–3.13 on every push (see the badge above).

The suite checks each specific repair and the core guarantee: every revised term is a fixpoint, and — unless flagged NEEDS REVIEW — a second validation run reports no errors.

## Limitations — please read

- **Syntax, not substance.** The tool guarantees revised terms are valid, runnable dtSearch. It cannot know whether they capture what the lawyers *meant*; every meaning-changing revision and every `ASSUMPTION APPLIED` finding needs case-team sign-off.
- **Index-specific behavior.** Noise words and the alphabet file (how `&`, `_`, `-`, `@` are treated) are configurable per dtSearch index. The tool assumes Relativity's defaults; adjust `NOISE_WORDS` and the special-character guidance if your index is customized.
- **Offline.** It does not connect to Relativity; verify hit counts in the workspace.
- Rules are based on the RelativityOne help pages *Using dtSearch syntax options*, *Create and edit search term reports* and *Search terms reports*.

Not affiliated with or endorsed by Relativity ODA LLC. Relativity is a trademark of its owner. Nothing here is legal advice.

## License

[MIT](LICENSE) © 2026 Inquestica Consulting — Inhaber: Neeraj Khandelwal
