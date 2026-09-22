# PptxValidator

A Python wrapper around Seismic's internal `PptxValidatorNet8.exe` (.NET
8/10) for diagnosing and
partially auto-fixing broken `.pptx`/`.docx`/`.xlsx` files.

It also ships as a Claude Code skill (`SKILL.md`) so Claude can validate and
fix OOXML files as part of a coding session.

## Why wrap instead of reimplement in Python

The underlying .NET tool layers custom checks — missing parts, broken
relationships, content-type/root-element mismatches, orphaned animation
references, broken ink annotations, a severity classifier, and a large
root-cause hint engine — on top of
`DocumentFormat.OpenXml.Validation.OpenXmlValidator`, Microsoft's own OOXML
schema validator. There is no Python equivalent of that validator, and
reimplementing full ECMA-376/ISO-29500 schema validation in pure Python would
be a large, error-prone undertaking with no guarantee of matching the SDK's
accuracy. Wrapping the existing, already-correct .NET binary is the right
tradeoff — see `SKILL.md` for the full rationale.

## Requirements

- A build of `PptxValidatorNet8.exe` (Release, net10.0). By default the
  scripts look for it at:
  ```
  Release\net10.0\PptxValidatorNet8.exe
  ```
  Pass `--exe <path>` to override.
- Python 3.10+, stdlib only (`zipfile`, `xml.etree.ElementTree`, `subprocess`,
  `json`) — no third-party packages to install.

## Usage

```bash
# Diagnose a single file
python scripts/validate.py deck.pptx

# Diagnose a whole folder (optionally recursive), including clean files
python scripts/validate.py ./decks --recursive --all

# Raw JSON, e.g. for scripting or feeding a schemaFinding's xPath/nodeXml to an editor
python scripts/validate.py deck.pptx --json

# Auto-fix structural issues only (missing parts, broken rels, orphaned
# animation refs, broken ink annotations) -> writes deck.fixed.pptx
python scripts/autofix.py deck.pptx

# Fix in place (keeps a .bak)
python scripts/autofix.py deck.pptx --in-place
```

Or via `make` — see `make help`.

### What gets auto-fixed, and what doesn't

`autofix.py` only touches dangling references at the ZIP/XML structural
layer, because those have one unambiguous safe fix (remove the dangling
reference):

| Finding | Auto-fixed? |
|---|---|
| `missingParts` | yes — drop the `[Content_Types].xml` Override |
| `brokenRels` | yes — drop the dangling `<Relationship>` |
| `orphanedAnimRefs` | yes — drop the animation block |
| `inkRelErrors` | yes — drop the ink annotation block |
| `contentTypeMismatches` | no — ambiguous which side is wrong; needs a human |
| `schemaFindings` (arbitrary OOXML schema errors) | no — too varied to rule-ify safely; use the reported `hint`/`xPath`/`nodeXml` to fix by hand, then re-validate |

## Repo layout

```
SKILL.md              Claude Code skill definition (workflow guidance for an agent)
README.md             this file
Makefile              convenience targets, see `make help`
scripts/validate.py   diagnose: run the .NET validator, print findings + root-cause hints
scripts/autofix.py    fix: diagnose then repair structural issues only
```

## License

Internal Seismic tooling; no license file included — do not distribute outside the org.
