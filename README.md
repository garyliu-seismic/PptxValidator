# PptxValidator

Pure-Python OOXML validator for `.pptx` / `.docx` / `.xlsx` files — **no .NET
dependency required**. The `.NET` backend (`PptxValidatorNet8.exe`) is still
supported as an opt-in fallback via `--exe`.

Validated against real-world SRDP output files: **8/8 schema-finding counts
match the .NET SDK** after gap analysis and targeted fixes.

---

## Architecture

```
validate.py / autofix.py   ← unified entry-points (backend-agnostic)
      │
      ├── py_validate.py   ← default: pure-Python engine (lxml + ISO-29500 XSD)
      │       ├── ZIP/OPC structural checks  (5 check types)
      │       ├── lxml XMLSchema validation  (ISO-IEC 29500-4 XSD set)
      │       ├── PPTX semantic checks       (charts, themes, layouts …)
      │       ├── Severity classifier        (Critical / High / Medium / Low)
      │       └── Root-cause hint engine     (~100 rules)
      │
      └── PptxValidatorNet8.exe  ← opt-in .NET fallback (--exe flag)
```

## Validation layers (py_validate.py)

| Layer | What it checks |
|---|---|
| **ZIP / OPC structure** | Missing parts, broken `.rels` targets, content-type root-element mismatches, orphaned animation `spid` refs, broken ink-annotation `r:id` refs |
| **XSD schema** | Every XML part validated against ISO-IEC 29500-4 (2016) XSD via `lxml.XMLSchema`; charts via `dml-chart.xsd`; SmartArt via `dml-diagram.xsd`; `.rels` via OPC XSD |
| **PPTX semantics** | Chart axis references, stacked bar label positions, master-theme sharing, slide layout ID integrity, duplicate layout refs, notes-slide ref uniqueness |
| **Severity** | Critical / High / Medium / Low — matched to PowerPoint repair-dialog likelihood |
| **Hints** | Human-readable root-cause hints for every schema error pattern |

## Gap analysis vs .NET SDK (tested on SRDP files)

| Gap | Root cause | Fix applied |
|---|---|---|
| `customXml/item*.xml` false positives | App-specific XML, no OOXML XSD | Added to skip list |
| `ppt/diagrams/drawing*.xml` MinInclusive misses | MS-proprietary root element; `.NET` validates embedded `a:ext/@cx/cy` | `_validate_diagrams_drawing()` checks negative coords directly |
| `buSzPct` integer format false positive | ISO XSD requires `"80%"` pattern; Office stores `80000`; `.NET` accepts both | Added to ignored-error list |
| `UID` attribute severity mismatch (Medium→Low) | lxml says `"not allowed"`; .NET says `"not declared"` | Extended `get_severity()` to catch `"is not allowed"` → Low |

**Result: 8/8 files MATCH between Python and .NET schema finding counts.**

## Requirements

- Python 3.10+
- `lxml` (`pip install lxml`) — already installed if you use `python-pptx`
- `defusedxml` (`pip install defusedxml`) — for safe XML parsing in autofix

No .NET runtime needed for the default backend.

## Usage

```bash
# Validate (pure-Python, default)
python scripts/validate.py deck.pptx
python scripts/validate.py ./decks --recursive --all
python scripts/validate.py deck.pptx --json

# Validate with .NET backend (fallback, requires PptxValidatorNet8.exe)
python scripts/validate.py deck.pptx --exe path/to/PptxValidatorNet8.exe

# Auto-fix structural issues -> writes deck.fixed.pptx
python scripts/autofix.py deck.pptx

# Fix in place (keeps a .bak)
python scripts/autofix.py deck.pptx --in-place

# Direct Python engine
python scripts/py_validate.py deck.pptx --json
```

Or via `make` — see `make help`.

## What gets auto-fixed

`autofix.py` only removes dangling references (one unambiguous safe fix):

| Finding | Auto-fixed? |
|---|---|
| `missingParts` | yes — drop the `[Content_Types].xml` Override |
| `brokenRels` | yes — drop the dangling `<Relationship>` |
| `orphanedAnimRefs` | yes — drop the animation block |
| `inkRelErrors` | yes — drop the ink annotation block |
| `contentTypeMismatches` | no — ambiguous which side is wrong |
| `schemaFindings` | no — too varied; use `hint`/`xPath` to fix by hand |

## Repo layout

```
SKILL.md                   Claude Code skill definition
README.md                  this file
Makefile                   convenience targets (make help)
scripts/
  validate.py              entry-point: auto-selects Python or .NET backend
  py_validate.py           pure-Python validation engine
  autofix.py               structural auto-fixer (uses py_validate directly)
  helpers/
    __init__.py            OPC path utilities
    pptx_chart.py          chart axis / label-position semantic checks
    pptx_slide.py          fatal slide XSD error patterns
    pptx_theme.py          master-theme sharing detection
schemas/
  ISO-IEC29500-4_2016/     pml.xsd, wml.xsd, dml-main.xsd, dml-chart.xsd …
  ecma/fouth-edition/      OPC content-types, relationships, core-properties
  microsoft/               wml-2010/2012/2018 extension schemas
  mce/                     mc.xsd (markup compatibility)
```

## License

No license file included.
