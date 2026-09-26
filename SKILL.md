---
name: pptx-net-validate
description: "Use this skill to validate a .pptx/.docx/.xlsx file for OOXML/OPC correctness (why PowerPoint says a file is 'corrupt' or shows a repair dialog), get a root-cause hint for each error, and — for structural issues only — auto-fix the file. Uses a pure-Python engine (lxml + ISO-IEC 29500-4 XSD) by default; falls back to PptxValidatorNet8.exe with --exe. Trigger on: 'validate this pptx', 'why is this pptx corrupt', 'check OOXML/OPC compliance', 'fix this broken pptx', 'diagnose PowerPoint repair dialog'."
---

# PPTX/OOXML validation and structural auto-fix

Pure-Python OOXML validator — **no .NET dependency** for the default path.
Gap-tested against `PptxValidatorNet8.exe` on real SRDP files: **8/8 schema
finding counts match**. The `.NET` exe is still available as `--exe` fallback.

## Backends

| Backend | How to invoke | Requirement |
|---|---|---|
| **Pure Python** (default) | `python scripts/validate.py deck.pptx` | `lxml`, `defusedxml` |
| **.NET exe** (fallback) | `python scripts/validate.py deck.pptx --exe path/to/PptxValidatorNet8.exe` | .NET 8+ runtime |

## Workflow

1. **Diagnose**: `python scripts/validate.py deck.pptx`
   Prints per-finding: `partUri`, `xPath`/line number, severity, and a
   root-cause hint. Add `--json` for structured output suitable for scripting.

2. **Auto-fix (structural only)**: `python scripts/autofix.py deck.pptx`
   Removes dangling references — missing parts, broken rels, orphaned anim
   refs, broken ink annotations. Writes `deck.fixed.pptx` by default;
   `--in-place` overwrites (keeps `.bak`). Re-validates and reports what's left.

   **Not auto-fixed, by design:**
   - `contentTypeMismatches` — ambiguous which side is wrong.
   - `schemaFindings` — too varied. Use `hint`/`description`/`xPath` from
     `--json` to locate the node, hand-edit the part XML, re-zip, re-validate.

3. **When findings remain**: open the fixed file in PowerPoint to confirm no
   repair dialog. If it still repairs, Save-As and diff the zip entries.

## What the Python engine checks

### 1. ZIP / OPC structural checks
- **missingParts** — parts declared in `[Content_Types].xml` that don't exist in the ZIP
- **brokenRels** — `.rels` targets pointing to nonexistent parts
- **contentTypeMismatches** — declared content-type vs actual XML root element
- **orphanedAnimRefs** — `p:spTgt`/`p:bldP` `@spid` values with no matching shape in `p:spTree`
- **inkRelErrors** — `mc:AlternateContent` ink blocks with missing/broken `r:id` rels

### 2. XSD schema validation (lxml + ISO-IEC 29500-4)
- `ppt/**` → `pml.xsd`
- `word/**` → `wml.xsd`
- `ppt/charts/**` → `dml-chart.xsd` (semantic checks, not XSD)
- `ppt/diagrams/data*` → `dml-diagram.xsd`
- `ppt/diagrams/drawing*` → direct MinInclusive coord check (ms-proprietary root)
- `.rels` files → `opc-relationships.xsd`
- `[Content_Types].xml` → `opc-contentTypes.xsd`
- Skipped (no ISO XSD): `customXml/item*.xml`, `customXml/itemProps*.xml`

### 3. PPTX semantic checks
- Chart axis reference integrity (axis IDs, stacked-bar label positions)
- Master-theme sharing (causes PowerPoint to refuse the file)
- Slide layout ID references in slide masters
- Duplicate slide layout refs per slide
- Notes-slide ref uniqueness

### 4. Severity classification
| Level | Meaning |
|---|---|
| **Critical** | File won't open or PowerPoint shows repair dialog |
| **High** | Visible corruption: shapes/tables broken, text missing |
| **Medium** | Subtle rendering difference; file opens |
| **Low** | Office 2016+ extension attributes; ignored at runtime |

### 5. Root-cause hints
~100 pattern rules mapping XSD error messages to actionable fix descriptions.

## Known gaps vs .NET SDK (acceptable, documented)

| Situation | Behaviour |
|---|---|
| `ppt/diagrams/drawing*.xml` — complex SmartArt layout errors beyond coord checks | Not fully detected (ms-proprietary root, no ISO XSD) |
| Office 2016/2019/2021/365 version-specific schema targets | Python uses ISO-29500-4 (2016) only; `.NET` supports per-version XSD |
| `NodeXml` / `RelatedNodeXml` fields | Always `null` in Python output (lxml doesn't expose element XML per error) |

## Scripts

| Script | Purpose |
|---|---|
| `scripts/validate.py <file\|dir> [-r] [-a] [--json] [--exe path]` | Validate one or more files; auto-selects Python backend |
| `scripts/py_validate.py <file\|dir> [-r] [-a] [--json]` | Pure-Python engine standalone |
| `scripts/autofix.py <file.pptx> [-o out \| --in-place]` | Structural auto-fixer |

## Output JSON schema

Same as `PptxValidatorNet8.exe --format json`:

```json
[{
  "path":                  "deck.pptx",
  "fileType":              "pptx",
  "fileNotFound":          false,
  "sdkException":          null,
  "missingParts":          [],
  "brokenRels":            [],
  "contentTypeMismatches": [],
  "orphanedAnimRefs":      [],
  "inkRelErrors":          [],
  "schemaFindings": [{
    "partUri":       "/ppt/slides/slide1.xml",
    "xPath":         "line:42",
    "description":   "Element '...rPr', attribute 'sz': ...",
    "errorType":     "SchemaError",
    "severity":      "High",
    "hint":          "RunProperties 'sz' out of range ...",
    "nodeLocalName": null,
    "nodeXml":       null
  }],
  "semanticIssues": []
}]
```
