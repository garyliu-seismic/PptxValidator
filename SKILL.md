---
name: pptx-net-validate
description: "Use this skill to validate a .pptx/.docx/.xlsx file for OOXML/OPC correctness (why PowerPoint says a file is 'corrupt' or shows a repair dialog), get a root-cause hint for each error, and — for structural issues only — auto-fix the file. Wraps the existing PptxValidatorNet8.exe (built from app-livedoc-construction-engine/src/core/PptxValidatorToolNet8) rather than reimplementing OOXML schema validation in Python. Trigger on: 'validate this pptx', 'why is this pptx corrupt', 'check OOXML/OPC compliance', 'fix this broken pptx', 'diagnose PowerPoint repair dialog'."
---

# PPTX/OOXML validation and structural auto-fix

Thin Python wrapper around a precompiled .NET tool — it does not reimplement OOXML schema validation (that's infeasible to match in pure Python; see rationale below). Two scripts, meant to be used in sequence.

## Why a wrapper, not a Python reimplementation

`PptxValidatorNet8.exe` layers custom checks (missing parts, broken relationships,
content-type/root-element mismatches, orphaned animation refs, broken ink
annotations, a severity classifier, and a ~450-line root-cause hint engine) on
top of `DocumentFormat.OpenXml.Validation.OpenXmlValidator` — Microsoft's own
OOXML schema validator. There is no Python equivalent of that validator;
reimplementing it against the full ECMA-376/ISO-29500 XSD tree would be a
large, error-prone undertaking with no guarantee of matching the SDK's
accuracy. Wrapping the existing, already-correct .NET binary is the right
tradeoff.

## Workflow

1. **Diagnose**: `python scripts/validate.py deck.pptx`
   Prints, per finding, exactly where the problem is (`partUri`/`xPath`) and
   a human-readable root-cause hint — this is the fast way to find *why* a
   file is broken, not just *that* it is. Add `--json` for the raw structured
   output (needed if you're going to hand-edit XML based on a `schemaFinding`).

2. **Auto-fix (structural only)**: `python scripts/autofix.py deck.pptx`
   Fixes dangling references — missing parts, broken relationships, orphaned
   animation refs, broken ink annotations — by removing the dangling
   reference itself. Writes `deck.fixed.pptx` by default (`--in-place` to
   overwrite, always keeps a `.bak`). Re-runs validate.py afterward and
   reports what's left.

   **Not auto-fixed, by design:**
   - `contentTypeMismatches` — ambiguous which side is wrong (the declared
     content type or the part's actual content); guessing wrong silently
     corrupts the file worse. Needs a human/Claude judgment call.
   - `schemaFindings` (arbitrary `OpenXmlValidator` errors) — too varied to
     rule-ify safely. Use the `hint`, `description`, `xPath`, and `nodeXml`
     fields from `validate.py --json` to locate the exact node (unzip the
     part, find it by `xPath`/`nodeXml`), hand-edit it, then re-zip and
     re-run `validate.py` to confirm the fix worked and didn't introduce a
     new error. This is the same isolate-one-variable-and-retest loop that
     works for manual PPTX repair generally — don't guess multiple fixes at
     once, fix one thing and re-validate.

3. **When schema findings or content-type mismatches remain**: don't trust
   "validator clean" alone as proof — PowerPoint's own repair can still
   trigger on things the SDK validator doesn't catch. If you have PowerPoint
   available, open the fixed file and confirm no repair dialog appears; if it
   still repairs, Save-As the repaired copy and diff the zip entries against
   your fixed version to see exactly what PowerPoint itself changed.

## Requirements

- `PptxValidatorNet8.exe` must already be built (Release, net10.0) at:
  `C:\project_new\app-livedoc-construction-engine\src\core\PptxValidatorToolNet8\bin\Release\net10.0\PptxValidatorNet8.exe`
  If missing, build it: `dotnet build -c Release` in that project directory.
  If the tool moves or is rebuilt elsewhere, pass `--exe <path>` to `validate.py`.
- Python 3.10+ (uses `X | None` union syntax), no third-party packages —
  stdlib `zipfile`/`xml.etree.ElementTree`/`subprocess`/`json` only.

## Scripts

| Script | Purpose |
|---|---|
| `scripts/validate.py <file\|dir...> [-r/--recursive] [-a/--all] [--json] [--office-version V] [--verbose] [--exe path]` | Run the validator, print a summary with root-cause hints, or raw JSON |
| `scripts/autofix.py <file.pptx> [-o out.pptx \| --in-place]` | Diagnose then auto-fix structural (dangling-reference) issues only |

`validate.py` accepts directories (optionally `--recursive`) in addition to
files, expanding them to every OOXML file inside (docx/docm/dotm/dotx,
pptx/pptm/potm/potx/ppam/ppsm/ppsx, xlsx/xlsm/xltm/xltx/xlam), skipping Office
lock files (`~$...`). By default only files with findings are shown; `--all`
also lists clean files. If a single file in a batch is unreadable enough to
crash the validator's own zip parsing (seen on a truncated/lock file), the
wrapper retries file-by-file so the rest of the batch's results aren't lost.
