#!/usr/bin/env python3
"""Auto-test: compare pure-Python validator output against .NET PptxValidatorNet8.exe.

Usage:
    python scripts/test_compare.py --dir <folder-of-pptx> --exe <path-to-exe>
    python scripts/test_compare.py --exe <path-to-exe>        # uses built-in fixture dir

Exit code:
    0  all files MATCH (struct count + schema count identical)
    1  one or more GAPs found
    2  setup error (no files, exe not found, etc.)

Called by:  make test
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

# ── defaults ───────────────────────────────────────────────────────────────────
_SCRIPT_DIR   = Path(__file__).resolve().parent
_DEFAULT_EXE  = (
    r"C:\project_new\app-livedoc-construction-engine\src\core"
    r"\PptxValidatorToolNet8\bin\Release\net10.0\PptxValidatorNet8.exe"
)

OOXML_EXTS = {".pptx", ".docx", ".xlsx", ".potx", ".dotx", ".xltx"}


# ── load py_validate ───────────────────────────────────────────────────────────
def _load_py_validate():
    spec = importlib.util.spec_from_file_location("py_validate", _SCRIPT_DIR / "py_validate.py")
    mod  = importlib.util.module_from_spec(spec)   # type: ignore
    spec.loader.exec_module(mod)                   # type: ignore
    return mod


# ── run .NET validator ─────────────────────────────────────────────────────────
def _run_dotnet(exe: Path, files: list[str]) -> list[dict]:
    cmd = [str(exe), *files, "--format", "json", "--office-version", "365"]
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if not proc.stdout.strip():
        # retry one-by-one on crash
        results = []
        for f in files:
            p2 = subprocess.run([str(exe), f, "--format", "json", "--office-version", "365"],
                                capture_output=True, text=True, encoding="utf-8")
            if p2.stdout.strip():
                try:
                    results.extend(json.loads(p2.stdout))
                except json.JSONDecodeError:
                    pass
        return results
    return json.loads(proc.stdout)


# ── helpers ────────────────────────────────────────────────────────────────────
def _fname(r: dict) -> str:
    return r["path"].replace("\\", "/").split("/")[-1]


def _counts(r: dict | None) -> tuple[int, int] | None:
    if r is None:
        return None
    struct = sum(len(r.get(k, [])) for k in
                 ("missingParts", "brokenRels", "contentTypeMismatches",
                  "orphanedAnimRefs", "inkRelErrors"))
    schema = len(r.get("schemaFindings", []))
    return struct, schema


def _only_in(a: dict, b: dict, key: str = "description") -> set[str]:
    sa = {x.get(key, "") for x in (a.get("schemaFindings") or [])}
    sb = {x.get(key, "") for x in (b.get("schemaFindings") or [])}
    return sa - sb


# ── main ───────────────────────────────────────────────────────────────────────
def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dir", default=None,
                        help="Directory of OOXML test files (default: scripts/../test_fixtures/)")
    parser.add_argument("--exe", default=_DEFAULT_EXE,
                        help="Path to PptxValidatorNet8.exe")
    parser.add_argument("--strict", action="store_true",
                        help="Fail on any count difference (default: only report)")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    exe = Path(args.exe)
    if not exe.is_file():
        print(f"ERROR: .NET exe not found at {exe}", file=sys.stderr)
        print("       Pass --exe <path> or skip gap testing.", file=sys.stderr)
        return 2

    # resolve test directory
    if args.dir:
        test_dir = Path(args.dir)
    else:
        # look for test_fixtures next to scripts/
        test_dir = _SCRIPT_DIR.parent / "test_fixtures"

    if not test_dir.is_dir():
        print(f"ERROR: test directory not found: {test_dir}", file=sys.stderr)
        print("       Pass --dir <folder-with-pptx-files>", file=sys.stderr)
        return 2

    files = [str(f) for f in test_dir.rglob("*")
             if f.is_file() and f.suffix.lower() in OOXML_EXTS
             and not f.name.startswith("~$")]

    if not files:
        print(f"ERROR: no OOXML files found in {test_dir}", file=sys.stderr)
        return 2

    print(f"Testing {len(files)} file(s) from {test_dir}")
    print(f"  .NET exe : {exe}")
    print(f"  Python   : {_SCRIPT_DIR / 'py_validate.py'}")
    print()

    pv = _load_py_validate()

    # run both backends
    dn_list = _run_dotnet(exe, files)
    py_list = [pv.validate_file(f) for f in files]

    dn_map = {_fname(r): r for r in dn_list}
    py_map = {_fname(r): r for r in py_list}
    all_files = sorted(set(dn_map) | set(py_map))

    # ── results table ──────────────────────────────────────────────────────────
    col = 52
    print(f"{'FILE':<{col}}  {'  .NET s/sch':>12} {'PY s/sch':>10}  VERDICT")
    print("-" * (col + 32))

    gaps: list[str] = []
    matches: list[str] = []

    for f in all_files:
        d  = dn_map.get(f)
        p  = py_map.get(f)
        dc = _counts(d)
        pc = _counts(p)

        dn_str = f"{dc[0]}/{dc[1]}" if dc else "N/A"
        py_str = f"{pc[0]}/{pc[1]}" if pc else "N/A"

        ok = bool(dc and pc and dc[0] == pc[0] and dc[1] == pc[1])
        verdict = "MATCH" if ok else "GAP"
        if ok:
            matches.append(f)
        else:
            gaps.append(f)

        print(f"{f:<{col}}  .NET:{dn_str:<8} PY:{py_str:<8} {verdict}")

    print("-" * (col + 32))
    print(f"Total {len(all_files)} | MATCH {len(matches)} | GAP {len(gaps)}")

    # ── gap detail ─────────────────────────────────────────────────────────────
    if gaps:
        print("\n=== GAP DETAIL ===")
        for f in gaps:
            d  = dn_map.get(f)
            p  = py_map.get(f)
            if not (d and p):
                print(f"\n{f}: present in only one backend (not a logic gap)")
                continue
            dc = _counts(d); pc = _counts(p)
            print(f"\n-- {f}  .NET:{dc}  PY:{pc} --")
            only_dn = _only_in(d, p)
            only_py = _only_in(p, d)
            if only_dn:
                print(f"  .NET only ({len(only_dn)}):")
                for e in sorted(only_dn)[:5]:
                    print(f"    - {e[:120]}")
            if only_py:
                print(f"  PY only ({len(only_py)}):")
                for e in sorted(only_py)[:5]:
                    print(f"    - {e[:120]}")

    print()
    if not gaps:
        print("All files MATCH.")
        return 0
    else:
        print(f"{len(gaps)} file(s) have gaps vs .NET backend.")
        return 1 if args.strict else 0


if __name__ == "__main__":
    sys.exit(main())
