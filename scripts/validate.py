#!/usr/bin/env python3
"""Thin Python wrapper around the precompiled PptxValidatorNet8.exe.

Runs the existing .NET OOXML validator (built from
app-livedoc-construction-engine/src/core/PptxValidatorToolNet8) against one
or more .pptx/.docx/.xlsx files and returns/prints a parsed, structured
result. Does not reimplement any validation logic itself.
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

# Update this if the tool is rebuilt to a different TFM/configuration.
DEFAULT_EXE = (
    r"C:\project_new\app-livedoc-construction-engine\src\core"
    r"\PptxValidatorToolNet8\bin\Release\net10.0\PptxValidatorNet8.exe"
)

# Every OOXML extension the validator's DetectFileType can route (it sniffs
# [Content_Types].xml for "presentationml"/"wordprocessingml"/"spreadsheetml",
# so macro-enabled/template variants all resolve correctly).
OOXML_EXTENSIONS = {
    ".docx", ".docm", ".dotm", ".dotx",
    ".pptx", ".pptm", ".potm", ".potx", ".ppam", ".ppsm", ".ppsx",
    ".xlsx", ".xlsm", ".xltm", ".xltx", ".xlam",
}


def expand_paths(paths: list[str], recursive: bool) -> list[str]:
    """Directories become the OOXML files inside them (like OOXML-Validator's
    directory/--recursive mode); PptxValidatorNet8 itself only takes files."""
    expanded: list[str] = []
    for p in paths:
        path = Path(p)
        if path.is_dir():
            it = path.rglob("*") if recursive else path.glob("*")
            expanded.extend(
                str(f) for f in it
                if f.is_file() and f.suffix.lower() in OOXML_EXTENSIONS
                and not f.name.startswith("~$")  # Office lock/temp files: same extension, not a real zip
            )
        else:
            expanded.append(p)
    return expanded


def find_exe(explicit: str | None) -> Path:
    candidate = Path(explicit) if explicit else Path(DEFAULT_EXE)
    if candidate.is_file():
        return candidate
    raise FileNotFoundError(
        f"PptxValidatorNet8.exe not found at {candidate}. "
        "Build it in app-livedoc-construction-engine (dotnet build -c Release) "
        "or pass --exe <path>."
    )


def _invoke(exe: Path, files: list[str], office_version: str, extra_args: list[str]) -> subprocess.CompletedProcess:
    cmd = [str(exe), *files, "--format", "json", "--office-version", office_version, *extra_args]
    return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")


def run_validator(exe: Path, files: list[str], office_version: str, extra_args: list[str]) -> list[dict]:
    """Runs the validator over all files in one batch. The exe has been seen to
    throw an unhandled exception (crashing the whole batch, no JSON output) on
    a single unreadable/non-zip file — e.g. an Office lock file that slipped
    through, or a genuinely truncated .pptx. When that happens, fall back to
    validating one file at a time so one bad file doesn't blank out results
    for the rest of the batch."""
    proc = _invoke(exe, files, office_version, extra_args)
    if proc.stdout.strip():
        try:
            return json.loads(proc.stdout)
        except json.JSONDecodeError as e:
            raise RuntimeError(f"Could not parse validator JSON output: {e}\nRaw output:\n{proc.stdout}")

    if len(files) == 1:
        raise RuntimeError(
            f"Validator produced no output for {files[0]} (exit {proc.returncode}). stderr:\n{proc.stderr}"
        )

    results: list[dict] = []
    for f in files:
        try:
            results.extend(run_validator(exe, [f], office_version, extra_args))
        except RuntimeError as e:
            results.append({"path": f, "fileNotFound": False, "sdkException": f"WrapperError: {e}"})
    return results


def has_errors(result: dict) -> bool:
    if result.get("fileNotFound") or result.get("sdkException"):
        return True
    for key in ("missingParts", "brokenRels", "contentTypeMismatches", "orphanedAnimRefs", "inkRelErrors", "schemaFindings"):
        if result.get(key):
            return True
    return False


def summarize(result: dict) -> str:
    path = result.get("path", "?")
    if result.get("fileNotFound"):
        return f"[FILE NOT FOUND] {path}"
    if result.get("sdkException"):
        return f"[SDK EXCEPTION] {path}: {result['sdkException']}"

    lines = [f"{'OK' if not has_errors(result) else 'INVALID'}: {path} ({result.get('fileType', '?')})"]
    for key, label in (
        ("missingParts", "Missing parts"),
        ("brokenRels", "Broken relationships"),
        ("contentTypeMismatches", "Content-type mismatches"),
        ("orphanedAnimRefs", "Orphaned animation refs"),
        ("inkRelErrors", "Ink relationship errors"),
    ):
        items = result.get(key) or []
        if items:
            lines.append(f"  {label}: {len(items)}")
            for item in items:
                lines.append(f"    - {item.get('detail', item)}")

    findings = result.get("schemaFindings") or []
    if findings:
        by_sev: dict[str, int] = {}
        for f in findings:
            by_sev[f.get("severity", "Unknown")] = by_sev.get(f.get("severity", "Unknown"), 0) + 1
        lines.append(f"  Schema findings: {len(findings)} ({', '.join(f'{k}={v}' for k, v in by_sev.items())})")
        for f in findings:
            hint = f" — {f['hint']}" if f.get("hint") else ""
            lines.append(f"    - [{f.get('severity')}] {f.get('partUri')}: {f.get('description')}{hint}")

    return "\n".join(lines)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "files", nargs="+",
        help="Path(s) to OOXML files or directories to validate (directories expand to the "
             "OOXML files inside them)",
    )
    parser.add_argument("--exe", help="Path to PptxValidatorNet8.exe (overrides default)")
    parser.add_argument("-r", "--recursive", action="store_true",
                         help="When a directory is passed, walk it recursively")
    parser.add_argument("-a", "--all", action="store_true",
                         help="Include files with no errors in the output (default: only show invalid files)")
    parser.add_argument(
        "--office-version",
        default="365",
        choices=["2007", "2010", "2013", "2016", "2019", "2021", "365"],
        help="Office version to validate against (default: 365)",
    )
    parser.add_argument("--json", action="store_true", help="Print raw JSON instead of a human summary")
    parser.add_argument("--verbose", action="store_true", help="Pass --verbose through to the validator")
    args = parser.parse_args()

    exe = find_exe(args.exe)
    files = expand_paths(args.files, args.recursive)
    if not files:
        print("No OOXML files found.", file=sys.stderr)
        return 2

    extra = ["--verbose"] if args.verbose else []
    results = run_validator(exe, files, args.office_version, extra)

    if not args.all:
        results = [r for r in results if has_errors(r)]

    if args.json:
        print(json.dumps(results, indent=2))
    else:
        if not results:
            print(f"OK: no errors in {len(files)} file(s).")
        for r in results:
            print(summarize(r))
            print()

    return 1 if any(has_errors(r) for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
