#!/usr/bin/env python3
"""Structural-only auto-fixer for PptxValidatorNet8 findings.

Scope (deliberately narrow — see SKILL.md for why): removes dangling
references reported at the ZIP/XML structural layer:
  - missingParts            -> drop the Override entry in [Content_Types].xml
  - brokenRels              -> drop the dangling <Relationship> in its .rels
  - orphanedAnimRefs        -> drop the animation block referencing a
                                nonexistent shape id (p:spTgt -> enclosing
                                p:par under p:childTnLst; p:bldP -> itself)
  - inkRelErrors            -> drop the mc:AlternateContent block whose
                                p:contentPart r:id doesn't resolve

Deliberately NOT auto-fixed:
  - contentTypeMismatches   -> ambiguous which side is wrong (declared type
                                vs. actual content); needs a human call
  - schemaFindings          -> arbitrary OpenXmlValidator errors; use the
                                `hint`/`description`/`xPath`/`nodeXml` fields
                                from validate.py to hand-edit the XML, then
                                re-run validate.py to confirm

Always writes to a new file unless --in-place is given, and --in-place
always keeps a .bak copy of the original.
"""

import argparse
import io
import json
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent

NS = {
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def q(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def run_validate(pptx_path: str) -> dict:
    proc = subprocess.run(
        [sys.executable, str(SCRIPT_DIR / "validate.py"), pptx_path, "--json"],
        capture_output=True,
        text=True,
    )
    if not proc.stdout.strip():
        raise RuntimeError(f"validate.py produced no output. stderr:\n{proc.stderr}")
    results = json.loads(proc.stdout)
    return results[0]


def build_parent_map(root: ET.Element) -> dict:
    return {child: parent for parent in root.iter() for child in parent}


def fix_content_types(xml_bytes: bytes, missing_parts: list[str]) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    wanted = {p.lstrip("/").lower() for p in missing_parts}
    removed = 0
    for override in list(root.findall(q("ct", "Override"))):
        part_name = (override.get("PartName") or "").lstrip("/").lower()
        if part_name in wanted:
            root.remove(override)
            removed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), removed


def fix_rels(xml_bytes: bytes, ids_to_remove: set[str]) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    removed = 0
    for rel in list(root.findall(q("rel", "Relationship"))):
        if rel.get("Id") in ids_to_remove:
            root.remove(rel)
            removed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), removed


def fix_orphaned_anim_refs(xml_bytes: bytes, findings: list[dict]) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    parent_map = build_parent_map(root)
    removed = 0
    for f in findings:
        ref_type = f["refType"]
        spid = f["shapeId"]
        for el in list(root.iter()):
            if local_name(el.tag) != ref_type or el.get("spid") != spid:
                continue
            if ref_type == "bldP":
                target = el
            else:  # spTgt: walk up to the direct child of p:childTnLst
                target = el
                while target in parent_map and local_name(parent_map[target].tag) != "childTnLst":
                    target = parent_map[target]
                if target not in parent_map:
                    continue
            parent = parent_map.get(target)
            if parent is not None:
                parent.remove(target)
                removed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), removed


def fix_ink_rel_errors(xml_bytes: bytes, findings: list[dict]) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    parent_map = build_parent_map(root)
    bad_ids = {f["relationshipId"] for f in findings}
    removed = 0
    for block in list(root.iter(q("mc", "AlternateContent"))):
        content_part = next((e for e in block.iter() if local_name(e.tag) == "contentPart"), None)
        if content_part is None:
            continue
        rid = content_part.get(q("r", "id"))
        if rid in bad_ids:
            parent = parent_map.get(block)
            if parent is not None:
                parent.remove(block)
                removed += 1
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), removed


def apply_fixes(pptx_path: Path, result: dict, entries: dict[str, bytes]) -> int:
    total_fixed = 0

    missing_parts = result.get("missingParts") or []
    if missing_parts and "[Content_Types].xml" in entries:
        entries["[Content_Types].xml"], n = fix_content_types(entries["[Content_Types].xml"], missing_parts)
        total_fixed += n
        print(f"  [Content_Types].xml: removed {n} Override entr{'y' if n == 1 else 'ies'} for missing parts")

    broken_rels = result.get("brokenRels") or []
    by_rels_file: dict[str, set[str]] = {}
    for r in broken_rels:
        by_rels_file.setdefault(r["relsFile"].lstrip("/"), set()).add(r["relationshipId"])
    for rels_file, ids in by_rels_file.items():
        if rels_file not in entries:
            continue
        entries[rels_file], n = fix_rels(entries[rels_file], ids)
        total_fixed += n
        print(f"  {rels_file}: removed {n} dangling relationship(s) — "
              f"re-run validate.py after this fix; anything that still referenced "
              f"the removed r:id will now show up as a new schema finding")

    anim_by_part: dict[str, list[dict]] = {}
    for f in result.get("orphanedAnimRefs") or []:
        anim_by_part.setdefault(f["partUri"].lstrip("/"), []).append(f)
    for part_uri, findings in anim_by_part.items():
        if part_uri not in entries:
            continue
        entries[part_uri], n = fix_orphaned_anim_refs(entries[part_uri], findings)
        total_fixed += n
        print(f"  {part_uri}: removed {n} orphaned animation reference(s)")

    ink_by_part: dict[str, list[dict]] = {}
    for f in result.get("inkRelErrors") or []:
        ink_by_part.setdefault(f["partUri"].lstrip("/"), []).append(f)
    for part_uri, findings in ink_by_part.items():
        if part_uri not in entries:
            continue
        entries[part_uri], n = fix_ink_rel_errors(entries[part_uri], findings)
        total_fixed += n
        print(f"  {part_uri}: removed {n} broken ink annotation reference(s)")

    return total_fixed


def rewrite_zip(src: Path, dst: Path, entries: dict[str, bytes]) -> None:
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, entries[item.filename])


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pptx", help="Path to the .pptx file to fix")
    parser.add_argument("-o", "--output", help="Output path (default: <name>.fixed.pptx)")
    parser.add_argument("--in-place", action="store_true", help="Overwrite the input, keeping a .bak copy")
    args = parser.parse_args()

    src = Path(args.pptx)
    if not src.is_file():
        print(f"File not found: {src}", file=sys.stderr)
        return 2

    print(f"Diagnosing {src} ...")
    result = run_validate(str(src))

    fixable_count = sum(len(result.get(k) or []) for k in
                         ("missingParts", "brokenRels", "orphanedAnimRefs", "inkRelErrors"))
    if fixable_count == 0:
        print("No structural (auto-fixable) issues found. "
              "Check contentTypeMismatches / schemaFindings via validate.py for anything remaining.")
        return 0

    with zipfile.ZipFile(src, "r") as zin:
        entries = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    print(f"Found {fixable_count} structural issue(s). Applying fixes:")
    fixed = apply_fixes(src, result, entries)

    if args.in_place:
        backup = src.with_suffix(src.suffix + ".bak")
        shutil.copy2(src, backup)
        dst = src
        print(f"Backed up original to {backup}")
    else:
        dst = Path(args.output) if args.output else src.with_suffix(".fixed" + src.suffix)

    tmp = dst.with_suffix(dst.suffix + ".tmp")
    rewrite_zip(src, tmp, entries)
    tmp.replace(dst)
    print(f"Wrote {fixed} fix(es) to {dst}")

    print("\nRe-validating fixed file ...")
    after = run_validate(str(dst))
    remaining = sum(len(after.get(k) or []) for k in
                     ("missingParts", "brokenRels", "orphanedAnimRefs", "inkRelErrors"))
    schema_left = len(after.get("schemaFindings") or [])
    ct_left = len(after.get("contentTypeMismatches") or [])
    print(f"Remaining structural issues: {remaining} | schema findings: {schema_left} | "
          f"content-type mismatches: {ct_left}")
    if remaining or schema_left or ct_left:
        print("Some issues remain — run validate.py --verbose on the output for hints, "
              "and open the file in PowerPoint to confirm it repairs cleanly.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
