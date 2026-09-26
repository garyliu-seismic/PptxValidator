#!/usr/bin/env python3
"""Pure-Python OOXML validator — no .NET dependency.

Replaces the subprocess call to PptxValidatorNet8.exe with:
  1. ZIP/OPC structural checks  (missing parts, broken rels, content-type
     mismatches, orphaned animation refs, broken ink annotation rels)
  2. lxml + ISO-IEC 29500-4 XSD schema validation (pml.xsd / wml.xsd)
  3. PPTX-specific semantic checks  (chart axis refs, stacked label positions,
     shared master-theme, slide layout IDs, duplicate slide layouts,
     notes-slide refs, master-theme uniqueness)
  4. Severity classification  (Critical / High / Medium / Low)
  5. Root-cause hints  (ported from C# GetRootCauseHint)

Output schema is deliberately identical to PptxValidatorNet8's JSON so that
the existing validate.py / autofix.py wrappers can use either backend.

Usage (standalone):
    python scripts/py_validate.py deck.pptx [deck2.pptx ...] [--json] [--all]
    python scripts/py_validate.py ./folder --recursive --json

Programmatic:
    from scripts.py_validate import validate_file
    result = validate_file("deck.pptx")   # returns the same dict as the .NET JSON
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Any

import lxml.etree

# ── paths ──────────────────────────────────────────────────────────────────────
_SCRIPT_DIR  = Path(__file__).resolve().parent
_SCHEMAS_DIR = _SCRIPT_DIR.parent / "schemas"

# ── namespace constants ────────────────────────────────────────────────────────
_NS_CT   = "http://schemas.openxmlformats.org/package/2006/content-types"
_NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_P    = "http://schemas.openxmlformats.org/presentationml/2006/main"
_NS_MC   = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_NS_R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_W14  = "http://schemas.microsoft.com/office/word/2010/wordml"
_NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"

_OOXML_NAMESPACES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    "http://www.w3.org/XML/1998/namespace",
}

_IGNORED_VALIDATION_ERRORS = ["hyphenationZone", "purl.org/dc/terms"]

# lxml error message substrings that are false-positives vs .NET SDK:
#   buSzPct: ISO XSD requires "%"-suffix pattern (e.g. "80%") but Office
#            stores integer thousandths (e.g. 80000). The .NET SDK accepts
#            both forms internally; we suppress the pattern-facet failure.
_IGNORED_ERROR_SUBSTRINGS = [
    "buSzPct", # catches all buSzPct pattern/value errors
]

# Parts whose root element belongs to a Microsoft-proprietary or application-
# specific namespace with no matching ISO-29500-4 XSD.
# ppt/diagrams/drawing*.xml  root = dsp:drawing (ms/office/drawing/2008/diagram)
#   — .NET SDK has a built-in schema for this; ISO XSD does not. Skip to avoid
#     "no matching global declaration" noise, BUT re-validate with dml-diagram.xsd
#     later for the dml/dgm content errors the .NET tool does catch.
# customXml/item*.xml        — application-custom XML (Seismic/Word SDT store);
#     no standard XSD; .NET SDK skips schema validation for these entirely.
_SKIP_PART_PATTERNS = [
    # customXml/item*.xml: app-specific XML with no OOXML XSD; .NET skips schema validation
    re.compile(r"customXml/item\d+\.xml$", re.IGNORECASE),
    # customXml/itemProps*.xml: datastoreItem schema (oc-customXmlDataProperties.xsd)
    # root is {officeDocument/2006/customXml}datastoreItem — causes false-positive
    re.compile(r"customXml/itemProps\d+\.xml$", re.IGNORECASE),
    # NOTE: ppt/diagrams/drawing*.xml is NOT skipped here — it's handled by
    # _validate_diagrams_drawing() below which does targeted MinInclusive checks.
]

# Parts that have a dedicated XSD different from their parent-folder default.
# These are validated separately in the main loop.
_DIAGRAMS_DATA_RE  = re.compile(r"ppt/diagrams/(data|layout|colors|quickStyle)[^/]*\.xml$", re.IGNORECASE)
_DIAGRAMS_DRAW_RE  = re.compile(r"ppt/diagrams/drawing[^/]*\.xml$",                         re.IGNORECASE)

# XSD schema mapping: folder-prefix or file-name → schema path relative to _SCHEMAS_DIR
_SCHEMA_MAP = {
    "ppt":              "ISO-IEC29500-4_2016/pml.xsd",
    "word":             "ISO-IEC29500-4_2016/wml.xsd",
    "xl":               "ISO-IEC29500-4_2016/sml.xsd",
    "[Content_Types].xml": "ecma/fouth-edition/opc-contentTypes.xsd",
    "app.xml":          "ISO-IEC29500-4_2016/shared-documentPropertiesExtended.xsd",
    "core.xml":         "ecma/fouth-edition/opc-coreProperties.xsd",
    "custom.xml":       "ISO-IEC29500-4_2016/shared-documentPropertiesCustom.xsd",
    ".rels":            "ecma/fouth-edition/opc-relationships.xsd",
    "people.xml":       "microsoft/wml-2012.xsd",
    "commentsIds.xml":  "microsoft/wml-cid-2016.xsd",
    "commentsExtensible.xml": "microsoft/wml-cex-2018.xsd",
    "commentsExtended.xml":   "microsoft/wml-2012.xsd",
    "chart":            "ISO-IEC29500-4_2016/dml-chart.xsd",
    "theme":            "ISO-IEC29500-4_2016/dml-main.xsd",
    "drawing":          "ISO-IEC29500-4_2016/dml-main.xsd",
}

OOXML_EXTENSIONS = {
    ".docx", ".docm", ".dotm", ".dotx",
    ".pptx", ".pptm", ".potm", ".potx", ".ppam", ".ppsm", ".ppsx",
    ".xlsx", ".xlsm", ".xltm", ".xltx", ".xlam",
}

# ── XSD loader (cached) ────────────────────────────────────────────────────────
@lru_cache(maxsize=None)
def _load_schema(schema_path: str) -> lxml.etree.XMLSchema:
    with open(schema_path, "rb") as f:
        doc = lxml.etree.parse(f, parser=lxml.etree.XMLParser(), base_url=schema_path)
    return lxml.etree.XMLSchema(doc)


def _get_schema_path(rel_path: str) -> Path | None:
    """Return the XSD schema Path for a given relative part path, or None."""
    name = Path(rel_path).name
    if name in _SCHEMA_MAP:
        return _SCHEMAS_DIR / _SCHEMA_MAP[name]
    if rel_path.endswith(".rels"):
        return _SCHEMAS_DIR / _SCHEMA_MAP[".rels"]
    parts = rel_path.split("/")
    if len(parts) >= 2:
        folder = parts[0]
        if folder in _SCHEMA_MAP:
            return _SCHEMAS_DIR / _SCHEMA_MAP[folder]
        # charts/chartN.xml  or  theme/themeN.xml
        if len(parts) >= 3 and parts[1] == "charts" and name.startswith("chart"):
            return _SCHEMAS_DIR / _SCHEMA_MAP["chart"]
        if len(parts) >= 3 and parts[1] == "theme" and name.startswith("theme"):
            return _SCHEMAS_DIR / _SCHEMA_MAP["theme"]
    return None


# ── helpers ────────────────────────────────────────────────────────────────────
def _detect_file_type(path: str) -> str:
    """Return 'pptx', 'docx', 'xlsx', or 'unknown' by peeking at Content_Types."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            entry = next(
                (e for e in zf.infolist()
                 if e.filename.lower() == "[content_types].xml"), None)
            if entry is None:
                return "unknown"
            xml = zf.read(entry.filename).decode("utf-8", "replace")
        if "presentationml" in xml:  return "pptx"
        if "wordprocessingml" in xml: return "docx"
        if "spreadsheetml" in xml:    return "xlsx"
    except Exception:
        pass
    return "unknown"


def _resolve_target(base_folder: str, target: str) -> str:
    """Resolve a relationship Target relative to its parent folder."""
    if target.startswith("/"):
        return target.lstrip("/")
    parts: list[str] = [p for p in base_folder.split("/") if p]
    for seg in target.split("/"):
        if seg == "..":
            if parts:
                parts.pop()
        elif seg and seg != ".":
            parts.append(seg)
    return "/".join(parts)


def _read_zip(path: str) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as zf:
        return {e.filename: zf.read(e.filename) for e in zf.infolist()}


# ══════════════════════════════════════════════════════════════════════════════
# 1. OPC / ZIP structural checks (ported from PptxPartInspector.cs)
# ══════════════════════════════════════════════════════════════════════════════

def _find_missing_parts(entries: dict[str, bytes]) -> list[str]:
    ct_data = entries.get("[Content_Types].xml")
    if ct_data is None:
        return ["[Content_Types].xml is missing entirely"]
    root = lxml.etree.fromstring(ct_data)
    declared = [
        (e.get("PartName") or "").lstrip("/")
        for e in root.iter(f"{{{_NS_CT}}}Override")
    ]
    existing = {k.lower() for k in entries}
    return [d for d in declared if d and d.lower() not in existing]


def _find_broken_rels(entries: dict[str, bytes]) -> list[dict]:
    errors: list[dict] = []
    existing = {k.lower() for k in entries}
    for name, data in entries.items():
        if not (name.lower().endswith(".rels") and ("/_rels/" in name or name.startswith("_rels/"))):
            continue
        try:
            root = lxml.etree.fromstring(data)
        except Exception as e:
            errors.append({"relsFile": f"/{name}", "relationshipId": "(parse error)",
                           "target": "", "detail": f"Cannot parse .rels: {e}"})
            continue
        # base folder: ppt/slides/_rels/slide1.xml.rels → ppt/slides
        parent = re.sub(r"/_rels/[^/]+$", "", name)
        parent = re.sub(r"^_rels/[^/]+$", "", parent)
        for rel in root.iter(f"{{{_NS_RELS}}}Relationship"):
            mode = (rel.get("TargetMode") or "").lower()
            if mode == "external":
                continue
            rid    = rel.get("Id",     "")
            target = rel.get("Target", "")
            resolved = _resolve_target(parent, target)
            if resolved and resolved.lower() not in existing:
                errors.append({
                    "relsFile":       f"/{name}",
                    "relationshipId": rid,
                    "target":         resolved,
                    "detail": (f"Target part '{resolved}' declared in {name} "
                               f"(Id={rid}) does not exist in the package."),
                })
    return errors


_CT_ROOT_EXPECT = {
    "presentation.main": "presentation",
    "slide.main":        "sld",
    "slideLayout":       "sldLayout",
    "slideMaster":       "sldMaster",
    "theme":             "theme",
    "chart+xml":         "chartSpace",
    "drawingml.chart":   "chartSpace",
    "tableStyles":       "tblStyleLst",
    "viewProps":         "viewPr",
    "presProps":         "presentationPr",
}

def _find_content_type_mismatches(entries: dict[str, bytes]) -> list[dict]:
    results: list[dict] = []
    ct_data = entries.get("[Content_Types].xml")
    if ct_data is None:
        return results
    root = lxml.etree.fromstring(ct_data)
    for override in root.iter(f"{{{_NS_CT}}}Override"):
        part    = (override.get("PartName") or "").lstrip("/")
        ct      = override.get("ContentType") or ""
        expected_root = next((v for k, v in _CT_ROOT_EXPECT.items() if k in ct), None)
        if expected_root is None or part not in entries:
            continue
        try:
            actual_root = lxml.etree.fromstring(entries[part]).tag.rsplit("}", 1)[-1]
        except Exception:
            continue
        if actual_root.lower() != expected_root.lower():
            results.append({
                "partUri":      f"/{part}",
                "declaredType": ct,
                "actualRoot":   actual_root,
                "detail": (f"Part '{part}': content-type '{ct}' expects root "
                           f"<{expected_root}> but found <{actual_root}>."),
            })
    return results


_SLIDE_RE = re.compile(r"^ppt/slides/slide\d+\.xml$", re.IGNORECASE)

def _check_orphaned_anim_refs(entries: dict[str, bytes]) -> list[dict]:
    """Port of PptxPartInspector.CheckOrphanedAnimationRefs."""
    findings: list[dict] = []
    for name, data in entries.items():
        if not _SLIDE_RE.match(name):
            continue
        try:
            root = lxml.etree.fromstring(data)
        except Exception:
            continue
        sp_tree = root.find(f".//{{{_NS_P}}}spTree")
        if sp_tree is None:
            continue
        shape_ids = {
            el.get("id")
            for el in sp_tree.iter()
            if el.tag.split("}")[-1] == "cNvPr" and el.get("id")
        }
        timing = root.find(f".//{{{_NS_P}}}timing")
        if timing is None:
            continue
        seen: set[str] = set()
        for el in timing.iter():
            local = el.tag.split("}")[-1]
            if local not in ("spTgt", "bldP"):
                continue
            spid = el.get("spid")
            if spid and spid not in shape_ids and spid not in seen:
                seen.add(spid)
                findings.append({
                    "partUri": f"/{name}",
                    "shapeId": spid,
                    "refType": local,
                    "detail": (f"p:{local} spid=\"{spid}\" references a shape that "
                               f"does not exist in p:spTree — causes PowerPoint repair dialog."),
                })
    return findings


def _check_ink_rel_errors(entries: dict[str, bytes]) -> list[dict]:
    """Port of PptxPartInspector.CheckInkWithBrokenRels."""
    findings: list[dict] = []
    existing = {k.lower() for k in entries}
    for name, data in entries.items():
        if not _SLIDE_RE.match(name):
            continue
        try:
            root = lxml.etree.fromstring(data)
        except Exception:
            continue
        ink_blocks = [
            el for el in root.iter(f"{{{_NS_MC}}}AlternateContent")
            if any(c.tag.split("}")[-1] == "contentPart" for c in el.iter())
        ]
        if not ink_blocks:
            continue
        # load .rels for this slide
        slide_dir  = name.rsplit("/", 1)[0]
        slide_name = name.rsplit("/", 1)[-1]
        rels_path  = f"{slide_dir}/_rels/{slide_name}.rels"
        rel_targets: dict[str, str] = {}
        if rels_path in entries:
            try:
                rels_root = lxml.etree.fromstring(entries[rels_path])
                for rel in rels_root.iter(f"{{{_NS_RELS}}}Relationship"):
                    mode = (rel.get("TargetMode") or "").lower()
                    rid = rel.get("Id"); tgt = rel.get("Target")
                    if rid and tgt and mode != "external":
                        rel_targets[rid] = tgt
            except Exception:
                pass
        for block in ink_blocks:
            cp = next((e for e in block.iter() if e.tag.split("}")[-1] == "contentPart"), None)
            if cp is None:
                continue
            rid = cp.get(f"{{{_NS_R}}}id")
            if not rid:
                continue
            shape_name = next(
                (e.get("name") for e in block.iter() if e.tag.split("}")[-1] == "cNvPr"),
                "(unknown)"
            )
            if rid not in rel_targets:
                findings.append({
                    "partUri": f"/{name}", "relationshipId": rid,
                    "shapeName": shape_name, "resolvedTarget": None,
                    "detail": (f"Ink shape \"{shape_name}\" references r:id=\"{rid}\" "
                               f"but no such relationship exists in {rels_path}."),
                })
                continue
            resolved = _resolve_target(slide_dir, rel_targets[rid])
            if resolved.lower() not in existing:
                findings.append({
                    "partUri": f"/{name}", "relationshipId": rid,
                    "shapeName": shape_name, "resolvedTarget": resolved,
                    "detail": (f"Ink shape \"{shape_name}\" r:id=\"{rid}\" → "
                               f"target \"{resolved}\" does not exist in the package."),
                })
    return findings


# ══════════════════════════════════════════════════════════════════════════════
# 2. XSD schema validation (lxml + ISO-IEC 29500-4)
# ══════════════════════════════════════════════════════════════════════════════

def _clean_for_xsd(root_elem: lxml.etree._Element) -> lxml.etree._Element:
    """Strip non-OOXML namespace attrs/elements and mc:Ignorable so XSD validates cleanly."""
    xml_str = lxml.etree.tostring(root_elem, encoding="unicode")
    copy    = lxml.etree.fromstring(xml_str)
    # remove mc:Ignorable from root
    ignorable_attr = f"{{{_NS_MC}}}Ignorable"
    if ignorable_attr in copy.attrib:
        del copy.attrib[ignorable_attr]
    # strip non-OOXML namespace attrs and elements recursively
    _strip_non_ooxml(copy)
    return copy


def _strip_non_ooxml(elem: lxml.etree._Element) -> None:
    to_remove: list[lxml.etree._Element] = []
    for child in list(elem):
        if not callable(getattr(child, "tag", None)):
            tag = str(child.tag)
            if tag.startswith("{"):
                ns = tag.split("}")[0][1:]
                if ns not in _OOXML_NAMESPACES:
                    to_remove.append(child)
                    continue
        _strip_non_ooxml(child)
    for child in to_remove:
        elem.remove(child)
    attrs_to_remove = [
        a for a in list(elem.attrib)
        if a.startswith("{") and a.split("}")[0][1:] not in _OOXML_NAMESPACES
    ]
    for a in attrs_to_remove:
        del elem.attrib[a]


def _strip_template_tags(text: str | None) -> str | None:
    if text is None:
        return None
    return re.sub(r"\{\{[^}]*\}\}", "", text)


def _preprocess_presentation_xml(root: lxml.etree._Element) -> lxml.etree._Element:
    """Move <p:notesMasterIdLst> before <p:sldIdLst> if needed (XSD requires that order)."""
    ns = f"{{{_NS_P}}}"
    notes = root.find(f"{ns}notesMasterIdLst")
    slides = root.find(f"{ns}sldIdLst")
    if notes is None or slides is None:
        return root
    children = list(root)
    if children.index(notes) > children.index(slides):
        root.remove(notes)
        root.insert(list(root).index(slides), notes)
    return root


def _validate_part_xsd(rel_path: str, data: bytes,
                        is_presentation_xml: bool = False,
                        _override_schema: str | None = None) -> list[dict]:
    """Validate one XML part against its XSD. Returns list of schemaFinding dicts."""
    if _override_schema:
        schema_path = Path(_override_schema)
    else:
        schema_path = _get_schema_path(rel_path)
    if schema_path is None or not schema_path.exists():
        return []
    try:
        schema = _load_schema(str(schema_path))
    except Exception as e:
        e_str = str(e)
        if any(ign in e_str for ign in _IGNORED_VALIDATION_ERRORS):
            return []
        return [{
            "partUri":     f"/{rel_path}",
            "xPath":       "",
            "description": f"Could not load XSD for this part: {e}",
            "errorType":   "SchemaLoadError",
            "severity":    "Low",
            "hint":        "XSD schema load failure — part skipped.",
            "nodeLocalName": None,
            "nodeXml":     None,
        }]
    try:
        root = lxml.etree.fromstring(data)
        for el in root.iter():
            if el.text:  el.text = _strip_template_tags(el.text)
            if el.tail:  el.tail = _strip_template_tags(el.tail)
        if is_presentation_xml:
            root = _preprocess_presentation_xml(root)
        cleaned = _clean_for_xsd(root)
        doc     = lxml.etree.ElementTree(cleaned)
        if schema.validate(doc):
            return []
        findings = []
        for err in schema.error_log:
            msg = err.message
            if any(ign in msg for ign in _IGNORED_VALIDATION_ERRORS):
                continue
            # Gap fix #4: buSzPct integer format — .NET SDK accepts 80000, ISO XSD requires "80%"
            if any(sub in msg for sub in _IGNORED_ERROR_SUBSTRINGS):
                continue
            xpath = f"line:{err.line}"
            hint  = get_root_cause_hint(msg, xpath)
            sev   = get_severity(msg, xpath)
            findings.append({
                "partUri":     f"/{rel_path}",
                "xPath":       xpath,
                "description": msg,
                "errorType":   "SchemaError",
                "severity":    sev,
                "hint":        hint or None,
                "nodeLocalName": None,
                "nodeXml":     None,
            })
        return findings
    except Exception as e:
        return [{
            "partUri":     f"/{rel_path}",
            "xPath":       "",
            "description": f"Parse/validation error: {e}",
            "errorType":   "SchemaError",
            "severity":    "High",
            "hint":        None,
            "nodeLocalName": None,
            "nodeXml":     None,
        }]


def _validate_diagrams_drawing(rel_path: str, data: bytes) -> list[dict]:
    """Gap fix #3b: validate a:xfrm subtrees inside dsp:drawing (diagrams/drawing*.xml).

    The root element is a Microsoft-proprietary dsp:drawing which has no ISO XSD.
    However it embeds standard DrawingML (a: namespace) children that .NET validates
    against dml-main.xsd, catching MinInclusive violations on a:ext/@cx/cy etc.
    We replicate this by extracting each a:xfrm subtree and wrapping it in a minimal
    a:spPr envelope that dml-main.xsd can validate.
    """
    _NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    findings: list[dict] = []
    try:
        root   = lxml.etree.fromstring(data)
        schema = _load_schema(str(_SCHEMAS_DIR / "ISO-IEC29500-4_2016/dml-main.xsd"))
    except Exception:
        return findings

    # Check every a:off and a:ext for negative values directly
    # (avoids the complexity of re-wrapping; MinInclusive on ST_Coordinate /
    # ST_PositiveCoordinate is a numeric check we can do ourselves)
    _COORD_ATTRS = ("x", "y", "cx", "cy")
    _POSITIVE_ELEMS = {"ext", "chExt"}   # must be >= 0
    _ANY_COORD_ELEMS = {"off", "chOff"}  # can be negative per spec but
                                          # .NET flags them too in this context

    for el in root.iter():
        local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        ns    = el.tag.split("}")[0][1:] if "}" in el.tag else ""
        if ns != _NS_A:
            continue
        if local in _POSITIVE_ELEMS:
            for attr in _COORD_ATTRS:
                val = el.get(attr)
                if val is None:
                    continue
                try:
                    if int(val) < 0:
                        findings.append({
                            "partUri":     f"/{rel_path}",
                            "xPath":       f"line:{el.sourceline if hasattr(el, 'sourceline') else '?'}",
                            "description": (f"The MinInclusive constraint failed. "
                                            f"The value must be greater than or equal to 0."),
                            "errorType":   "SchemaError",
                            "severity":    "Medium",
                            "hint":        (f"a:{local} @{attr}={val} is negative. "
                                           f"SmartArt layout coordinates must be >= 0."),
                            "nodeLocalName": local,
                            "nodeXml":     None,
                        })
                except ValueError:
                    pass
    return findings


# ══════════════════════════════════════════════════════════════════════════════
# 3. PPTX semantic checks (ported from PPTXSchemaValidator)
# ══════════════════════════════════════════════════════════════════════════════

def _pptx_semantic_checks(entries: dict[str, bytes]) -> list[str]:
    """Run chart, theme-sharing, slide layout and notes-slide checks.
    Returns a list of human-readable problem strings."""
    problems: list[str] = []

    # lazy import helpers (they live next to this file)
    try:
        from helpers.pptx_chart import find_chart_problems
        from helpers.pptx_theme import live_shared_master_themes
    except ImportError:
        try:
            import importlib, sys as _sys
            _helpers = _SCRIPT_DIR / "helpers"
            if str(_helpers) not in _sys.path:
                _sys.path.insert(0, str(_SCRIPT_DIR))
            from helpers.pptx_chart import find_chart_problems
            from helpers.pptx_theme import live_shared_master_themes
        except ImportError:
            return problems   # helpers not available — skip

    # chart checks
    problems.extend(find_chart_problems(entries))

    # master-theme uniqueness
    problems.extend(live_shared_master_themes(entries))

    return problems


def _pptx_slide_layout_issues(entries: dict[str, bytes]) -> list[str]:
    """Check slide layout ID references in slide masters."""
    problems: list[str] = []
    master_re = re.compile(r"ppt/slideMasters/slideMaster\d+\.xml$", re.IGNORECASE)
    for name, data in entries.items():
        if not master_re.match(name):
            continue
        try:
            root = lxml.etree.fromstring(data)
        except Exception:
            continue
        # load rels
        parts      = name.rsplit("/", 1)
        rels_path  = f"{parts[0]}/_rels/{parts[1]}.rels"
        valid_rids: set[str] = set()
        if rels_path in entries:
            try:
                rr = lxml.etree.fromstring(entries[rels_path])
                for rel in rr.iter(f"{{{_NS_RELS}}}Relationship"):
                    if "slideLayout" in (rel.get("Type") or ""):
                        valid_rids.add(rel.get("Id") or "")
            except Exception:
                pass
        for el in root.iter(f"{{{_NS_P}}}sldLayoutId"):
            rid = el.get(f"{{{_NS_R}}}id")
            if rid and rid not in valid_rids:
                problems.append(
                    f"{name}: sldLayoutId r:id='{rid}' not found in slideLayout relationships")
    return problems


def _pptx_duplicate_layout_refs(entries: dict[str, bytes]) -> list[str]:
    problems: list[str] = []
    slide_rels_re = re.compile(r"ppt/slides/_rels/.+\.xml\.rels$", re.IGNORECASE)
    for name, data in entries.items():
        if not slide_rels_re.match(name):
            continue
        try:
            root = lxml.etree.fromstring(data)
        except Exception:
            continue
        layout_rels = [
            r for r in root.iter(f"{{{_NS_RELS}}}Relationship")
            if "slideLayout" in (r.get("Type") or "")
        ]
        if len(layout_rels) > 1:
            problems.append(
                f"{name}: slide has {len(layout_rels)} slideLayout references (max 1)")
    return problems


def _pptx_notes_slide_refs(entries: dict[str, bytes]) -> list[str]:
    """Each notes slide may be referenced by at most one regular slide."""
    import posixpath
    problems: list[str] = []
    slide_rels_re = re.compile(r"ppt/slides/_rels/.+\.xml\.rels$", re.IGNORECASE)
    notes_owners: dict[str, list[str]] = {}
    for name, data in entries.items():
        if not slide_rels_re.match(name):
            continue
        try:
            root = lxml.etree.fromstring(data)
        except Exception:
            continue
        for rel in root.iter(f"{{{_NS_RELS}}}Relationship"):
            if "notesSlide" not in (rel.get("Type") or ""):
                continue
            tgt = rel.get("Target") or ""
            base = posixpath.dirname(posixpath.dirname(name))   # ppt/slides
            resolved = _resolve_target(base, tgt)
            notes_owners.setdefault(resolved, []).append(name)
    for target, owners in notes_owners.items():
        if len(owners) > 1:
            problems.append(
                f"Notes slide '{target}' referenced by multiple slides: "
                + ", ".join(owners))
    return problems


# ══════════════════════════════════════════════════════════════════════════════
# 4. Severity classifier (ported from PptxValidator.GetSeverity)
# ══════════════════════════════════════════════════════════════════════════════

def get_severity(msg: str, xpath: str = "") -> str:
    # Low: unrecognised / extension attributes (both .NET and lxml phrasing)
    if "is not declared" in msg or "unrecognised attribute" in msg:
        return "Low"
    if "is not allowed" in msg and "attribute" in msg:
        return "Low"
    _low_children = ("leaderLines", "showLeaderLines", "showDLblsOverMax",
                     "showDataLabelsRange", "dlblFieldTable", "xForSave",
                     "dispBlanksAs", "smooth", "invertIfNegative", "bubble3D",
                     "explosion", "extLst")
    if "unexpected child element" in msg and any(c in msg for c in _low_children):
        return "Low"
    _font_locals = (":ea>", ":latin>", ":cs>", ":sym>")
    if ("unexpected child element" in msg and
            any(f in msg for f in _font_locals) and
            ("a:rPr" in xpath or "a:endParaRPr" in xpath)):
        return "Low"

    # Critical: presentation/master/layout incomplete
    if ("incomplete content" in msg and
            any(x in xpath for x in ("p:presentation", "p:sldMaster", "p:sldLayout"))):
        return "Critical"

    # Critical: bullet choice-group conflicts
    _bullet_elems = (":buNone", ":buChar", ":buAutoNum", ":buBlip", ":buClr",
                     ":buClrTx", ":buSzTx", ":buSzPct", ":buSzPts", ":buFont",
                     ":buFontTx", ":noFill", ":solidFill", ":gradFill", ":blipFill",
                     ":pattFill", ":grpFill", ":effectDag", ":uLnTx", ":uFillTx")
    if "unexpected child element" in msg and any(b in msg for b in _bullet_elems):
        return "Critical"

    # Critical: DOCX document/body incomplete
    if ("incomplete content" in msg and
            any(x in xpath for x in ("w:document", "w:body"))):
        return "Critical"

    # High: wrong namespace on cNvPr etc.
    _wrong_ns = ("drawingml/2006/main:cNvPr", "drawingml/2006/main:cNvSpPr",
                 "drawingml/2006/main:cNvPicPr", "drawingml/2006/main:cNvCxnSpPr",
                 "drawingml/2006/main:cNvGraphicFramePr")
    if any(n in msg for n in _wrong_ns):
        return "High"

    # High: table structure
    if ("required attribute 'h' is missing" in msg and "a:tr" in xpath) or \
       ("required attribute 'w' is missing" in msg and "a:gridCol" in xpath) or \
       ("'p:txBody'" in msg and "a:tc" in xpath):
        return "High"

    # High: core shape/content incomplete
    _high_xpaths = ("p:sp", "p:pic", "p:graphicFrame", "a:txBody",
                    "p:txBody", "p:blipFill", "p:spTree")
    if "incomplete content" in msg and any(x in xpath for x in _high_xpaths):
        return "High"

    if "required attribute" in msg and "is missing" in msg:
        return "High"

    if "InvalidAttribute" in msg and "r:id" in msg:
        return "High"

    # High: DOCX table/cell
    _docx_high = ("w:tbl", "w:tr", "w:tc")
    if "incomplete content" in msg and any(x in xpath for x in _docx_high):
        return "High"

    return "Medium"


# ══════════════════════════════════════════════════════════════════════════════
# 5. Root-cause hints (ported from PptxValidator.GetRootCauseHint)
# ══════════════════════════════════════════════════════════════════════════════

def get_root_cause_hint(msg: str, xpath: str = "") -> str:  # noqa: C901
    """Return a human-readable root-cause hint for a schema error message."""

    # ── Namespace mismatches ──────────────────────────────────────────────────
    if "drawingml/2006/main:cNvPr" in msg:
        return ("Wrong namespace: used A.NonVisualDrawingProperties (a:cNvPr) "
                "instead of P.NonVisualDrawingProperties (p:cNvPr) inside a slide shape.")
    if "drawingml/2006/main:cNvSpPr" in msg and "p:nvSpPr" in xpath:
        return "Wrong namespace: a:cNvSpPr inside p:nvSpPr. Use P.NonVisualShapeDrawingProperties."
    if "drawingml/2006/main:cNvPicPr" in msg and "p:nvPicPr" in xpath:
        return "Wrong namespace: a:cNvPicPr inside p:nvPicPr. Use P.NonVisualPictureDrawingProperties."
    if "drawingml/2006/main:cNvGraphicFramePr" in msg:
        return "Wrong namespace: a:cNvGraphicFramePr inside p:nvGraphicFramePr. Use P.NonVisualGraphicFrameDrawingProperties."

    # ── Table ────────────────────────────────────────────────────────────────
    if "'p:txBody'" in msg and "a:tc" in xpath:
        return "Table cell (a:tc) must use a:txBody, not p:txBody."
    if "required attribute 'h' is missing" in msg and "a:tr" in xpath:
        return "TableRow (a:tr) is missing required 'h' (height) attribute. Set e.g. 457200 EMU."
    if "required attribute 'w' is missing" in msg and "a:gridCol" in xpath:
        return "GridCol (a:gridCol) is missing required 'w' (width) attribute (in EMUs)."
    if "<a:tblGrid>" in msg:
        return "Table (a:tbl) is missing required a:tblGrid. Each column needs an a:gridCol with 'w'."
    if "incomplete content" in msg and "a:tc" in xpath:
        return "Table cell (a:tc) child order must be: a:txBody first, then a:tcPr."

    # ── Shape / spTree ────────────────────────────────────────────────────────
    if "incomplete content" in msg and "p:spTree" in xpath:
        return "ShapeTree (p:spTree) requires p:nvGrpSpPr and p:grpSpPr as first two children."
    if "incomplete content" in msg and "p:sp" in xpath:
        return "Shape (p:sp) needs p:nvSpPr, p:spPr, and p:txBody (in that order)."
    if "incomplete content" in msg and "p:grpSpPr" in xpath:
        return "GroupShapeProperties (p:grpSpPr) requires a:xfrm with a:off, a:ext, a:chOff, a:chExt."
    if "incomplete content" in msg and "p:graphicFrame" in xpath:
        return "GraphicFrame needs p:nvGraphicFramePr, p:xfrm (not p:spPr!), and a:graphic."
    if "incomplete content" in msg and "p:pic" in xpath:
        return "Picture (p:pic) needs p:nvPicPr, p:blipFill (a:blip r:embed + a:stretch), and p:spPr."

    # ── TextBody / paragraph ─────────────────────────────────────────────────
    if "incomplete content" in msg and "p:txBody" in xpath:
        return "TextBody (p:txBody) needs a:bodyPr and a:lstStyle before paragraphs."
    if "incomplete content" in msg and "a:txBody" in xpath:
        return "Drawing TextBody (a:txBody) requires a:bodyPr, then a:lstStyle, then a:p."
    if "incomplete content" in msg and "a:r" in xpath:
        return "Text run (a:r) requires at least an a:t child element."

    # ── RunProperties ────────────────────────────────────────────────────────
    if "InvalidValue" in msg and "sz" in msg and ("a:rPr" in xpath or "a:endParaRPr" in xpath):
        if 'sz="0"' in msg or "sz=0" in msg:
            return "RunProperties 'sz' is 0. Minimum valid is 100 (=1pt). Remove or set ≥100."
        return "RunProperties 'sz' (font size) out of range. Must be 100–400000 (hundredths of a pt)."
    if "'lang'" in msg and ("a:rPr" in xpath or "a:endParaRPr" in xpath):
        return "RunProperties 'lang' must be a valid BCP-47 tag (e.g. 'en-US'). Empty string is invalid."
    if "val" in msg and "srgbClr" in xpath:
        return "SrgbColor 'val' must be exactly 6 hex characters without '#' (e.g. val=\"FF0000\")."

    # ── Presentation ─────────────────────────────────────────────────────────
    if "incomplete content" in msg and "p:presentation" in xpath:
        return "Presentation element is missing required children (notesSz, sldMasterIdLst, or sldSz)."
    if "incomplete content" in msg and "p:sldMaster" in xpath:
        return "SlideMaster is empty — needs cSld > spTree with nvGrpSpPr + grpSpPr, plus txStyles and clrMap."
    if "incomplete content" in msg and "p:sldLayout" in xpath:
        return "SlideLayout needs at least cSld > spTree."

    # ── Chart ────────────────────────────────────────────────────────────────
    if "incomplete content" in msg and "c:valAx" in xpath:
        return "ValueAxis (c:valAx) needs: c:axId, c:scaling, c:delete, c:axPos, c:crossAx, c:crosses, c:crossBetween."
    if "incomplete content" in msg and "c:catAx" in xpath:
        return "CategoryAxis (c:catAx) needs: c:axId, c:scaling, c:delete, c:axPos, c:crossAx, c:crosses, c:auto, c:lblAlgn, c:lblOffset."
    if "incomplete content" in msg and "c:barChart" in xpath:
        return "BarChart (c:barChart) requires c:barDir, c:barGrouping, then series, then c:axId refs."
    if "incomplete content" in msg and "c:plotArea" in xpath:
        return "PlotArea (c:plotArea) needs at least one chart-type child (c:barChart, c:lineChart, etc.)."
    if "incomplete content" in msg and "c:chart" in xpath:
        return "Chart element (c:chart) requires a c:plotArea child."

    # ── Theme ────────────────────────────────────────────────────────────────
    if "incomplete content" in msg and "a:theme" in xpath:
        return "Theme (a:theme) requires a:themeElements > (a:clrScheme, a:fontScheme, a:fmtScheme)."
    if "incomplete content" in msg and "a:clrScheme" in xpath:
        return "ColorScheme (a:clrScheme) requires 12 color children: dk1, lt1, dk2, lt2, accent1–6, hlink, folHlink."

    # ── Generic fallbacks ────────────────────────────────────────────────────
    if "required attribute" in msg and "is missing" in msg:
        m = re.search(r"'([^']+)'", msg)
        attr = m.group(1) if m else "unknown"
        return f"Required attribute '{attr}' is missing."
    if "unexpected child element" in msg:
        m = re.search(r"'([^']+)'", msg)
        elem = m.group(1) if m else "unknown"
        return (f"Element '{elem}' is in the wrong position or uses the wrong namespace. "
                "Check child element order and namespace prefixes.")
    if "incomplete content" in msg:
        return "Element is missing required child elements. Check the OOXML schema for what is expected."

    return ""


# ══════════════════════════════════════════════════════════════════════════════
# 6. Top-level: validate_file
# ══════════════════════════════════════════════════════════════════════════════

def validate_file(path: str) -> dict[str, Any]:
    """Validate a single OOXML file. Returns a dict matching PptxValidatorNet8's JSON schema."""
    result: dict[str, Any] = {
        "path":                  path,
        "fileType":              "unknown",
        "fileNotFound":          False,
        "sdkException":          None,
        "missingParts":          [],
        "brokenRels":            [],
        "contentTypeMismatches": [],
        "orphanedAnimRefs":      [],
        "inkRelErrors":          [],
        "schemaFindings":        [],
        "semanticIssues":        [],   # new: PPTX-specific semantic problems
    }

    if not Path(path).is_file():
        result["fileNotFound"] = True
        return result

    file_type = _detect_file_type(path)
    result["fileType"] = file_type

    try:
        entries = _read_zip(path)
    except Exception as e:
        result["sdkException"] = f"CannotReadZip: {e}"
        return result

    # 1. structural checks
    result["missingParts"]          = _find_missing_parts(entries)
    result["brokenRels"]            = _find_broken_rels(entries)
    result["contentTypeMismatches"] = _find_content_type_mismatches(entries)
    result["orphanedAnimRefs"]      = _check_orphaned_anim_refs(entries)
    result["inkRelErrors"]          = _check_ink_rel_errors(entries)

    # 2. schema validation (all XML parts that have a known XSD)
    schema_findings: list[dict] = []
    if file_type in ("pptx", "docx"):
        for name, data in entries.items():
            if not name.endswith(".xml") and not name.endswith(".rels"):
                continue
            # skip charts in pptx — handled by semantic check
            if file_type == "pptx" and "charts/" in name and name.endswith(".xml"):
                continue
            # Gap fix #1/#3: skip ms-proprietary and app-custom parts
            if any(p.search(name) for p in _SKIP_PART_PATTERNS):
                continue
            # Gap fix #3a: diagrams/data|layout|colors|quickStyle — use dml-diagram.xsd
            if _DIAGRAMS_DATA_RE.search(name):
                schema_path = _SCHEMAS_DIR / "ISO-IEC29500-4_2016/dml-diagram.xsd"
                schema_findings.extend(_validate_part_xsd(name, data,
                                                           _override_schema=str(schema_path)))
                continue
            # Gap fix #3b: diagrams/drawing*.xml — root is ms-proprietary dsp:drawing,
            # but it embeds standard a:xfrm children. Validate only the a: subtree
            # against dml-main.xsd to catch MinInclusive violations (.NET catches these).
            if _DIAGRAMS_DRAW_RE.search(name):
                schema_findings.extend(_validate_diagrams_drawing(name, data))
                continue
            rel_path = name.lstrip("/")
            is_pres  = (rel_path == "ppt/presentation.xml")
            try:
                findings = _validate_part_xsd(rel_path, data, is_pres)
                schema_findings.extend(findings)
            except Exception as e:
                schema_findings.append({
                    "partUri": f"/{rel_path}",
                    "xPath": "", "description": str(e),
                    "errorType": "SchemaError", "severity": "High",
                    "hint": None, "nodeLocalName": None, "nodeXml": None,
                })
    result["schemaFindings"] = schema_findings

    # 3. PPTX semantic checks
    if file_type == "pptx":
        semantic: list[str] = []
        semantic.extend(_pptx_semantic_checks(entries))
        semantic.extend(_pptx_slide_layout_issues(entries))
        semantic.extend(_pptx_duplicate_layout_refs(entries))
        semantic.extend(_pptx_notes_slide_refs(entries))
        result["semanticIssues"] = semantic

    return result


def has_errors(result: dict) -> bool:
    if result.get("fileNotFound") or result.get("sdkException"):
        return True
    for key in ("missingParts", "brokenRels", "contentTypeMismatches",
                "orphanedAnimRefs", "inkRelErrors", "schemaFindings", "semanticIssues"):
        if result.get(key):
            return True
    return False


def summarize(result: dict) -> str:
    path = result.get("path", "?")
    if result.get("fileNotFound"):
        return f"[FILE NOT FOUND] {path}"
    if result.get("sdkException"):
        return f"[SDK EXCEPTION] {path}: {result['sdkException']}"

    status = "INVALID" if has_errors(result) else "OK"
    lines = [f"{status}: {path} ({result.get('fileType', '?')})"]

    for key, label in (
        ("missingParts",          "Missing parts"),
        ("brokenRels",            "Broken relationships"),
        ("contentTypeMismatches", "Content-type mismatches"),
        ("orphanedAnimRefs",      "Orphaned animation refs"),
        ("inkRelErrors",          "Ink relationship errors"),
    ):
        items = result.get(key) or []
        if items:
            lines.append(f"  {label}: {len(items)}")
            for item in items:
                lines.append(f"    - {item.get('detail', item)}")

    semantic = result.get("semanticIssues") or []
    if semantic:
        lines.append(f"  Semantic issues: {len(semantic)}")
        for s in semantic:
            lines.append(f"    - {s}")

    findings = result.get("schemaFindings") or []
    if findings:
        by_sev: dict[str, int] = {}
        for f in findings:
            sev = f.get("severity", "Unknown")
            by_sev[sev] = by_sev.get(sev, 0) + 1
        lines.append(f"  Schema findings: {len(findings)} "
                     f"({', '.join(f'{k}={v}' for k, v in by_sev.items())})")
        for f in findings:
            hint = f" — {f['hint']}" if f.get("hint") else ""
            lines.append(f"    - [{f.get('severity')}] {f.get('partUri')}: "
                         f"{f.get('description')}{hint}")

    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def _expand_paths(paths: list[str], recursive: bool) -> list[str]:
    expanded: list[str] = []
    for p in paths:
        path = Path(p)
        if path.is_dir():
            it = path.rglob("*") if recursive else path.glob("*")
            expanded.extend(
                str(f) for f in it
                if f.is_file()
                and f.suffix.lower() in OOXML_EXTENSIONS
                and not f.name.startswith("~$")
            )
        else:
            expanded.append(p)
    return expanded


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("files", nargs="+",
                        help="OOXML file(s) or directories to validate")
    parser.add_argument("-r", "--recursive", action="store_true",
                        help="Walk directories recursively")
    parser.add_argument("-a", "--all", action="store_true",
                        help="Show clean files too (default: only invalid)")
    parser.add_argument("--json", action="store_true",
                        help="Print raw JSON output")
    args = parser.parse_args()

    files = _expand_paths(args.files, args.recursive)
    if not files:
        print("No OOXML files found.", file=sys.stderr)
        return 2

    results = [validate_file(f) for f in files]

    if not args.all:
        display = [r for r in results if has_errors(r)]
    else:
        display = results

    if args.json:
        print(json.dumps(display, indent=2, default=str))
    else:
        if not display:
            print(f"OK: no errors in {len(files)} file(s).")
        for r in display:
            print(summarize(r))
            print()

    return 1 if any(has_errors(r) for r in results) else 0


if __name__ == "__main__":
    sys.exit(main())
