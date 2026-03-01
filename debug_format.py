#!/usr/bin/env python3
"""
debug_format.py – DOCX diagnostic toolkit
==========================================

Subcommands:
  inspect    General overview: styles, paragraphs, fonts, footnote count
  check      Corruption / validity checks (rsids, paraIds, rels, XML, body structure)
  headings   Heading structure analysis + property-based inference preview
  footnotes  Detailed footnote structure (run styles, separators, indentation)
  compare    Side-by-side style/heading/paragraph comparison of two documents
  styles     Full style dump (type, outline level, font, size, bold, italic)
  xml        Pretty-print any XML part from the ZIP archive

Usage:
  python debug_format.py inspect    doc.docx
  python debug_format.py check      doc.docx
  python debug_format.py headings   doc.docx
  python debug_format.py footnotes  doc.docx  [--id N]
  python debug_format.py compare    blueprint.docx  source.docx
  python debug_format.py styles     doc.docx  [--type paragraph|character|table]
  python debug_format.py xml        doc.docx  word/document.xml
"""

import argparse
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

# ── Dependency check ─────────────────────────────────────────────────────────

try:
    from lxml import etree
except ImportError:
    print("Missing dependency: pip install lxml")
    sys.exit(1)

try:
    from docx import Document
    from docx.oxml.ns import qn as docx_qn
except ImportError:
    print("Missing dependency: pip install python-docx")
    sys.exit(1)

# ── Namespace helpers ─────────────────────────────────────────────────────────

W    = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14  = "http://schemas.microsoft.com/office/word/2010/wordml"
W_NS = {"w": W, "w14": W14}

def w(tag: str)   -> str: return f"{{{W}}}{tag}"
def w14(tag: str) -> str: return f"{{{W14}}}{tag}"

def xpath(elem, expr: str) -> list:
    """XPath helper that works on both BaseOxmlElement and plain lxml._Element."""
    return etree._Element.xpath(elem, expr, namespaces=W_NS)

def strip_ns(xml_str: str) -> str:
    """Remove xmlns:* declarations for readable output."""
    return re.sub(r' xmlns:[^=]+="[^"]*"', "", xml_str)

def _get_text(elem) -> str:
    return "".join(t.text or "" for t in elem.iter(w("t")))

def _half_pt(val: str) -> str:
    """Convert half-point string to pt string."""
    try:
        return f"{int(val) / 2:.1f}pt"
    except (ValueError, TypeError):
        return val or ""

def _emu_pt(val: str) -> str:
    """Convert EMU string to pt string (1pt = 12700 EMU)."""
    try:
        return f"{int(val) / 12700:.1f}pt"
    except (ValueError, TypeError):
        return val or ""

def _twip_pt(val: str) -> str:
    """Convert twip string to pt string (1pt = 20 twips)."""
    try:
        return f"{int(val) / 20:.1f}pt"
    except (ValueError, TypeError):
        return val or ""

# ── XML part loader ───────────────────────────────────────────────────────────

def load_xml(path: str, part: str):
    """Open a ZIP archive and parse one XML part; return the lxml root."""
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(part))

def zip_names(path: str) -> Set[str]:
    with zipfile.ZipFile(path) as z:
        return set(z.namelist())

def zip_read(path: str, part: str) -> bytes:
    with zipfile.ZipFile(path) as z:
        return z.read(part)

# ── Style helpers ─────────────────────────────────────────────────────────────

def load_style_index(path: str) -> Dict[str, dict]:
    """Return dict style_id → {name, type, outline_level, sz_pt, bold, italic}."""
    root = load_xml(path, "word/styles.xml")
    idx: Dict[str, dict] = {}
    for s in root.iter(w("style")):
        sid  = s.get(w("styleId"), "")
        nm_e = s.find(w("name"))
        nm   = nm_e.get(w("val"), "") if nm_e is not None else ""
        stype = s.get(w("type"), "")

        ol_e = s.find(f".//{w('outlineLvl')}")
        ol   = int(ol_e.get(w("val"), "-1")) if ol_e is not None else -1

        sz_e = s.find(f".//{w('sz')}")
        sz   = _half_pt(sz_e.get(w("val"), "")) if sz_e is not None else ""

        b_e  = s.find(f".//{w('b')}")
        i_e  = s.find(f".//{w('i')}")
        idx[sid] = {
            "name": nm, "type": stype, "outline_level": ol,
            "sz": sz, "bold": b_e is not None, "italic": i_e is not None,
        }
    return idx

def style_name_from_pPr(pPr, style_idx: Dict[str, dict]) -> str:
    if pPr is None:
        return ""
    ps = pPr.find(w("pStyle"))
    if ps is None:
        return ""
    sid = ps.get(w("val"), "")
    return style_idx.get(sid, {}).get("name", sid)


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: inspect
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_inspect(args):
    path = args.doc
    print(f"\nInspecting: {path}")
    print("=" * 72)

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        doc_xml      = z.read("word/document.xml")
        settings_xml = z.read("word/settings.xml")

    style_idx = load_style_index(path)
    root      = etree.fromstring(doc_xml)
    body      = root.find(w("body"))

    # ── ZIP inventory ────────────────────────────────────────────────────────
    xml_parts = sorted(n for n in names if n.endswith(".xml") or n.endswith(".rels"))
    print(f"\n── ZIP ({len(names)} entries, {len(xml_parts)} XML/rels parts) ──")
    for n in xml_parts[:20]:
        print(f"   {n}")
    if len(xml_parts) > 20:
        print(f"   … and {len(xml_parts) - 20} more")

    # ── Page geometry ────────────────────────────────────────────────────────
    print("\n── Page geometry (first section) ──")
    pgSz = body.find(f".//{w('pgSz')}")
    pgMar = body.find(f".//{w('pgMar')}")
    if pgSz is not None:
        w_val = _twip_pt(pgSz.get(w("w"), ""))
        h_val = _twip_pt(pgSz.get(w("h"), ""))
        print(f"   Page size: {w_val} × {h_val}")
    if pgMar is not None:
        l = _twip_pt(pgMar.get(w("left"),   ""))
        r = _twip_pt(pgMar.get(w("right"),  ""))
        t = _twip_pt(pgMar.get(w("top"),    ""))
        b_ = _twip_pt(pgMar.get(w("bottom"), ""))
        print(f"   Margins: L={l} R={r} T={t} B={b_}")

    # ── Styles summary ───────────────────────────────────────────────────────
    para_styles = [(v["name"], k, v["outline_level"], v["sz"], v["bold"])
                   for k, v in style_idx.items() if v["type"] == "paragraph"]
    char_styles = [v["name"] for v in style_idx.values() if v["type"] == "character"]
    print(f"\n── Styles: {len(para_styles)} paragraph, {len(char_styles)} character ──")
    heading_styles = [(nm, sid, ol) for nm, sid, ol, *_ in para_styles if ol >= 0]
    if heading_styles:
        print("   Heading styles (with outlineLvl):")
        for nm, sid, ol in sorted(heading_styles, key=lambda x: x[2]):
            print(f"     H{ol + 1}  {nm!s:35} id={sid}")
    else:
        print("   No styles with outlineLvl found")

    # ── Body paragraph inventory ─────────────────────────────────────────────
    print("\n── Body paragraphs ──")
    style_freq: Dict[str, int] = {}
    all_paras = list(body.findall(w("p")))
    tables    = list(body.findall(w("tbl")))
    for p in all_paras:
        pPr  = p.find(w("pPr"))
        snm  = style_name_from_pPr(pPr, style_idx) or "(default)"
        style_freq[snm] = style_freq.get(snm, 0) + 1
    print(f"   {len(all_paras)} paragraphs, {len(tables)} tables")
    print("   Style frequencies:")
    for nm, cnt in sorted(style_freq.items(), key=lambda x: -x[1])[:15]:
        print(f"     {cnt:3}×  {nm}")

    # ── Footnotes ────────────────────────────────────────────────────────────
    fn_count = 0
    if "word/footnotes.xml" in names:
        fn_root = etree.fromstring(zip_read(path, "word/footnotes.xml"))
        fn_count = sum(1 for fn in fn_root
                       if int(fn.get(w("id"), "0") or "0") > 0)
    print(f"\n── Footnotes: {fn_count} ──")
    print()


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: check
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_check(args):
    path = args.doc
    print(f"\nCorruption/validity check: {path}")
    print("=" * 72)

    issues: List[str] = []
    ok_msgs: List[str] = []

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        xml_parts = [n for n in names if n.endswith(".xml") or n.endswith(".rels")]

        # ── 1. XML parse validity ─────────────────────────────────────────────
        xml_errors = []
        for name in xml_parts:
            try:
                etree.fromstring(z.read(name))
            except Exception as e:
                xml_errors.append(f"{name}: {e}")
        if xml_errors:
            for e in xml_errors:
                issues.append(f"XML parse error: {e}")
        else:
            ok_msgs.append(f"All {len(xml_parts)} XML/rels parts parse cleanly")

        # ── 2. rsid values vs settings.xml <w:rsids> ─────────────────────────
        doc_xml  = z.read("word/document.xml")
        doc_root = etree.fromstring(doc_xml)

        rsid_keys = [f"{{{W}}}{a}" for a in
                     ("rsidR", "rsidRPr", "rsidDel", "rsidRDefault", "rsidRPrChange")]
        body_rsids: Set[str] = set()
        for p in doc_root.iter(w("p")):
            for key in rsid_keys:
                val = p.get(key)
                if val:
                    body_rsids.add(val)

        sroot = etree.fromstring(z.read("word/settings.xml"))
        rsids_el = sroot.find(w("rsids"))
        settings_rsids: Set[str] = set()
        if rsids_el is not None:
            for c in rsids_el:
                val = c.get(w("val"))
                if val:
                    settings_rsids.add(val)

        missing = body_rsids - settings_rsids
        if missing:
            issues.append(
                f"{len(missing)} paragraph rsid value(s) not in settings.xml "
                f"<w:rsids> — causes 'Word found unreadable content'. "
                f"Sample: {sorted(missing)[:4]}"
            )
        elif body_rsids:
            ok_msgs.append(f"{len(body_rsids)} rsid value(s), all declared in settings.xml")
        else:
            ok_msgs.append("No rsid attributes in body paragraphs")

        # ── 3. w14:paraId uniqueness across all XML parts ────────────────────
        all_para_ids: List[Tuple[str, str]] = []  # (id_value, part_name)
        for name in xml_parts:
            if not name.endswith(".xml"):
                continue
            try:
                part_root = etree.fromstring(z.read(name))
                for p in part_root.iter(w("p")):
                    pid = p.get(w14("paraId"))
                    if pid:
                        all_para_ids.append((pid, name))
            except Exception:
                pass

        pid_counts = Counter(pid for pid, _ in all_para_ids)
        dupes = {pid: cnt for pid, cnt in pid_counts.items() if cnt > 1}
        if dupes:
            for pid, cnt in list(dupes.items())[:3]:
                parts = [n for p, n in all_para_ids if p == pid]
                issues.append(
                    f"Duplicate w14:paraId {pid!r} appears {cnt}× in: {parts}"
                )
        else:
            ok_msgs.append(
                f"{len(all_para_ids)} w14:paraId values across all parts, all unique"
            )

        # ── 4. Relationship targets present in ZIP ────────────────────────────
        rels_missing: List[Tuple[str, str, str]] = []
        for rel_name in (n for n in names if n.endswith(".rels")):
            try:
                rels_root = etree.fromstring(z.read(rel_name))
                # Base path of the file that owns these rels
                # e.g. word/_rels/document.xml.rels → base is word/
                base = rel_name.replace("_rels/", "").rsplit("/", 1)[0]
                for rel in rels_root:
                    if rel.get("TargetMode") == "External":
                        continue
                    target = rel.get("Target", "")
                    if target.startswith("/"):
                        full = target.lstrip("/")
                    else:
                        parts_ = (base + "/" + target).split("/")
                        resolved: List[str] = []
                        for part_ in parts_:
                            if part_ == "..":
                                if resolved:
                                    resolved.pop()
                            elif part_ and part_ != ".":
                                resolved.append(part_)
                        full = "/".join(resolved)
                    if full and full not in names:
                        rels_missing.append((rel_name, target, full))
            except Exception:
                pass

        if rels_missing:
            for rn, t, f in rels_missing[:5]:
                issues.append(f"Missing rel target {t!r} (resolved: {f!r}) in {rn}")
        else:
            ok_msgs.append("All relationship targets present in ZIP")

        # ── 5. Body structure ─────────────────────────────────────────────────
        body = doc_root.find(w("body"))
        if body is not None:
            children  = list(body)
            valid_tags = {w("p"), w("tbl"), w("sectPr")}
            bad_tags  = [c.tag for c in children if c.tag not in valid_tags]
            sect_last = bool(children) and children[-1].tag == w("sectPr")
            if bad_tags:
                issues.append(f"Body has unexpected element tags: {sorted(set(bad_tags))}")
            elif not sect_last and any(c.tag == w("sectPr") for c in children):
                issues.append("Body <w:sectPr> is not the last child")
            else:
                ok_msgs.append(
                    f"Body structure valid ({len(children)} children, "
                    f"sectPr at end: {sect_last})"
                )

        # ── 6. Bookmark ID uniqueness ─────────────────────────────────────────
        bm_ids = [p.get(w("id")) for p in doc_root.iter(w("bookmarkStart"))
                  if p.get(w("id"))]
        bm_dupes = {k: v for k, v in Counter(bm_ids).items() if v > 1}
        if bm_dupes:
            for bid, cnt in list(bm_dupes.items())[:3]:
                issues.append(f"Duplicate bookmarkStart id={bid!r} appears {cnt}×")
        else:
            ok_msgs.append(f"{len(bm_ids)} bookmarkStart ID(s), all unique")

        # ── 7. Inline relationship references in body ─────────────────────────
        body_xml_str = z.read("word/document.xml").decode(errors="replace")
        r_ids_in_body = set(re.findall(
            r'r:(?:id|embed|link)="(rId\d+)"', body_xml_str
        ))
        rels_file = "word/_rels/document.xml.rels"
        if rels_file in names:
            rels_xml_str = z.read(rels_file).decode(errors="replace")
            missing_rids = [rid for rid in r_ids_in_body
                            if f'Id="{rid}"' not in rels_xml_str]
            if missing_rids:
                issues.append(
                    f"Body references {len(missing_rids)} rId(s) not in "
                    f"document.xml.rels: {sorted(missing_rids)}"
                )
            else:
                ok_msgs.append(
                    f"{len(r_ids_in_body)} body relationship reference(s), "
                    f"all resolved"
                )

    # ── Summary ───────────────────────────────────────────────────────────────
    print()
    for msg in ok_msgs:
        print(f"  OK    {msg}")
    for msg in issues:
        print(f"  FAIL  {msg}")
    print()
    if issues:
        print(f"Result: {len(issues)} ISSUE(S) FOUND")
        return 1
    else:
        print("Result: PASS — no issues found")
        return 0


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: headings
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_headings(args):
    path = args.doc
    print(f"\nHeading analysis: {path}")
    print("=" * 72)

    style_idx = load_style_index(path)
    doc_root  = load_xml(path, "word/document.xml")
    body      = doc_root.find(w("body"))

    # ── Styles with outlineLvl ───────────────────────────────────────────────
    print("\n── Styles with explicit outline level (outlineLvl in XML) ──")
    heading_styles = [
        (v["outline_level"], v["name"], k, v["sz"], v["bold"])
        for k, v in style_idx.items()
        if v["outline_level"] >= 0 and v["type"] == "paragraph"
    ]
    if heading_styles:
        for ol, nm, sid, sz, bold in sorted(heading_styles, key=lambda x: x[0]):
            b_str = "bold" if bold else ""
            print(f"  H{ol + 1}  {sz:7}  {b_str:4}  {nm!s:40} id={sid}")
    else:
        print("  (none)")

    # ── Property-based heading inference (mirrors _infer_headings) ───────────
    print("\n── Property-based heading inference ──")
    print("   Signals: all-runs-bold OR pPr/rPr/w:b  +  short text (< 100 chars)")
    print()

    candidates: List[Tuple[str, float, str]] = []  # (text, size_pt, reason)
    body_sizes: List[float] = []

    for p in body.findall(w("p")):
        pPr     = p.find(w("pPr"))
        text    = _get_text(p).strip()
        if not text:
            continue

        # Paragraph-default bold and size
        ppr_bold   = False
        ppr_sz_pt: Optional[float] = None
        if pPr is not None:
            ppr_rPr = pPr.find(w("rPr"))
            if ppr_rPr is not None:
                ppr_bold = ppr_rPr.find(w("b")) is not None
                sz_el = ppr_rPr.find(w("sz"))
                if sz_el is not None:
                    try:
                        ppr_sz_pt = int(sz_el.get(w("val"), "0")) / 2.0
                    except (ValueError, TypeError):
                        pass

        # Run-level bold and size
        text_runs = [r for r in p.findall(w("r"))
                     if _get_text(r).strip()]
        run_bold_flags = []
        run_szs: List[float] = []
        for r in text_runs:
            rPr  = r.find(w("rPr"))
            r_bold = False
            if rPr is not None:
                r_bold = rPr.find(w("b")) is not None
                sz_el  = rPr.find(w("sz"))
                if sz_el is not None:
                    try:
                        run_szs.append(int(sz_el.get(w("val"), "0")) / 2.0)
                    except (ValueError, TypeError):
                        pass
            run_bold_flags.append(r_bold or ppr_bold)

        all_bold       = bool(text_runs) and all(run_bold_flags)
        effective_bold = all_bold or ppr_bold
        effective_sz   = (sum(run_szs) / len(run_szs)) if run_szs else ppr_sz_pt

        if effective_bold and 0 < len(text) < 100:
            reason = "pPr/rPr bold" if (ppr_bold and not all_bold) else "all-runs bold"
            candidates.append((text, effective_sz or 0.0, reason))
        elif effective_sz:
            body_sizes.append(effective_sz)

    if not candidates:
        print("  No heading candidates detected by property inference.")
    else:
        body_sz = Counter(body_sizes).most_common(1)[0][0] if body_sizes else 0.0
        unique_szs = sorted({sz for _, sz, _ in candidates if sz > 0}, reverse=True)
        heading_szs = [sz for sz in unique_szs
                       if body_sz == 0.0 or sz > body_sz + 0.4]
        if not heading_szs:
            heading_szs = [0.0]

        def _lvl(sz: float) -> int:
            if heading_szs == [0.0]:
                return 1
            for lvl_, threshold in enumerate(heading_szs, start=1):
                if sz >= threshold - 0.4:
                    return lvl_
            return len(heading_szs)

        print(f"  Body text reference size: "
              f"{body_sz:.1f}pt  (from {len(body_sizes)} non-heading para(s))")
        print(f"  Heading size tiers: {[f'{s:.1f}pt' for s in heading_szs] if heading_szs != [0.0] else ['all same']}")
        print()
        print(f"  {'Lvl':<5} {'Size':>7}  {'Signal':<16}  Text")
        print(f"  {'-'*5} {'-'*7}  {'-'*16}  {'-'*42}")
        for text, sz, reason in candidates:
            lvl = _lvl(sz)
            sz_str = f"{sz:.1f}pt" if sz else "n/a"
            print(f"  H{lvl:<4} {sz_str:>7}  {reason:<16}  {text[:55]!r}")

    print()


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: footnotes
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_footnotes(args):
    path    = args.doc
    target  = args.id  # optional footnote ID filter

    print(f"\nFootnote structure: {path}")
    print("=" * 72)

    names = zip_names(path)
    if "word/footnotes.xml" not in names:
        print("  This document has no footnotes.xml part.")
        return

    style_idx = load_style_index(path)
    fn_root   = etree.fromstring(zip_read(path, "word/footnotes.xml"))

    fn_count = 0
    for fn in fn_root:
        fn_id = fn.get(w("id"), "0")
        try:
            if int(fn_id) <= 0:
                continue
        except (ValueError, TypeError):
            continue

        if target is not None and fn_id != str(target):
            continue

        fn_count += 1
        text_preview = _get_text(fn)[:80]
        print(f"\nFootnote #{fn_id}: {text_preview!r}")

        for pi, p in enumerate(fn.findall(w("p"))):
            pPr = p.find(w("pPr"))

            # Style
            style_nm = style_name_from_pPr(pPr, style_idx) or "(default)"

            # Indentation
            ind_left = ind_hang = ind_first = ""
            if pPr is not None:
                ind = pPr.find(w("ind"))
                if ind is not None:
                    ind_left  = _twip_pt(ind.get(w("left"),      ""))
                    ind_hang  = _twip_pt(ind.get(w("hanging"),   ""))
                    ind_first = _twip_pt(ind.get(w("firstLine"), ""))

            ind_str = ""
            if ind_left:  ind_str += f"left={ind_left}"
            if ind_hang:  ind_str += f" hanging={ind_hang}"
            if ind_first: ind_str += f" firstLine={ind_first}"

            print(f"  Para {pi}: style={style_nm!r:28} indent={ind_str.strip() or 'none'}")

            # Runs
            for ri, r in enumerate(p.findall(w("r"))):
                rPr       = r.find(w("rPr"))
                t_text    = _get_text(r)
                fn_ref    = r.find(f".//{w('footnoteRef')}")
                fn_ref_in = r.find(f".//{w('footnoteReference')}")
                has_tab   = r.find(f".//{w('tab')}") is not None

                rStyle_val = ""
                sz_val     = ""
                va_val     = ""
                bold       = False
                italic     = False
                pos_val    = ""

                if rPr is not None:
                    rs = rPr.find(w("rStyle"))
                    if rs is not None:
                        rStyle_val = rs.get(w("val"), "")
                    sz = rPr.find(w("sz"))
                    if sz is not None:
                        sz_val = _half_pt(sz.get(w("val"), ""))
                    va = rPr.find(w("vertAlign"))
                    if va is not None:
                        va_val = va.get(w("val"), "")
                    bold   = rPr.find(w("b"))   is not None
                    italic = rPr.find(w("i"))   is not None
                    pos    = rPr.find(w("position"))
                    if pos is not None:
                        pos_val = pos.get(w("val"), "")

                details: List[str] = []
                if rStyle_val: details.append(f"rStyle={rStyle_val!r}")
                if sz_val:     details.append(f"sz={sz_val}")
                if va_val:     details.append(f"vertAlign={va_val}")
                if pos_val:    details.append(f"position={pos_val}")
                if bold:       details.append("bold")
                if italic:     details.append("italic")

                if fn_ref is not None:
                    label = "(footnoteRef marker)"
                elif fn_ref_in is not None:
                    label = "(footnoteReference anchor)"
                elif has_tab:
                    label = "<w:tab/> separator"
                elif t_text == "\t":
                    label = "TAB separator (text)"
                elif t_text == " ":
                    label = "SPACE separator"
                elif t_text == "":
                    label = "(empty run)"
                else:
                    label = repr(t_text[:30])

                print(f"    Run {ri}: {label:32} {', '.join(details)}")

            # Only show first paragraph unless --all flag
            if not getattr(args, "all_paras", False):
                remaining = len(fn.findall(w("p"))) - 1
                if remaining > 0:
                    print(f"    … +{remaining} more paragraph(s)")
                break

        if target is not None:
            break

    if fn_count == 0:
        if target is not None:
            print(f"  Footnote #{target} not found.")
        else:
            print("  No numbered footnotes found.")
    print()


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: compare
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_compare(args):
    a_path = args.doc_a
    b_path = args.doc_b

    print(f"\nComparing:")
    print(f"  A: {a_path}")
    print(f"  B: {b_path}")
    print("=" * 72)

    def _summary(path: str) -> dict:
        style_idx = load_style_index(path)
        root      = load_xml(path, "word/document.xml")
        body      = root.find(w("body"))

        heading_styles = sorted(
            [(v["outline_level"], v["name"], v["sz"])
             for v in style_idx.values()
             if v["type"] == "paragraph" and v["outline_level"] >= 0],
            key=lambda x: x[0]
        )

        paras  = list(body.findall(w("p")))
        tables = list(body.findall(w("tbl")))

        fn_count = 0
        if "word/footnotes.xml" in zip_names(path):
            fn_root  = etree.fromstring(zip_read(path, "word/footnotes.xml"))
            fn_count = sum(1 for fn in fn_root
                           if int(fn.get(w("id"), "0") or "0") > 0)

        # Style frequency in body
        style_freq: Dict[str, int] = {}
        for p in paras:
            pPr = p.find(w("pPr"))
            nm  = style_name_from_pPr(pPr, style_idx) or "(default)"
            style_freq[nm] = style_freq.get(nm, 0) + 1

        # para_styles set
        para_style_names = {v["name"] for v in style_idx.values()
                            if v["type"] == "paragraph"}
        char_style_names = {v["name"] for v in style_idx.values()
                            if v["type"] == "character"}

        return {
            "heading_styles": heading_styles,
            "para_count":  len(paras),
            "table_count": len(tables),
            "fn_count":    fn_count,
            "style_freq":  style_freq,
            "para_styles": para_style_names,
            "char_styles": char_style_names,
        }

    a = _summary(a_path)
    b = _summary(b_path)

    # ── Heading style comparison ─────────────────────────────────────────────
    print("\n── Heading styles (from outlineLvl) ──")
    max_len = max(len(a["heading_styles"]), len(b["heading_styles"]), 1)
    print(f"  {'Level':<6}  {'A':^40}  {'B':^40}")
    print(f"  {'-'*6}  {'-'*40}  {'-'*40}")
    all_levels = sorted(set(
        [ol for ol, _, _ in a["heading_styles"]] +
        [ol for ol, _, _ in b["heading_styles"]]
    ))
    a_map = {ol: (nm, sz) for ol, nm, sz in a["heading_styles"]}
    b_map = {ol: (nm, sz) for ol, nm, sz in b["heading_styles"]}
    for ol in all_levels:
        an, asz = a_map.get(ol, ("—", ""))
        bn, bsz = b_map.get(ol, ("—", ""))
        a_str = f"{an!s} ({asz})" if asz else an
        b_str = f"{bn!s} ({bsz})" if bsz else bn
        print(f"  H{ol + 1:<4}  {a_str:<40}  {b_str:<40}")

    # ── Style inventory comparison ────────────────────────────────────────────
    print("\n── Paragraph style inventory ──")
    only_a = a["para_styles"] - b["para_styles"]
    only_b = b["para_styles"] - a["para_styles"]
    shared = a["para_styles"] & b["para_styles"]
    print(f"  Shared:     {len(shared)}")
    print(f"  Only in A:  {len(only_a)}  {sorted(only_a)[:6]}")
    print(f"  Only in B:  {len(only_b)}  {sorted(only_b)[:6]}")

    # ── Body paragraph counts ─────────────────────────────────────────────────
    print("\n── Body content ──")
    print(f"  {'':30}  {'A':>8}  {'B':>8}")
    print(f"  {'Paragraphs':30}  {a['para_count']:>8}  {b['para_count']:>8}")
    print(f"  {'Tables':30}  {a['table_count']:>8}  {b['table_count']:>8}")
    print(f"  {'Footnotes':30}  {a['fn_count']:>8}  {b['fn_count']:>8}")

    # ── Style frequency comparison ────────────────────────────────────────────
    print("\n── Top styles used in body paragraphs ──")
    all_style_nms = sorted(
        set(a["style_freq"]) | set(b["style_freq"]),
        key=lambda n: -(a["style_freq"].get(n, 0) + b["style_freq"].get(n, 0))
    )
    print(f"  {'Style name':40}  {'A':>6}  {'B':>6}")
    print(f"  {'-'*40}  {'-'*6}  {'-'*6}")
    for nm in all_style_nms[:20]:
        a_cnt = a["style_freq"].get(nm, 0)
        b_cnt = b["style_freq"].get(nm, 0)
        flag  = " ←" if (a_cnt > 0) != (b_cnt > 0) else ""
        print(f"  {nm!s:40}  {a_cnt:>6}  {b_cnt:>6}{flag}")

    print()


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: styles
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_styles(args):
    path      = args.doc
    type_filt = getattr(args, "type", None)

    print(f"\nStyle dump: {path}")
    print("=" * 72)

    style_idx = load_style_index(path)

    type_map = {
        "paragraph": "paragraph",
        "character": "character",
        "table":     "table",
        "numbering": "numbering",
    }
    filt = type_map.get(type_filt or "", "")

    rows = []
    for sid, info in style_idx.items():
        if filt and info["type"] != filt:
            continue
        ol_str = f"H{info['outline_level'] + 1}" if info["outline_level"] >= 0 else "  "
        b_str  = "B" if info["bold"]   else " "
        i_str  = "I" if info["italic"] else " "
        rows.append((info["type"], ol_str, info["sz"], b_str, i_str, info["name"], sid))

    type_order = {"paragraph": 0, "character": 1, "table": 2, "numbering": 3}
    rows.sort(key=lambda r: (type_order.get(r[0], 9), r[1], r[5]))

    print(f"\n  {'Type':12}  {'Lvl':4}  {'Size':7}  {'BI':2}  {'Style name':40}  {'ID'}")
    print(f"  {'-'*12}  {'-'*4}  {'-'*7}  {'-'*2}  {'-'*40}  {'-'*30}")
    for stype, ol_str, sz, b_str, i_str, nm, sid in rows:
        print(f"  {stype:12}  {ol_str:4}  {sz:7}  {b_str}{i_str}   {nm!s:40}  {sid}")

    print()


# ═══════════════════════════════════════════════════════════════════════════════
# SUBCOMMAND: xml
# ═══════════════════════════════════════════════════════════════════════════════

def cmd_xml(args):
    path = args.doc
    part = args.part

    names = zip_names(path)

    # Fuzzy match if exact not found
    if part not in names:
        candidates = [n for n in names if part in n]
        if len(candidates) == 1:
            part = candidates[0]
        elif len(candidates) > 1:
            print(f"Ambiguous part '{part}'. Matches:")
            for c in candidates:
                print(f"  {c}")
            return
        else:
            print(f"Part '{part}' not found in {path}.")
            print("Available parts:")
            for n in sorted(names):
                print(f"  {n}")
            return

    raw = zip_read(path, part)

    if part.endswith(".xml") or part.endswith(".rels"):
        try:
            root = etree.fromstring(raw)
            pretty = etree.tostring(root, pretty_print=True).decode()
            # Optionally strip namespace declarations for readability
            if getattr(args, "strip_ns", False):
                pretty = strip_ns(pretty)
            print(pretty)
        except Exception as e:
            print(f"Parse error: {e}")
            print(raw.decode(errors="replace"))
    else:
        print(raw.decode(errors="replace"))


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

def main() -> int:
    parser = argparse.ArgumentParser(
        prog="debug_format.py",
        description="DOCX diagnostic toolkit for format-transplant debugging",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="cmd", metavar="COMMAND")

    # inspect
    p_inspect = sub.add_parser("inspect", help="General document overview")
    p_inspect.add_argument("doc", help="DOCX file")
    p_inspect.set_defaults(func=cmd_inspect)

    # check
    p_check = sub.add_parser(
        "check",
        help="Corruption/validity checks (rsids, paraIds, rels, XML, body structure)"
    )
    p_check.add_argument("doc", help="DOCX file")
    p_check.set_defaults(func=cmd_check)

    # headings
    p_head = sub.add_parser(
        "headings",
        help="Heading structure analysis + property-based inference preview"
    )
    p_head.add_argument("doc", help="DOCX file")
    p_head.set_defaults(func=cmd_headings)

    # footnotes
    p_fn = sub.add_parser(
        "footnotes",
        help="Detailed footnote structure (run styles, separators, indentation)"
    )
    p_fn.add_argument("doc", help="DOCX file")
    p_fn.add_argument("--id",  type=int, default=None, metavar="N",
                      help="Show only footnote #N")
    p_fn.add_argument("--all-paras", action="store_true",
                      help="Show all paragraphs per footnote (default: first only)")
    p_fn.set_defaults(func=cmd_footnotes)

    # compare
    p_cmp = sub.add_parser(
        "compare",
        help="Side-by-side comparison of two documents"
    )
    p_cmp.add_argument("doc_a", help="First DOCX (e.g. blueprint)")
    p_cmp.add_argument("doc_b", help="Second DOCX (e.g. source)")
    p_cmp.set_defaults(func=cmd_compare)

    # styles
    p_styles = sub.add_parser("styles", help="Full style dump")
    p_styles.add_argument("doc", help="DOCX file")
    p_styles.add_argument(
        "--type", choices=["paragraph", "character", "table", "numbering"],
        default=None, help="Filter by style type"
    )
    p_styles.set_defaults(func=cmd_styles)

    # xml
    p_xml = sub.add_parser(
        "xml",
        help="Pretty-print any XML part from the ZIP (e.g. word/document.xml)"
    )
    p_xml.add_argument("doc",  help="DOCX file")
    p_xml.add_argument("part", help="ZIP entry path, e.g. word/styles.xml")
    p_xml.add_argument("--strip-ns", action="store_true",
                       help="Remove xmlns:* declarations for readability")
    p_xml.set_defaults(func=cmd_xml)

    args = parser.parse_args()
    if not hasattr(args, "func"):
        parser.print_help()
        return 1

    result = args.func(args)
    return result if isinstance(result, int) else 0


if __name__ == "__main__":
    sys.exit(main())
