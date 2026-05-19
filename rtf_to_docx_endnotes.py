#!/usr/bin/env python3
r"""Convert an RTF (or Markdown / DOCX) with inline `[N]` citation markers and
a trailing "Endnotes" (or "Notes" / "Footnotes" / "Anmerkungen" / "References")
section into a DOCX whose notes are real Word footnotes (default) or endnotes.

Pipeline:
  1. pandoc <input> → Markdown.
  2. Find the notes section, parse each `[N] …` paragraph, rewrite inline
     `[N]` markers in the body to pandoc footnote refs `[^N]`, append
     `[^N]: …` definitions.
  3. pandoc Markdown → DOCX, using an auto-generated *reference docx* that
     applies Times New Roman 14pt to body and Arial bold to headings (a
     reasonable approximation of the source RTF's prevailing look).
  4. If --notes=endnotes, post-process the DOCX to convert footnotes to
     endnotes (rename word/footnotes.xml → word/endnotes.xml, rewrite the
     references in document.xml, fix content-types and rels).

textutil-based earlier versions of this script preserved formatting better
but produced docx that Word's strict validator rejects ("found unreadable
content"). The pandoc path produces clean OOXML that Word accepts; the
reference docx is how we recover *some* visual fidelity.

Usage:
  rtf_to_docx_endnotes.py INPUT [-o OUTPUT.docx]
                              [--notes footnotes|endnotes]
                              [--reference-doc REF.docx]
                              [--body-font NAME] [--body-size PT]
                              [--heading-font NAME]
                              [--keep-intermediates]
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


# --- markdown processing ---------------------------------------------------

NOTES_HEADER_RE = re.compile(
    r"^\s*#{1,6}\s*(end\s?notes?|notes|footnotes|anmerkungen|endnoten|"
    r"fußnoten|fussnoten|references)\s*$",
    re.IGNORECASE,
)
INLINE_RE = re.compile(r"\\?\[(\d+)\\?\]")
NOTE_START_RE = re.compile(r"^(?:\*\*\s*)?\\?\[(?P<num>\d+)\\?\]\s*")


def split_body_notes(md_text: str) -> tuple[str, str]:
    lines = md_text.splitlines()
    for i, line in enumerate(lines):
        if NOTES_HEADER_RE.match(line):
            return "\n".join(lines[:i]), "\n".join(lines[i + 1:])
    raise SystemExit(
        "Could not find a notes section header (e.g. '## Endnotes')."
    )


def parse_notes(notes_md: str) -> dict[int, str]:
    lines = notes_md.splitlines()
    starts: list[tuple[int, int]] = []
    for idx, line in enumerate(lines):
        m = NOTE_START_RE.match(line)
        if m:
            starts.append((idx, int(m.group("num"))))
    if not starts:
        return {}
    starts.append((len(lines), -1))

    notes: dict[int, str] = {}
    for (start_idx, num), (next_idx, _) in zip(starts, starts[1:]):
        chunk = "\n".join(lines[start_idx:next_idx]).strip()
        if chunk.startswith("**"):
            chunk = chunk[2:]
            if chunk.rstrip().endswith("**"):
                chunk = chunk.rstrip()[:-2].rstrip()
        chunk = NOTE_START_RE.sub("", chunk, count=1)
        chunk = re.sub(r"\s*\n\s*", " ", chunk).strip()
        if chunk:
            notes[num] = chunk
    return notes


def rewrite_body(body_md: str, valid_nums: set[int]) -> str:
    def repl(m: re.Match) -> str:
        n = int(m.group(1))
        return f"[^{n}]" if n in valid_nums else m.group(0)
    return INLINE_RE.sub(repl, body_md)


# Detect a *whole* paragraph wrapped in **...** — source RTF cosmetically
# bolds entire body paragraphs in some authoring workflows. We strip that
# wrapper while leaving intra-paragraph emphasis (`some **word** here`) alone.
_WHOLE_PARA_BOLD_RE = re.compile(
    r"\A(\*\*)(?P<inner>(?:(?!\*\*).)+)\*\*\Z", re.DOTALL,
)


def strip_paragraph_bold(body_md: str) -> tuple[str, int]:
    """Remove `**...**` wrappers around entire paragraphs. Returns (text, count)."""
    paragraphs = re.split(r"(\n[ \t]*\n)", body_md)
    out: list[str] = []
    stripped = 0
    for chunk in paragraphs:
        if chunk and not chunk.startswith("\n"):
            m = _WHOLE_PARA_BOLD_RE.match(chunk.strip())
            if m:
                # Replace the chunk preserving leading/trailing whitespace.
                lead_ws = chunk[: len(chunk) - len(chunk.lstrip())]
                trail_ws = chunk[len(chunk.rstrip()):]
                inner = m.group("inner")
                # Don't strip if inner itself contains a `**` that would
                # leave an unmatched marker.
                if "**" not in inner:
                    chunk = lead_ws + inner + trail_ws
                    stripped += 1
        out.append(chunk)
    return "".join(out), stripped


def build_footnoted_markdown(body: str, notes: dict[int, str]) -> str:
    parts = [body.rstrip(), ""]
    for n in sorted(notes):
        parts.append(f"[^{n}]: {notes[n]}")
        parts.append("")
    return "\n".join(parts)


# --- reference docx generation --------------------------------------------

def generate_reference_docx(
    out_path: Path, body_font: str, body_size_pt: float, heading_font: str
) -> None:
    """Start from pandoc's default reference docx and patch Normal/Heading
    fonts. Keeps all of pandoc's other styles (FootnoteText, FootnoteReference,
    SourceCode, etc.) intact — important because pandoc references them when
    writing footnotes/code blocks.
    """
    from docx import Document
    from docx.shared import Pt
    from docx.oxml.ns import qn
    from lxml import etree

    # Pull pandoc's default reference docx.
    try:
        default = subprocess.check_output(
            ["pandoc", "--print-default-data-file=reference.docx"]
        )
    except subprocess.CalledProcessError as e:
        raise SystemExit(f"pandoc cannot emit default reference docx: {e}")
    out_path.write_bytes(default)

    doc = Document(str(out_path))

    def set_font(style, font: str, size_pt: float, bold: bool | None = None) -> None:
        style.font.name = font
        style.font.size = Pt(size_pt)
        if bold is not None:
            style.font.bold = bold
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = etree.SubElement(rpr, qn("w:rFonts"))
        rfonts.set(qn("w:ascii"), font)
        rfonts.set(qn("w:hAnsi"), font)
        rfonts.set(qn("w:cs"), font)
        rfonts.set(qn("w:eastAsia"), font)

    try:
        set_font(doc.styles["Normal"], body_font, body_size_pt)
    except KeyError:
        pass

    heading_sizes = {
        "Heading 1": 20.0,
        "Heading 2": 16.0,
        "Heading 3": 14.0,
        "Heading 4": 12.0,
    }
    for name, size in heading_sizes.items():
        try:
            set_font(doc.styles[name], heading_font, size, bold=True)
        except KeyError:
            pass

    try:
        set_font(doc.styles["Title"], heading_font, 22.0, bold=True)
    except KeyError:
        pass

    doc.save(str(out_path))


# --- conversion driver ----------------------------------------------------

def rtf_to_markdown(src: Path, dst: Path) -> None:
    subprocess.run(
        ["pandoc", str(src), "-t", "markdown", "--wrap=preserve", "-o", str(dst)],
        check=True, capture_output=True, text=True,
    )


def markdown_to_docx(md_path: Path, docx_path: Path, ref_doc: Path | None) -> None:
    args = ["pandoc", str(md_path), "-f", "markdown", "-t", "docx",
            "-o", str(docx_path)]
    if ref_doc is not None:
        args.extend(["--reference-doc", str(ref_doc)])
    subprocess.run(args, check=True, capture_output=True, text=True)


# --- footnotes → endnotes post-processing (optional) ----------------------

def footnotes_to_endnotes(docx_path: Path) -> None:
    """Rewrite a pandoc-emitted DOCX so its footnotes become endnotes.

    Pandoc has no built-in option for endnotes in DOCX, so we patch the parts.
    """
    with zipfile.ZipFile(docx_path, "r") as zf:
        parts = {n: zf.read(n) for n in zf.namelist()}

    if "word/footnotes.xml" not in parts:
        return

    fn = parts.pop("word/footnotes.xml").decode("utf-8")
    en = fn
    en = en.replace("w:footnotes", "w:endnotes")
    en = en.replace("w:footnote ", "w:endnote ")
    en = en.replace("w:footnote>", "w:endnote>")
    en = en.replace("w:footnoteRef", "w:endnoteRef")
    en = en.replace('w:val="FootnoteText"', 'w:val="EndnoteText"')
    en = en.replace('w:val="FootnoteReference"', 'w:val="EndnoteReference"')
    parts["word/endnotes.xml"] = en.encode("utf-8")

    fn_rels = parts.pop("word/_rels/footnotes.xml.rels", None)
    if fn_rels is not None:
        parts["word/_rels/endnotes.xml.rels"] = fn_rels

    doc = parts["word/document.xml"].decode("utf-8")
    doc = doc.replace("w:footnoteReference", "w:endnoteReference")
    doc = doc.replace('w:val="FootnoteReference"', 'w:val="EndnoteReference"')
    parts["word/document.xml"] = doc.encode("utf-8")

    rels_key = "word/_rels/document.xml.rels"
    rels = parts[rels_key].decode("utf-8")
    rels = rels.replace(
        '/officeDocument/2006/relationships/footnotes"',
        '/officeDocument/2006/relationships/endnotes"',
    )
    rels = rels.replace('Target="footnotes.xml"', 'Target="endnotes.xml"')
    parts[rels_key] = rels.encode("utf-8")

    ct_key = "[Content_Types].xml"
    ct = parts[ct_key].decode("utf-8")
    ct = ct.replace("/word/footnotes.xml", "/word/endnotes.xml")
    ct = ct.replace("wordprocessingml.footnotes+xml",
                    "wordprocessingml.endnotes+xml")
    parts[ct_key] = ct.encode("utf-8")

    tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, b in parts.items():
            zf.writestr(n, b)
    tmp.replace(docx_path)


# --- rsid / paraId stripping ----------------------------------------------

# Word's strict validator (and sometimes its loader) rejects docs whose
# <w:p> / <w:r> elements reference revision-session IDs that aren't listed
# in settings.xml's <w:rsids>. Stripping these attrs is safe — Word
# regenerates them on save. Insight from cstr/FormatTransplant.
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_RSID_PARA_ATTRS = {
    f"{{{_W14_NS}}}paraId",
    f"{{{_W14_NS}}}textId",
    f"{{{_W_NS}}}rsidR",
    f"{{{_W_NS}}}rsidRPr",
    f"{{{_W_NS}}}rsidDel",
    f"{{{_W_NS}}}rsidRDefault",
    f"{{{_W_NS}}}rsidP",
    f"{{{_W_NS}}}rsidTr",
    f"{{{_W_NS}}}rsidSect",
}
_RSID_RUN_ATTRS = _RSID_PARA_ATTRS  # same set; cheap reuse


def _strip_rsids_in_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    from lxml import etree

    if not xml_bytes:
        return xml_bytes, 0
    try:
        root = etree.fromstring(xml_bytes)
    except etree.XMLSyntaxError:
        return xml_bytes, 0

    removed = 0
    for elem in root.iter():
        if not isinstance(elem.tag, str):
            continue
        for attr in list(elem.attrib):
            if attr in _RSID_PARA_ATTRS:
                del elem.attrib[attr]
                removed += 1
    if not removed:
        return xml_bytes, 0
    return etree.tostring(root, encoding="utf-8", xml_declaration=True), removed


def strip_rsids_from_docx(docx_path: Path) -> int:
    """Remove rsid/paraId attrs from document.xml, footnotes.xml, endnotes.xml.
    Returns the total number of attributes removed."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        parts = {n: zf.read(n) for n in zf.namelist()}
    total = 0
    for key in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if key in parts:
            new_bytes, removed = _strip_rsids_in_xml(parts[key])
            if removed:
                parts[key] = new_bytes
                total += removed
    if total:
        tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for n, b in parts.items():
                zf.writestr(n, b)
        tmp.replace(docx_path)
    return total


# --- driver ----------------------------------------------------------------

def convert(
    input_path: Path,
    output_path: Path,
    kind: str,
    ref_doc_arg: Path | None,
    body_font: str,
    body_size_pt: float,
    heading_font: str,
    keep: bool,
    keep_bold: bool,
    strip_rsids: bool,
) -> None:
    workdir = Path(tempfile.mkdtemp(prefix="rtf2docx_"))
    try:
        md_raw = workdir / "raw.md"
        md_rewritten = workdir / "rewritten.md"

        rtf_to_markdown(input_path, md_raw)
        raw_md = md_raw.read_text(encoding="utf-8")

        body, notes_section = split_body_notes(raw_md)
        notes = parse_notes(notes_section)
        if not notes:
            raise SystemExit("Notes section found but no [N] entries parsed.")

        valid_nums = set(notes.keys())
        new_body = rewrite_body(body, valid_nums)
        if not keep_bold:
            new_body, stripped = strip_paragraph_bold(new_body)
            if stripped:
                print(
                    f"info: stripped paragraph-wide bold from {stripped} paragraphs "
                    f"(use --keep-bold to disable)",
                    file=sys.stderr,
                )
        used_nums = {int(m.group(1)) for m in INLINE_RE.finditer(body)}
        missing = used_nums - valid_nums
        unused = valid_nums - used_nums
        if missing:
            print(f"warning: body cites [N] without matching note: {sorted(missing)}",
                  file=sys.stderr)
        if unused:
            print(f"warning: note defs never cited: {sorted(unused)}",
                  file=sys.stderr)

        md_rewritten.write_text(
            build_footnoted_markdown(new_body, notes), encoding="utf-8"
        )

        # Reference docx: user-supplied wins; otherwise auto-generate one.
        if ref_doc_arg is not None:
            ref_doc = ref_doc_arg
        else:
            ref_doc = workdir / "reference.docx"
            try:
                generate_reference_docx(
                    ref_doc, body_font, body_size_pt, heading_font
                )
            except Exception as e:
                print(f"warning: failed to build reference docx ({e}); "
                      f"falling back to pandoc default styles",
                      file=sys.stderr)
                ref_doc = None

        markdown_to_docx(md_rewritten, output_path, ref_doc)

        if kind == "endnotes":
            footnotes_to_endnotes(output_path)

        if strip_rsids:
            removed = strip_rsids_from_docx(output_path)
            if removed:
                print(
                    f"info: stripped {removed} rsid/paraId attributes "
                    f"(--no-strip-rsids to disable)",
                    file=sys.stderr,
                )

        print(
            f"wrote {output_path}  "
            f"(refs in body: {len(used_nums)}, note defs: {len(notes)}, "
            f"notes-as: {kind})"
        )
        if keep:
            print(f"intermediates kept in {workdir}", file=sys.stderr)
    finally:
        if not keep:
            shutil.rmtree(workdir, ignore_errors=True)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("input", type=Path, help="source RTF / MD / DOCX")
    p.add_argument(
        "-o", "--output", type=Path, default=None,
        help="output .docx (default: same stem as input, .docx extension)",
    )
    p.add_argument(
        "--notes", choices=("footnotes", "endnotes"), default="footnotes",
        help="render notes as Word footnotes (default) or endnotes",
    )
    p.add_argument(
        "--reference-doc", type=Path, default=None,
        help="path to a pandoc reference docx; if omitted, one is "
             "auto-generated with the body/heading font options below",
    )
    p.add_argument("--body-font", default="Times New Roman",
                   help="body font for the auto-generated reference docx")
    p.add_argument("--body-size", type=float, default=14.0,
                   help="body font size (pt) for the reference docx")
    p.add_argument("--heading-font", default="Arial",
                   help="heading font for the reference docx")
    p.add_argument(
        "--keep-bold", action="store_true",
        help="keep paragraph-wide **bold** wrapping from the source RTF "
             "(default: strip it; intra-paragraph emphasis is always kept)",
    )
    p.add_argument(
        "--no-strip-rsids", dest="strip_rsids", action="store_false",
        help="don't strip rsid/paraId tracking attrs from the output docx",
    )
    p.set_defaults(strip_rsids=True)
    p.add_argument(
        "--keep-intermediates", action="store_true",
        help="leave temp files in place for debugging",
    )
    args = p.parse_args(argv)
    input_path: Path = args.input.expanduser().resolve()
    output_path = (
        input_path.with_suffix(".docx")
        if args.output is None
        else args.output.expanduser().resolve()
    )
    convert(
        input_path, output_path, args.notes,
        args.reference_doc.expanduser().resolve() if args.reference_doc else None,
        args.body_font, args.body_size, args.heading_font,
        args.keep_intermediates, args.keep_bold, args.strip_rsids,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
