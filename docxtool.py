#!/usr/bin/env python3
"""docxtool -- unified CLI for the CrispTranslator suite.

Subcommands dispatch to the standalone scripts in this repo, so every
existing tool keeps working on its own and gains a single, discoverable
entry point.

Usage:
  docxtool notes        -- RTF/MD with [N] citation markers -> DOCX with real
                          Word footnotes/endnotes  (rtf_to_docx_endnotes.py)
  docxtool transplant   -- apply a blueprint docx's formatting to source
                          content                       (format_transplant.py)
  docxtool translate    -- translate a docx, preserving formatting at the
                          run level                             (translator.py)
  docxtool debug        -- inspect / validate / compare docx XML
                                                            (debug_format.py)
  docxtool clean        -- strip rsid/paraId tracking attrs from a docx (the
                          "Word found unreadable content" cure)

Run `docxtool <subcommand> --help` for the full options of each.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


# --- subcommand handlers --------------------------------------------------


def _delegate(script_name: str, args: list[str]) -> int:
    """Forward to a sibling script using the current Python interpreter.

    Subprocess (rather than import) keeps the per-tool argparse code intact
    and avoids module-level side effects from the existing scripts.
    """
    script = SCRIPT_DIR / script_name
    if not script.exists():
        print(f"docxtool: missing sibling script {script}", file=sys.stderr)
        return 2
    return subprocess.call([sys.executable, str(script), *args])


def cmd_notes(args: list[str]) -> int:
    return _delegate("rtf_to_docx_endnotes.py", args)


def cmd_transplant(args: list[str]) -> int:
    return _delegate("format_transplant.py", args)


def cmd_translate(args: list[str]) -> int:
    return _delegate("translator.py", args)


def cmd_debug(args: list[str]) -> int:
    return _delegate("debug_format.py", args)


def cmd_clean(args: list[str]) -> int:
    """Strip rsid/paraId tracking attributes from a docx in place (or to OUT).

    Word's "found unreadable content" warning is often triggered by
    w14:paraId / w:rsidR / w:rsidRPr / w:rsidDel / w:rsidRDefault references
    pointing at session IDs not declared in settings.xml's <w:rsids>.
    Stripping these is safe - Word regenerates them on next save.
    """
    p = argparse.ArgumentParser(
        prog="docxtool clean",
        description=cmd_clean.__doc__.strip(),
    )
    p.add_argument("input", type=Path, help="docx to clean")
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="output path (default: edit input in place)",
    )
    p.add_argument(
        "--also-normalize-tags",
        action="store_true",
        help="also rewrite textutil's non-standard OOXML tags (w:sz-cs -> w:szCs)",
    )
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="report what would be changed but don't write",
    )
    ns = p.parse_args(args)

    src: Path = ns.input.expanduser().resolve()
    if not src.exists():
        print(f"docxtool clean: not found: {src}", file=sys.stderr)
        return 2
    dst: Path = src if ns.output is None else ns.output.expanduser().resolve()

    # Reuse the rsid-strip primitive from rtf_to_docx_endnotes.
    sys.path.insert(0, str(SCRIPT_DIR))
    from rtf_to_docx_endnotes import strip_rsids_from_docx  # type: ignore

    if dst != src:
        shutil.copyfile(src, dst)

    if ns.dry_run:
        # Count what *would* be stripped without writing.
        from lxml import etree

        total = 0
        with zipfile.ZipFile(src, "r") as zf:
            for key in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
                try:
                    data = zf.read(key)
                except KeyError:
                    continue
                root = etree.fromstring(data)
                from rtf_to_docx_endnotes import _RSID_PARA_ATTRS  # type: ignore

                for elem in root.iter():
                    if not isinstance(elem.tag, str):
                        continue
                    for a in elem.attrib:
                        if a in _RSID_PARA_ATTRS:
                            total += 1
        print(f"docxtool clean: would strip {total} rsid/paraId attrs from {src}")
        return 0

    removed = strip_rsids_from_docx(dst)
    print(f"docxtool clean: stripped {removed} rsid/paraId attrs -> {dst}")

    if ns.also_normalize_tags:
        renamed = _normalize_nonstandard_tags(dst)
        if renamed:
            print(f"docxtool clean: normalized {renamed} non-standard OOXML tags")
    return 0


# --- tag normalizer (Apple textutil quirks) -------------------------------

TAG_RENAMES = {
    b"w:sz-cs": b"w:szCs",
    b"w:b-cs": b"w:bCs",
    b"w:i-cs": b"w:iCs",
}


def _normalize_nonstandard_tags(docx_path: Path) -> int:
    """Rewrite textutil's non-OOXML local names (w:sz-cs -> w:szCs)."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        parts = {n: zf.read(n) for n in zf.namelist()}

    total = 0
    for key in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if key not in parts:
            continue
        data = parts[key]
        renamed_here = 0
        for old, new in TAG_RENAMES.items():
            cnt = data.count(old)
            if cnt:
                data = data.replace(old, new)
                renamed_here += cnt
        if renamed_here:
            parts[key] = data
            total += renamed_here

    if total:
        tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for n, b in parts.items():
                zf.writestr(n, b)
        tmp.replace(docx_path)
    return total


# --- dispatch -------------------------------------------------------------

COMMANDS = {
    "notes": (
        cmd_notes,
        "RTF/MD/DOCX with [N] markers -> DOCX with real footnotes/endnotes",
    ),
    "transplant": (
        cmd_transplant,
        "Apply blueprint formatting to source content",
    ),
    "translate": (
        cmd_translate,
        "Translate a docx, preserving formatting at the run level",
    ),
    "debug": (
        cmd_debug,
        "Inspect / validate / compare docx XML",
    ),
    "clean": (
        cmd_clean,
        "Strip rsid/paraId tracking attrs (Word 'unreadable content' cure)",
    ),
}


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else list(argv)
    if not args or args[0] in ("-h", "--help"):
        print(__doc__.strip(), file=sys.stderr)
        print("\nSubcommands:", file=sys.stderr)
        width = max(len(k) for k in COMMANDS) + 2
        for name, (_, desc) in COMMANDS.items():
            print(f"  {name:<{width}}{desc}", file=sys.stderr)
        return 0 if args else 1

    cmd = args[0]
    rest = args[1:]
    if cmd not in COMMANDS:
        print(f"docxtool: unknown subcommand '{cmd}'", file=sys.stderr)
        print(f"available: {', '.join(COMMANDS)}", file=sys.stderr)
        return 2
    handler, _ = COMMANDS[cmd]
    return handler(rest)


if __name__ == "__main__":
    raise SystemExit(main())
