"""Regression tests for `debug_format.cmd_check` body-structure and
relationship-target resolution.

Both checks had bugs that surfaced when the Rust port (`crisp-docx`)
ran cmd_check side-by-side against the real Vielfalt cs15.docx:

  * The body-structure check rejected `<w:bookmarkStart>` and
    `<w:bookmarkEnd>` as direct body children. The Vielfalt doc has
    them (2 leading, 2 trailing), and Word opens it cleanly — they
    are valid OOXML for bookmarks that span multiple paragraphs.

  * The relationship-target resolver computed a wrong base path for
    the package-root rels file `_rels/.rels`: it set base=".rels"
    instead of "" (empty). Every entry in that file then got the
    bogus prefix ".rels/", flagging legitimate targets as missing.

Tests below build minimal docx fixtures exercising each scenario and
assert cmd_check reports them clean.
"""

from __future__ import annotations

import io
import contextlib
from pathlib import Path

import argparse

from tests._helpers import W_DOC_HEAD, W_DOC_TAIL, write_docx
from debug_format import cmd_check  # type: ignore


def _run_check(path: Path) -> tuple[int, list[str], list[str]]:
    ns = argparse.Namespace(doc=str(path))
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = cmd_check(ns)
    oks = [ln.strip()[3:].strip() for ln in buf.getvalue().splitlines() if ln.lstrip().startswith("OK")]
    fails = [ln.strip()[5:].strip() for ln in buf.getvalue().splitlines() if ln.lstrip().startswith("FAIL")]
    return rc, oks, fails


def test_bookmark_start_end_allowed_as_direct_body_children(tmp_path: Path) -> None:
    """Vielfalt-style: 2 leading bookmarkStart + 2 trailing bookmarkEnd as
    direct body children. cmd_check must NOT flag these as 'unexpected
    element tags'."""
    body = (
        "<w:body>"
        '<w:bookmarkStart w:id="0" w:name="_top"/>'
        '<w:bookmarkStart w:id="1" w:name="intro"/>'
        "<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
        '<w:bookmarkEnd w:id="0"/>'
        '<w:bookmarkEnd w:id="1"/>'
        "<w:sectPr/>"
        "</w:body>"
    )
    doc = f"{W_DOC_HEAD}{body}{W_DOC_TAIL}"
    docx = write_docx(tmp_path / "bookmarks.docx", doc)

    rc, oks, fails = _run_check(docx)
    assert rc == 0, f"expected pass, got fails: {fails}"
    assert any("Body structure valid" in ok for ok in oks)
    assert not any("unexpected element tags" in f for f in fails)


def test_package_root_rels_resolves_targets_correctly(tmp_path: Path) -> None:
    """`_rels/.rels` is the package-level rels file. Its base path is the
    package root (empty string), not '.rels'. Targets like
    `word/document.xml` must resolve to actual ZIP entries."""
    # The standard _helpers fixture writes a sound _rels/.rels with a
    # valid Target="word/document.xml". A buggy resolver would compute
    # the resolved path as '.rels/word/document.xml' and flag it as
    # missing. Just check that this fixture passes cleanly.
    body = "<w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p><w:sectPr/></w:body>"
    doc = f"{W_DOC_HEAD}{body}{W_DOC_TAIL}"
    docx = write_docx(tmp_path / "rels.docx", doc)
    rc, oks, fails = _run_check(docx)
    assert rc == 0, f"expected pass; fails={fails}"
    assert any("All relationship targets present" in ok for ok in oks), f"oks={oks}"
    # Specifically: no FAIL line should mention '.rels/' as a resolved prefix.
    for f in fails:
        assert ".rels/word/" not in f, f"buggy resolver: {f}"


def test_legitimately_unexpected_body_child_still_flagged(tmp_path: Path) -> None:
    """The relaxation only adds bookmark elements to the allow-list;
    other unexpected children (e.g. a stray `<w:r>` outside a paragraph)
    should still be flagged."""
    body = "".join(
        [
            "<w:body>",
            "<w:r><w:t>orphan run</w:t></w:r>",
            "<w:p><w:r><w:t>hello</w:t></w:r></w:p>",
            "<w:sectPr/>",
            "</w:body>",
        ]
    )
    doc = f"{W_DOC_HEAD}{body}{W_DOC_TAIL}"
    docx = write_docx(tmp_path / "orphan.docx", doc)
    rc, _, fails = _run_check(docx)
    assert rc == 1, f"expected fail, got: {fails}"
    assert any("unexpected element tags" in f for f in fails), f"fails={fails}"
