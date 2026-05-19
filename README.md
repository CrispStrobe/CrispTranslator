# CrispTranslator

Four complementary tools for working with Word documents at the formatting level, plus a unified `docxtool` CLI:

| Tool | What it does |
|---|---|
| **Document Translator** | Translate `.docx` files across 200+ languages while preserving all formatting — down to bold/italic on individual words, footnotes, tables, headers, and footers |
| **Format Transplant** | Apply the complete formatting of a blueprint `.docx` to the content of a different document — page layout, styles, margins, everything — without translating anything |
| **DOCX Debugger** | Inspect, validate, and compare `.docx` files at the XML level — corruption checks, heading inference, footnote structure, style dumps, and side-by-side comparison |
| **RTF Notes → DOCX** | Convert RTF (or Markdown/DOCX) files whose citations are written as inline `[N]` markers followed by a trailing `Endnotes` list, producing a DOCX with *real* Word footnotes (or endnotes) — anchored, auto-numbered, and Word-clean |
| **docxtool** | One CLI with subcommands `notes`, `transplant`, `translate`, `debug`, and a standalone `clean` that strips rsid/paraId tracking attrs (the common cause of Word's "unreadable content" recovery dialog) |

All tools operate at the XML level of the OOXML format (`.docx`), preserving structure that higher-level APIs would silently discard.

---

## Table of Contents

- [Installation](#installation)
- [Document Translator](#document-translator)
  - [CLI](#translator-cli)
  - [Web UI](#translator-web-ui)
  - [How it works](#how-the-translator-works)
- [Format Transplant](#format-transplant)
  - [CLI](#transplant-cli)
  - [Web UI](#transplant-web-ui)
  - [How it works](#how-the-transplant-works)
- [DOCX Debugger](#docx-debugger)
  - [Subcommands](#debugger-subcommands)
- [RTF Notes → DOCX](#rtf-notes--docx)
- [docxtool (unified CLI)](#docxtool-unified-cli)
- [License](#license)

---

## Installation

### Requirements

- Python 3.10+
- The two core libraries are always required:

```bash
pip install python-docx lxml
```

- For the **web UIs**:

```bash
pip install gradio
```

- For the **Document Translator** (NMT models and alignment):

```bash
pip install torch ctranslate2 transformers huggingface_hub tqdm simalign
```

- For LLM backend backends (optional):

```bash
pip install openai anthropic fastapi-poe requests
```

### Optional: `fast_align`

Build the [`fast_align`](https://github.com/clab/fast_align) binary and put it on your `PATH` for an additional word-alignment backend. All other alignment backends work without it.

---

## Document Translator

Translate `.docx` files across 200+ languages while preserving formatting at run level: if word three in a sentence is bold in the source, word three's translation is bold in the output.

### Translator CLI

```bash
python translator.py input.docx output.docx -s en -t de
```

```
positional arguments:
  input                 Input .docx file
  output                Output .docx file

language:
  -s, --source          Source language code  (default: en)
  -t, --target          Target language code  (default: de)

mode:
  --mode {nmt, llm-align, llm-plain, hybrid}
                        nmt       – local NMT only (default)
                        hybrid    – NMT + optional LLM (recommended)
                        llm-align – LLM with local neural alignment
                        llm-plain – LLM, no alignment (fluent but loses inline formatting)

NMT backend:
  --nmt {nllb, madlad, opus, ct2, auto}
  --nllb-size {600M, 1.3B, 3.3B}

alignment:
  --aligner {awesome, simalign, lindat, fast_align, heuristic, auto}

LLM:
  --llm {openai, anthropic, ollama, groq}

  -v, --verbose         DEBUG logging
```

**Examples:**

```bash
# Fast general-purpose translation
python translator.py paper.docx paper_de.docx -s en -t de --nmt nllb

# High-quality academic text (3 GB RAM, slower)
python translator.py paper.docx paper_de.docx -s en -t de --nmt madlad

# LLM translation (Claude) with local alignment for formatting
python translator.py paper.docx paper_de.docx -s en -t es \
    --mode llm-align --llm anthropic

# Rare language with larger NLLB model
python translator.py doc.docx doc_uk.docx -s en -t uk --nmt nllb --nllb-size 1.3B

# Full debug trace
python translator.py doc.docx out.docx -s en -t fr -v 2>&1 | tee translate.log
```

### Translator Web UI

```bash
python translator-app.py
# → http://localhost:7860
```

The web UI includes a dynamic **Model Fetcher** that queries provider APIs (OpenAI, Groq, Anthropic, Ollama) to list available models and their capabilities.

---

## Format Transplant

Apply the complete formatting of a blueprint `.docx` to the content of a source `.docx`. No translation — pure layout transplant.

### Transplant CLI

```bash
python format_transplant.py blueprint.docx source.docx output.docx
```

```
positional arguments:
  blueprint             Blueprint DOCX — provides formatting
  source                Source DOCX — provides content
  output                Output DOCX path

options:
  -v, --verbose         DEBUG logging
  --style-map SRC=BP    Explicit style overrides
  --llm {openai, anthropic, groq, nebius, scaleway, openrouter, mistral, poe, ollama}
  --llm-model MODEL     Specific model ID (use 'auto' for default)
  --llm-batch N         Paragraphs per LLM call (default: 15, Groq: 5)
  --debug-limit N       Process only first N paragraphs (for testing)
  --styleguide-out PATH Save generated style guide to .md
  --styleguide-in PATH  Load pre-existing style guide
```

### How the Transplant Works

The transplant engine is designed for professional editorial standards:

1.  **Verbatim Reproduction:** LLM prompts are strictly constrained to ensure that text content is never summarized, paraphrased, or altered—only formatted.
2.  **Physical Tab Preservation:** Correctly detects and recreates physical `<w:tab/>` elements from the blueprint, ensuring professional spacing in footnotes.
3.  **Footnote Marker Precision:** Automatically extracts and applies the exact font, size, and vertical alignment of footnote numbers from the blueprint.
4.  **Robust Rate Limiting:** Implements exponential backoff, `retry-after` header parsing, and inter-batch delays to stay within strict provider tiers (e.g., Groq).
5.  **Environment Support:** Built-in lightweight `.env` loader for secure API key management.

---

## DOCX Debugger

`debug_format.py` is a standalone diagnostic toolkit for inspecting, validating, and comparing `.docx` files at the OOXML level.

### Subcommands

#### `footnotes` — Detailed footnote structure

```bash
python debug_format.py footnotes doc.docx
```

Inspects the internal structure of footnotes, identifying markers, text runs, and physical tab/space separators.

#### `xml` — Pretty-print ZIP parts

```bash
python debug_format.py xml doc.docx word/document.xml --strip-ns --exact
```

Directly inspects the raw XML of any part within the `.docx` archive. The `--exact` flag allows for surgical inspection of specific components like `word/footnotes.xml`.

---

## RTF Notes → DOCX

Many editorial workflows produce documents whose citations live inline as bracketed numbers (`…drawing its boundaries.[1]`) followed by a numbered `Endnotes` list at the bottom. `rtf_to_docx_endnotes.py` rewrites these into *real* Word notes that auto-number, anchor correctly, and survive editing.

### CLI

```bash
python rtf_to_docx_endnotes.py paper.rtf -o paper.docx
```

```
positional arguments:
  input                   source RTF / MD / DOCX

options:
  -o, --output            output .docx (default: same stem with .docx)
  --notes {footnotes,endnotes}
                          render notes as Word footnotes (default) or endnotes
  --reference-doc REF     pandoc reference docx; if omitted, one is built
                          on the fly with the body/heading-font options below
  --body-font NAME        body font for the auto-built reference (default: Times New Roman)
  --body-size PT          body size in points (default: 14)
  --heading-font NAME     heading font (default: Arial)
  --keep-bold             keep paragraph-wide **bold** wrappers from the source
                          (default: strip them; intra-paragraph emphasis is preserved)
  --no-strip-rsids        skip the rsid/paraId tracking-attr scrub
  --keep-intermediates    leave temp files in place for debugging
```

### How it works

1. **pandoc RTF → Markdown** with `--wrap=preserve`.
2. **Notes section detection**: header matching `/^#{1,6}\s*(end ?notes?|notes|footnotes|anmerkungen|endnoten|fußnoten|references)\s*$/i`, then per-note paragraphs starting with `[N]`. Numeric markers only — slide markers like `[S2]` and bracketed names like `[Liedhegener]` are left alone.
3. **Marker rewrite**: every digit-only `[N]` in the body becomes pandoc's footnote syntax `[^N]`; note bodies are appended as `[^N]: …` definitions.
4. **Whole-paragraph bold strip** (opt-out via `--keep-bold`): some editorial workflows cosmetically wrap every body paragraph in `**…**`; this is removed while leaving intra-paragraph emphasis intact.
5. **Auto-built reference docx**: starts from pandoc's default reference docx (so `FootnoteText`/`FootnoteReference` and friends remain defined), then patches `Normal` and `Heading 1-4` to the requested fonts/sizes, writing `w:rFonts` for all four scripts (`ascii`, `hAnsi`, `eastAsia`, `cs`) so Word doesn't fall back to the theme font.
6. **pandoc Markdown → DOCX** with that reference docx — output uses real `<w:footnoteReference>` elements anchored to entries in `word/footnotes.xml`.
7. **Endnotes mode** (`--notes endnotes`): post-processes the DOCX to rename `word/footnotes.xml` → `word/endnotes.xml`, rewrite references in `document.xml`, and patch `[Content_Types].xml` plus the relationship.
8. **rsid/paraId scrub**: strips `w14:paraId`, `w:rsidR`, `w:rsidRPr`, `w:rsidDel`, `w:rsidRDefault`, `w:rsidP`, `w:rsidTr`, `w:rsidSect` from every `<w:p>` and `<w:r>` (Word regenerates them on save; references to revision sessions that don't exist in `settings.xml` are a known cause of the "unreadable content" recovery dialog).

### Why these choices

A direct RTF→DOCX via Apple's `textutil` preserves the source's runs faithfully but emits OOXML that Word's strict validator rejects (non-standard tags like `w:sz-cs`, missing `styles.xml`, mis-ordered `<w:rPr>` children, malformed `customXml` relationships). The pandoc path produces Word-clean OOXML; the reference docx is how we recover *enough* visual fidelity (body and heading fonts/sizes) without inheriting textutil's quirks.

---

## docxtool (unified CLI)

`docxtool.py` is a single dispatcher that wraps every tool above plus a standalone `clean` subcommand. Each subcommand forwards to its sibling script, so the per-tool CLIs continue to work on their own.

```bash
python docxtool.py <subcommand> [options...]
```

| Subcommand | Wraps | What it does |
|---|---|---|
| `notes` | `rtf_to_docx_endnotes.py` | RTF/MD with `[N]` markers → DOCX with real footnotes/endnotes |
| `transplant` | `format_transplant.py` | Apply blueprint formatting to source content |
| `translate` | `translator.py` | Translate a docx, preserving run-level formatting |
| `debug` | `debug_format.py` | Inspect / validate / compare docx XML |
| `clean` | *(built-in)* | Strip rsid/paraId tracking attrs from a docx; optional non-standard tag normalization |

### `clean` standalone

```bash
python docxtool.py clean broken.docx                        # in place
python docxtool.py clean broken.docx -o fixed.docx          # to a new file
python docxtool.py clean broken.docx --dry-run              # report only
python docxtool.py clean broken.docx --also-normalize-tags  # + textutil quirks
python docxtool.py clean broken.docx --backend rust         # force native
python docxtool.py clean broken.docx --backend python       # force lxml
```

The `--backend` flag selects the implementation:

| Value | Behaviour |
|---|---|
| `auto` *(default)* | Use the [`crisp-docx`](https://github.com/CrispStrobe/crisp-docx) Rust wheel if installed; otherwise fall back to the lxml-backed Python implementation. |
| `rust` | Require the Rust wheel; fail with a clear error if it isn't available. |
| `python` | Force the Python implementation regardless of what's installed. |

`pip install crisp-docx` makes the Rust path available. Output reflects which backend was used (`stripped N attrs … (via crisp_docx)`). Both paths produce byte-identical results — the difference is throughput on large files.

Strips `w14:paraId`, `w14:textId`, `w:rsidR`, `w:rsidRPr`, `w:rsidDel`, `w:rsidRDefault`, `w:rsidP`, `w:rsidTr`, and `w:rsidSect` from every `<w:p>` and `<w:r>` in `word/document.xml`, `word/footnotes.xml`, and `word/endnotes.xml`. These attributes reference revision sessions registered in `settings.xml`'s `<w:rsids>`; when a body fragment from one document is grafted into another (transplant scenarios, partial recoveries, sed-style hand edits), the references go dangling and Word's strict validator fires the "unreadable content" recovery dialog. Stripping them is safe — Word regenerates fresh IDs the next time you save.

`--also-normalize-tags` additionally rewrites Apple `textutil`'s non-standard OOXML tags (`w:sz-cs` → `w:szCs`, `w:b-cs` → `w:bCs`, `w:i-cs` → `w:iCs`).

---

## Tests

Lightweight `unittest`-based suite covering the text-processing and
XML-surgery primitives behind `rtf_to_docx_endnotes.py` and `docxtool.py`.

```bash
python -m unittest discover tests -v
```

Requires only `python-docx` and `lxml`. No `pandoc` or `textutil` needed —
fixtures build minimal docx packages in-memory. CI runs the suite on Linux,
macOS, and Windows against Python 3.10–3.12 (see `.github/workflows/tests.yml`).

---

## License

GNU Affero General Public License v3.0 (AGPL-3.0). See [`LICENSE`](LICENSE).
