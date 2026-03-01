# CrispTranslator

Three complementary tools for working with Word documents at the formatting level:

| Tool | What it does |
|---|---|
| **Document Translator** | Translate `.docx` files across 200+ languages while preserving all formatting — down to bold/italic on individual words, footnotes, tables, headers, and footers |
| **Format Transplant** | Apply the complete formatting of a blueprint `.docx` to the content of a different document — page layout, styles, margins, everything — without translating anything |
| **DOCX Debugger** | Inspect, validate, and compare `.docx` files at the XML level — corruption checks, heading inference, footnote structure, style dumps, and side-by-side comparison |

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

## License

GNU Affero General Public License v3.0 (AGPL-3.0). See [`LICENSE`](LICENSE).
