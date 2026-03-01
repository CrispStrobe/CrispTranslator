# CrispTranslator

Two complementary tools for working with Word documents at the formatting level:

| Tool | What it does |
|---|---|
| **Document Translator** | Translate `.docx` files across 200+ languages while preserving all formatting — down to bold/italic on individual words, footnotes, tables, headers, and footers |
| **Format Transplant** | Apply the complete formatting of a blueprint `.docx` to the content of a different document — page layout, styles, margins, everything — without translating anything |

Both tools operate at the XML level of the OOXML format (`.docx`), preserving structure that higher-level APIs would silently discard.

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

- For LLM translation backends (optional):

```bash
pip install openai anthropic
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
  --llm {openai, anthropic, ollama}

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

The web UI wraps the engine with asynchronous processing, real-time log streaming, and environment auto-setup (CTranslate2 install check, API key detection). It is designed to deploy on Hugging Face Spaces without changes.

### How the Translator Works

The engine follows an **Extract → Translate → Align → Reconstruct** pipeline:

#### 1. Document parsing

`python-docx` walks the entire document tree — body paragraphs, table cells, headers, footers, and footnotes — and converts each paragraph into a `TranslatableParagraph` carrying `FormatRun` objects (text + bold/italic/underline/font/size/color) and full layout metadata (indents, spacing, alignment). The font hierarchy is resolved up the style chain to prevent Word from reverting to theme defaults.

#### 2. Translation backends

All local models are accelerated via **CTranslate2** for int8 inference (ARM64 / Apple Silicon optimised) or float16 (NVIDIA CUDA). Backends are tried in a fallback chain.

| Backend | Model | Languages | RAM | Speed |
|---|---|---|---|---|
| **NLLB-200** (default) | Meta, 600 M–3.3 B | 200+ | ~1–4 GB | Fast |
| **Madlad-400** | Google, 3 B | 200+ | ~3 GB | Medium |
| **Opus-MT** | Helsinki-NLP | Specific pairs | ~200 MB | Very fast |
| **WMT21/CT2** | Meta, dense | European | ~6 GB | Medium |
| **LLM** | OpenAI / Anthropic / Ollama | Any | API | Slow |

#### 3. Word alignment

Inline formatting (bold, italic) is preserved by finding where each source word landed in the translated sentence. The aligner chain, in priority order:

1. **Awesome-Align** — CT2-optimised BERT, Mutual Argmax (highest precision)
2. **SimAlign** — PyTorch BERT embeddings
3. **Lindat API** — zero local RAM, network-dependent
4. **fast_align** — classical IBM models (binary required)
5. **Heuristic** — shared-word matching (ultimate fallback)

#### 4. Reconstruction

Runs are cleared from the paragraph XML and rebuilt from scratch. Each target word gets the inline style of its aligned source word. Font name, size, and color are written directly into `<w:rFonts>` XML to bypass Word's theme system. Footnote anchors (`<w:footnoteReference>`) are extracted before clearing and re-attached after. The footnote XML blob is committed back into the document's binary part.

#### Translation modes

| Mode | Quality | Speed | Notes |
|---|---|---|---|
| `nmt` | Good | Fast | Local only, full privacy |
| `hybrid` | Better | Medium | NMT + alignment |
| `llm-align` | High | Slow | LLM quality + run-level formatting |
| `llm-plain` | High (fluent) | Slow | LLM only; loses inline formatting |

---

## Format Transplant

Apply the complete formatting of a blueprint `.docx` to the content of a source `.docx`. No translation — pure layout transplant.

**Blueprint** provides: page size, margins, section layout, every style definition (fonts, sizes, indents, spacing, colors), headers, footers, footnote formatting.

**Source** provides: all body text, bold/italic/underline of runs, tables, footnote text content.

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
  -v, --verbose         DEBUG logging (every XML element, every style resolution)
  --style-map SRC=BP    Explicit style overrides (repeatable)
```

**Examples:**

```bash
# Basic
python format_transplant.py template.docx manuscript.docx formatted.docx

# With verbose trace
python format_transplant.py template.docx manuscript.docx out.docx -v 2>&1 | tee run.log

# Debug style mapping
grep "\[MAPPER\]" run.log

# Explicit style overrides when auto-mapping misses
python format_transplant.py template.docx manuscript.docx out.docx \
    --style-map "My Body Text=Normal" \
    --style-map "Chapter Title=Heading 1" \
    --style-map "Blockzitat=Intense Quote"
```

### Transplant Web UI

```bash
python transplant-app.py
# → http://localhost:7861
```

Upload blueprint and source files, optionally enter style overrides (one `Source Style = Blueprint Style` pair per line), and click **Run**. The log shows a style mapping summary followed by the full pipeline trace.

### How the Transplant Works

The pipeline has four phases:

#### Phase 1 — Blueprint analysis

Every section's page geometry (size, margins, gutter, header/footer distance, orientation) and every style definition is extracted and indexed. Font properties are resolved by walking the style inheritance chain. This produces a `BlueprintSchema` used by all subsequent phases.

#### Phase 2 — Content extraction

Source paragraphs are extracted in body order alongside table placeholders (so table position within the flow is preserved). Each paragraph carries its raw `<w:p>` lxml element for deep-copy, plus semantic metadata (style name, heading level, bold/italic flags on runs, footnote reference flags).

Footnotes are extracted separately: each `<w:footnote>` element is deep-copied for later transplant.

#### Phase 3 — Style mapping

Every source style name is resolved to the best blueprint style name:

1. **User override** — explicit `--style-map` entry
2. **Exact name match** — identical style names across both documents
3. **Case-insensitive match** — handles `normal` vs `Normal`
4. **Semantic class** — heading level 1–9 detected across DE/FR/IT/ES/RU/ZH/PL/SE/EN, plus footnote text, captions, block quotes, abstracts
5. **Fallback** — blueprint's `Normal` style

The full mapping table is logged at `INFO` level under `[MAPPER]`.

#### Phase 4 — Document assembly

```
shutil.copy2(blueprint, output)   ← preserves styles.xml, settings, rels
↓
Clear body                        ← remove all <w:p> and <w:tbl>; keep final <w:sectPr>
↓
For each source element:
  paragraph → deep-copy <w:p> XML
              reset <w:pPr> → only mapped style reference (strips all direct formatting)
              clean <w:rPr> → keep bold/italic/underline, strip fonts/colors/sizes
  table     → deep-copy <w:tbl> XML, remap each cell paragraph's style
↓
Footnotes → remove blueprint's numbered footnotes
            insert source footnotes with blueprint's footnote text style
            commit updated footnotes.xml blob
↓
doc.save(output)
```

The `<w:pPr>` reset is the key operation: all direct paragraph formatting from the source (indents, spacing, alignment, section breaks) is discarded. Only the style reference remains, so the blueprint style governs the visual output completely.

#### Inline formatting

Run-level properties that carry semantic meaning are kept: `w:b` (bold), `w:i` (italic), `w:u` (underline), `w:strike`, `w:highlight`, `w:smallCaps`, `w:allCaps`, `w:vertAlign`, `w:vanish`.

Run-level properties that are purely aesthetic are stripped: `w:rFonts`, `w:sz`, `w:color`, `w:lang`, `w:kern`, `w:spacing` — the blueprint style defines all of these.

#### Debug log tags

Every log line is tagged for easy `grep`:

| Tag | What it covers |
|---|---|
| `[BLUEPRINT]` | Section geometry, every style attribute |
| `[EXTRACT]` | Every paragraph read from source (style, class, run count, text preview) |
| `[MAPPER]` | Every style resolution and the reason (exact/semantic/fallback) |
| `[BUILD]` | Every element inserted, every pPr reset, every rPr element stripped |

---

## License

GNU Affero General Public License v3.0 (AGPL-3.0). See [`LICENSE`](LICENSE).
