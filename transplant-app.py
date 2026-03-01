#!/usr/bin/env python3
"""
Gradio interface for Format Transplant
Designed for local use and Hugging Face Spaces deployment.
"""

import logging
import os
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple, List

import gradio as gr

# ============================================================================
# LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)-7s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ============================================================================
# IMPORT CORE ENGINE
# ============================================================================

try:
    from format_transplant import (
        FormatTransplanter,
        LLMFormatTransplanter,
        PROVIDER_DEFAULTS,
        llm_config_from_args,
    )
    ENGINE_OK = True
    ENGINE_ERROR = None
except Exception as _e:
    ENGINE_OK = False
    ENGINE_ERROR = str(_e)
    logger.error("Failed to import format_transplant: %s", _e)

# ============================================================================
# ENVIRONMENT STATUS
# ============================================================================

def _check_environment() -> str:
    lines = []

    def _ok(msg):  lines.append(f"✓ {msg}")
    def _err(msg): lines.append(f"✗ {msg}")
    def _inf(msg): lines.append(f"ℹ {msg}")

    # python-docx
    try:
        from docx import Document
        _ok("python-docx installed")
    except ImportError:
        _err("python-docx missing – run: pip install python-docx")

    # lxml
    try:
        from lxml import etree
        _ok("lxml installed")
    except ImportError:
        _err("lxml missing – run: pip install lxml")

    # Core engine
    if ENGINE_OK:
        _ok("format_transplant engine loaded")
    else:
        _err(f"format_transplant engine failed to load: {ENGINE_ERROR}")

    lines.append("")
    lines.append("── LLM providers ──")

    # openai (covers OpenAI / Nebius / Scaleway / OpenRouter / Mistral)
    try:
        import openai
        _ok(f"openai SDK {openai.__version__}  (covers OpenAI, Nebius, Scaleway, OpenRouter, Mistral)")
    except ImportError:
        _inf("openai SDK missing – run: pip install openai  (needed for OpenAI/Nebius/Scaleway/OpenRouter/Mistral)")

    # anthropic
    try:
        import anthropic
        _ok(f"anthropic SDK {anthropic.__version__}")
    except ImportError:
        _inf("anthropic SDK missing – run: pip install anthropic")

    # fastapi-poe
    try:
        import fastapi_poe
        _ok("fastapi-poe installed")
    except ImportError:
        _inf("fastapi-poe missing – run: pip install fastapi-poe  (needed for Poe)")

    lines.append("")
    lines.append("── Detected API keys ──")
    for provider, defaults in (PROVIDER_DEFAULTS.items() if ENGINE_OK else {}.items()):
        env = defaults.get("env", "")
        if env and os.getenv(env):
            _ok(f"{provider.capitalize()} key found  ({env})")

    return "\n".join(lines)


SETUP_STATUS = _check_environment()
logger.info("Environment status:\n%s", SETUP_STATUS)

# ============================================================================
# CORE PROCESSING FUNCTION
# ============================================================================

def run_transplant(
    blueprint_file: Optional[str],
    source_file: Optional[str],
    style_overrides_text: str,
    verbose: bool,
    # LLM parameters
    llm_provider: str,
    llm_model: str,
    llm_api_key: str,
    llm_mode: str,
    styleguide_in_file: Optional[str],
    extra_styleguide_files,          # list[str] | None from gr.File(file_count="multiple")
    llm_batch_size: int,
    llm_context_chars: int,
    progress=gr.Progress(),
) -> Tuple[Optional[str], Optional[str], str]:
    """
    Main handler called by the Gradio button.

    Returns (output_docx_path_or_None, styleguide_path_or_None, log_string).
    """

    use_llm = bool(llm_provider and llm_provider != "(none)")

    # ── Validate inputs ───────────────────────────────────────────────
    if blueprint_file is None:
        return None, None, "❌ No blueprint file uploaded."
    if source_file is None and llm_mode != "styleguide_only":
        return None, None, "❌ No source file uploaded."

    blueprint_path = Path(blueprint_file)
    source_path    = Path(source_file) if source_file else None

    if blueprint_path.suffix.lower() != ".docx":
        return None, None, "❌ Blueprint must be a .docx file."
    if source_path and source_path.suffix.lower() != ".docx":
        return None, None, "❌ Source must be a .docx file."

    if not ENGINE_OK:
        return None, None, f"❌ Engine not available: {ENGINE_ERROR}"

    # ── Parse style overrides ─────────────────────────────────────────
    overrides = {}
    if style_overrides_text.strip():
        for raw_line in style_overrides_text.strip().splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                logger.warning("Ignored override line (no '='): '%s'", line)
                continue
            src, _, bp = line.partition("=")
            overrides[src.strip()] = bp.strip()

    # ── Output paths ───────────────────────────────────────────────────
    temp_dir = Path(tempfile.mkdtemp())
    output_filename = f"{source_path.stem}_transplanted.docx" if source_path else "transplanted.docx"
    output_path = temp_dir / output_filename
    styleguide_out_path = temp_dir / "styleguide.md"

    # ── Log capture ───────────────────────────────────────────────────
    log_records: list = []

    class _Capture(logging.Handler):
        def emit(self, record):
            log_records.append(self.format(record))

    capture_handler = _Capture()
    capture_handler.setFormatter(logging.Formatter("[%(levelname)-7s] %(message)s"))

    root_log = logging.getLogger()
    saved_level = root_log.level
    root_log.addHandler(capture_handler)
    root_log.setLevel(logging.DEBUG if verbose else logging.INFO)

    saved_sg_path: Optional[str] = None

    try:
        progress(0.05, desc="Checking files…")

        if use_llm:
            # ── Build LLM config ──────────────────────────────────────
            progress(0.10, desc="Initialising LLM client…")

            llm_cfg = llm_config_from_args(
                provider_str=llm_provider,
                model=llm_model.strip() or None,
                api_key=llm_api_key.strip() or None,
            )
            llm_cfg.para_batch_size        = int(llm_batch_size)
            llm_cfg.blueprint_context_chars = int(llm_context_chars)

            extra_sg_paths: Optional[List[Path]] = None
            if extra_styleguide_files:
                files = extra_styleguide_files if isinstance(extra_styleguide_files, list) else [extra_styleguide_files]
                extra_sg_paths = [Path(f) for f in files if f]

            sg_in: Optional[Path] = None
            if styleguide_in_file:
                sg_in = Path(styleguide_in_file)

            transplanter = LLMFormatTransplanter()

            progress(0.15, desc="Phase 1 – Analysing blueprint…")
            progress(0.25, desc="Phase 1-LLM – Generating style guide…")
            progress(0.45, desc="Phase 2 – Extracting & LLM-formatting content…")
            progress(0.70, desc="Phase 3-4 – Style mapping & document assembly…")

            saved_sg = transplanter.run(
                blueprint_path=blueprint_path,
                source_path=source_path,
                output_path=output_path if llm_mode != "styleguide_only" else temp_dir / "_unused.docx",
                llm_config=llm_cfg,
                extra_styleguide_paths=extra_sg_paths,
                styleguide_in=sg_in,
                styleguide_out=styleguide_out_path,
                llm_mode=llm_mode,
                user_style_overrides=overrides or None,
            )

            if saved_sg and saved_sg.exists():
                saved_sg_path = str(saved_sg)

        else:
            # ── Structural transplant only (no LLM) ───────────────────
            transplanter = FormatTransplanter()

            progress(0.15, desc="Phase 1 – Analysing blueprint…")
            progress(0.30, desc="Phase 2 – Extracting source content…")
            progress(0.55, desc="Phase 3 – Mapping styles…")
            progress(0.70, desc="Phase 4 – Building output document…")

            transplanter.run(
                blueprint_path=blueprint_path,
                source_path=source_path,
                output_path=output_path,
                user_style_overrides=overrides or None,
            )

        progress(1.0, desc="✓ Complete!")

    except Exception as exc:
        root_log.removeHandler(capture_handler)
        root_log.setLevel(saved_level)
        log_text = "\n".join(log_records)
        logger.error("Transplant failed: %s", exc, exc_info=True)
        return None, None, (
            f"❌ Error: {exc}\n\n"
            f"── Log before error ──\n{log_text}"
        )

    root_log.removeHandler(capture_handler)
    root_log.setLevel(saved_level)

    # ── Build summary ─────────────────────────────────────────────────
    log_text = "\n".join(log_records)

    mapper_lines = [l for l in log_records if "[MAPPER]" in l and "→" in l]
    mapper_summary = "\n".join(mapper_lines) if mapper_lines else "(none)"

    llm_lines = [l for l in log_records if "Phase 1-LLM" in l or "Phase 2-LLM" in l]
    llm_summary = "\n".join(llm_lines) if llm_lines else ""

    out_filename = output_filename if llm_mode != "styleguide_only" else "(none – styleguide_only mode)"
    out_path_for_return = str(output_path) if (llm_mode != "styleguide_only" and output_path.exists()) else None

    summary_parts = [
        "✅ Format Transplant Complete!\n",
        f"📋 Blueprint : {blueprint_path.name}",
        f"📄 Source    : {source_path.name if source_path else '(none)'}",
        f"📤 Output    : {out_filename}",
        f"🎨 Overrides : {len(overrides)}",
    ]
    if use_llm:
        summary_parts += [
            f"🤖 LLM       : {llm_provider} / {llm_cfg.model}",
            f"🔧 Mode      : {llm_mode}",
        ]
    if llm_summary:
        summary_parts += ["\n── LLM phases ──", llm_summary]
    summary_parts += [
        "\n── Style mapping ──",
        mapper_summary,
        "\n── Full log ──",
        log_text,
    ]

    return out_path_for_return, saved_sg_path, "\n".join(summary_parts)


# ============================================================================
# GRADIO INTERFACE
# ============================================================================

_PROVIDER_CHOICES = ["(none)"] + list(PROVIDER_DEFAULTS.keys()) if ENGINE_OK else ["(none)"]


def _default_model_for_provider(provider: str) -> str:
    """Return the default model string for a provider name."""
    if not ENGINE_OK or provider == "(none)":
        return ""
    return PROVIDER_DEFAULTS.get(provider, {}).get("model", "")


def create_interface() -> gr.Blocks:

    with gr.Blocks(title="Format Transplant") as demo:

        gr.Markdown("""
# 🎨 Format Transplant

Apply the **complete formatting** of a blueprint document to the **content** of a source document — down to paragraph styles, page layout, margins, headers, footers, and footnotes.
Optionally run an **LLM style pass** that learns editorial conventions from the blueprint and re-formats source paragraphs and footnotes accordingly.

| What comes from the **blueprint** | What comes from the **source** |
|---|---|
| Page size, margins, section layout | All body text |
| All style definitions (fonts, indents, spacing) | Bold / italic / underline of runs |
| Headers & footers | Tables (with remapped styles) |
| Footnote formatting | Footnote text content |
""")

        with gr.Row():

            # ── Left column: inputs ────────────────────────────────────
            with gr.Column(scale=1):
                gr.Markdown("### 📋 Input files")

                blueprint_file = gr.File(
                    label="① Blueprint DOCX  (provides formatting)",
                    file_types=[".docx"],
                    type="filepath",
                )

                source_file = gr.File(
                    label="② Source DOCX  (provides content)",
                    file_types=[".docx"],
                    type="filepath",
                )

                gr.Markdown("### ⚙️ Options")

                style_overrides = gr.Textbox(
                    label="Style overrides  (optional)",
                    placeholder=(
                        "One mapping per line:\n"
                        "  Source Style Name = Blueprint Style Name\n\n"
                        "Examples:\n"
                        "  My Custom Body = Normal\n"
                        "  Big Header = Heading 1\n"
                        "  Zitat = Intense Quote"
                    ),
                    lines=5,
                    info=(
                        "Leave blank for automatic mapping. "
                        "Check the log for [MAPPER] lines to audit what was resolved."
                    ),
                )

                verbose = gr.Checkbox(
                    label="Verbose debug logging",
                    value=False,
                    info="Shows every XML element stripped/reset — helpful for troubleshooting.",
                )

                run_btn = gr.Button(
                    "🚀 Run Format Transplant",
                    variant="primary",
                    size="lg",
                )

            # ── Right column: outputs ──────────────────────────────────
            with gr.Column(scale=1):
                gr.Markdown("### 📥 Result")

                output_file = gr.File(
                    label="Transplanted document  (.docx)",
                    interactive=False,
                )

                styleguide_file = gr.File(
                    label="Generated style guide  (.md)  — only produced when LLM is enabled",
                    interactive=False,
                )

                log_output = gr.Textbox(
                    label="Log",
                    lines=24,
                    max_lines=60,
                    interactive=False,
                )

        # ── LLM accordion ──────────────────────────────────────────────
        with gr.Accordion("🤖 LLM Style Pass  (optional)", open=False):
            gr.Markdown("""
Select an LLM provider to add an **editorial style pass** on top of the structural format transplant:

1. **Style guide generation** — the LLM reads a sample of the blueprint and produces a `styleguide.md`
   describing conventions like: _names always italic, DMG transliteration for Arabic, citation markers,
   foreign terms in italics, quotation-mark style_, etc.
2. **Content formatting** — each source paragraph / footnote is sent (in batches) to the LLM together
   with the style guide.  The LLM returns Markdown with bold/italic applied; these runs replace the
   original runs in the transplanted document.

Leave **Provider** at `(none)` to skip the LLM pass entirely (fast, structural-only transplant).
""")

            with gr.Row():
                llm_provider = gr.Dropdown(
                    label="Provider",
                    choices=_PROVIDER_CHOICES,
                    value="(none)",
                    info="Select an LLM provider. API key must be available.",
                )
                llm_model = gr.Dropdown(
                    label="Model",
                    choices=["auto"],
                    value="auto",
                    allow_custom_value=True,
                    info="Select a model or type a custom one. Use 'Fetch Models' to update list.",
                )
                fetch_models_btn = gr.Button("🔄 Fetch Models", size="sm")
                llm_api_key = gr.Textbox(
                    label="API key  (optional)",
                    type="password",
                    placeholder="sk-…",
                    info="Overrides env variable if provided.",
                )

            # --- Logic to fetch models ---
            def _fetch_models(provider, api_key):
                if provider == "(none)":
                    return gr.update(choices=["auto"], value="auto")
                
                try:
                    # Temporary config to use for fetching
                    cfg = llm_config_from_args(provider, api_key=api_key)
                    client = MultiProviderLLMClient()
                    models = client.get_available_models(cfg)
                    
                    if not models:
                        return gr.update(choices=["auto"], value="auto")
                    
                    choices = [m["id"] for m in models]
                    # Also include the default model from PROVIDER_DEFAULTS if not in list
                    default_m = PROVIDER_DEFAULTS.get(provider, {}).get("model")
                    if default_m and default_m not in choices:
                        choices.insert(0, default_m)
                    
                    return gr.update(choices=choices, value=choices[0])
                except Exception as e:
                    logger.error(f"Fetch models failed: {e}")
                    return gr.update(choices=["auto", f"Error: {str(e)[:20]}..."], value="auto")

            fetch_models_btn.click(
                fn=_fetch_models,
                inputs=[llm_provider, llm_api_key],
                outputs=[llm_model]
            )

            with gr.Row():
                llm_mode = gr.Radio(
                    label="LLM mode",
                    choices=["both", "paragraphs", "footnotes", "styleguide_only"],
                    value="both",
                    info=(
                        "both — format paragraphs and footnotes  |  "
                        "styleguide_only — only generate the style guide, no output document"
                    ),
                )

            with gr.Row():
                llm_batch_size = gr.Slider(
                    label="Batch size  (paragraphs per LLM call)",
                    minimum=1,
                    maximum=50,
                    step=1,
                    value=15,
                    info="Smaller = more calls, larger = may hit context limits.",
                )
                llm_context_chars = gr.Slider(
                    label="Blueprint context  (chars sent for style guide generation)",
                    minimum=5_000,
                    maximum=120_000,
                    step=5_000,
                    value=40_000,
                    info="~4 chars ≈ 1 token. Adjust to fit your model's context window.",
                )

            with gr.Row():
                styleguide_in = gr.File(
                    label="Pre-existing style guide  (.md)  — skip generation if provided",
                    file_types=[".md", ".txt"],
                    type="filepath",
                )
                extra_styleguides = gr.File(
                    label="Extra style guide files  (optional, multiple)",
                    file_types=[".md", ".txt", ".pdf"],
                    type="filepath",
                    file_count="multiple",
                )

            # Auto-fill default model when provider changes
            def _on_provider_change(provider):
                return _default_model_for_provider(provider)

            llm_provider.change(
                fn=_on_provider_change,
                inputs=[llm_provider],
                outputs=[llm_model],
            )

        # ── System status ──────────────────────────────────────────────
        with gr.Accordion("System status", open=False):
            gr.Markdown(f"```\n{SETUP_STATUS}\n```")

        # ── Help / docs ────────────────────────────────────────────────
        with gr.Accordion("How it works", open=False):
            gr.Markdown("""
### Structural pipeline (always runs)

1. **Blueprint analysis** — reads every section (margins, page size, header/footer distance) and
   every style definition (font, size, bold, italic, indents) from the blueprint, resolving the
   full style inheritance chain.

2. **Content extraction** — pulls all body paragraphs and tables from the source in order,
   capturing text and semantic inline formatting (bold/italic/underline).
   Footnote content is also extracted.

3. **Style mapping** — each source paragraph style is mapped to the best blueprint style:
   exact name → case-insensitive → semantic class → fallback to Normal.
   Semantic classes include headings 1–9 (detected in DE/FR/IT/ES/RU/ZH/PL/SE/EN),
   footnote text, captions, block-quotes, abstracts.

4. **Document assembly** — a copy of the blueprint becomes the output.
   Its body is cleared (the final `<w:sectPr>` that holds page layout is kept).
   Source elements are inserted one by one with the mapped style applied.
   Footnotes are transplanted with the blueprint's footnote text style.

### LLM style pass (when provider ≠ none)

**Phase 1-LLM** — the LLM receives a sample of the blueprint text and produces a `styleguide.md`
with self-instructions about editorial conventions.  You can skip this by uploading a pre-existing
style guide (the file input above), or you can run `styleguide_only` mode first, inspect the result,
edit it, then re-run with the edited file as input.

**Phase 2-LLM** — source paragraphs and footnotes are sent in batches to the LLM together with the
style guide. The LLM returns `***bold+italic***` / `**bold**` / `*italic*` Markdown. These inline
markers are parsed and the resulting runs replace the original content in the transplanted document.

### Style override syntax

One mapping per line in the **Style overrides** box:

```
Source Style Name = Blueprint Style Name
My Custom Body    = Normal
Big Chapter Head  = Heading 1
Blockzitat        = Intense Quote
```

### Debug tips

- Enable **Verbose** logging and read the full log.
- `[MAPPER]` lines show every style resolution and why.
- `[BUILD]` lines show every element inserted into the output.
- `[BLUEPRINT]` lines show what was read from the blueprint.
- `[EXTRACT]` lines show what was read from the source.

If a style isn't mapping correctly, copy its exact name from a `[EXTRACT]` line
and add an override.
""")

        # ── Wire button ────────────────────────────────────────────────
        run_btn.click(
            fn=run_transplant,
            inputs=[
                blueprint_file,
                source_file,
                style_overrides,
                verbose,
                llm_provider,
                llm_model,
                llm_api_key,
                llm_mode,
                styleguide_in,
                extra_styleguides,
                llm_batch_size,
                llm_context_chars,
            ],
            outputs=[output_file, styleguide_file, log_output],
        )

    return demo


# ============================================================================
# MAIN
# ============================================================================

if __name__ == "__main__":
    demo = create_interface()
    # Respect the GRADIO_SERVER_PORT environment variable if set (standard for HF Spaces)
    server_port = int(os.getenv("GRADIO_SERVER_PORT", 7860))
    demo.launch(
        server_name="0.0.0.0",
        server_port=server_port,
        share=False,
    )
