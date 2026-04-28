# Repository Guidelines

## Project Overview

This is a **self-learning research workspace** for industrial machine vision defect detection (电池划痕检测). The core domain problem: a battery factory has defect labels for 5 product models, but when a 6th model arrives with no scratch labels, direct transfer fails due to domain shift — product appearance subspace dominates defect semantic subspace.

Work is organized by date batches (`YYYY-MM-DD/`), each containing `sources/` (input PDFs) and `results/` (generated outputs: reports, HTML slides, PPTX).

## Project Structure & Module Organization

```
YYYY-MM-DD/
├── sources/                    # Input PDFs and reference material
└── results/
    ├── <arxiv-id>/             # Paper analysis outputs
    │   ├── report.md           # Structured analysis report
    │   └── images/             # Extracted figures from the paper
    └── html_ppt/               # Presentation deck
        ├── index.html          # HTML slide deck
        ├── style.css           # Scoped deck styles (.tpl-<name>-report)
        ├── build_rich_pptx.py  # PPTX generator script
        ├── <Name>.pptx         # Generated PowerPoint
        └── assets/             # Themes, runtime, fonts, animations (from html-ppt-skill)
```

Keep new scripts close to the artifact family they generate. Store reusable assets beside the output format they support.

## Conda Environment

All Python execution uses the `pdf_trans` conda environment:

```bash
conda run -n pdf_trans python <SCRIPT_PATH>
```

## Build, Test, and Development Commands

```bash
# Regenerate PPTX from HTML deck scripts
conda run -n pdf_trans python 2026-04-27/results/subspacead_ppt/scripts/build_offline_ppt.py
conda run -n pdf_trans python 2026-04-27/results/subspacead_ppt/scripts/build_enriched_ppt.py
conda run -n pdf_trans python 2026-04-27/results/html_ppt/build_rich_pptx.py
conda run -n pdf_trans python 2026-04-28/results/html_ppt/build_rich_pptx.py
```

When adding a new script, provide a `--output` argument if the artifact path may vary.

## Key Workflow: Paper → Report → PPT

### 1. Paper Analysis & Report

Reports are markdown files at `results/<paper-id>/report.md` with images in `results/<paper-id>/images/`. Structure: background, method, experiments, application to our task, recommendations by audience level.

### 2. HTML PPT Deck

Located at `results/html_ppt/`. Built using the `html-ppt-skill` skill:

- Copy `assets/` from `~/.claude/skills/html-ppt-skill/assets/`
- Mirror the style of the most recent existing deck (same `.tpl-*` scoping, component patterns)
- Include `<aside class="notes">` on every slide for speaker scripts (150–300 words each)
- Default theme: `corporate-clean`; press T in browser to cycle themes
- Keyboard: S = presenter mode, F = fullscreen, ← → = navigate
- Images referenced via relative paths to `../<paper-id>/images/`

### 3. PPTX Generation

Each `html_ppt/` contains a `build_rich_pptx.py` using `python-pptx` + `Pillow`:

- Slide size: 13.33" × 7.5" (widescreen)
- Color palette: `NAVY(8,35,62)`, `BLUE(29,78,216)`, `GREEN(15,118,110)`, `ORANGE(217,119,6)`, `RED(185,28,28)`, `MUTED(71,85,105)`, `LIGHT(244,247,251)`
- Helper functions: `add_header()`, `add_card()`, `add_panel()`, `add_table()`, `add_badge()`, `add_image_fit()`, `add_bullets()`
- When creating a new script, copy the most recent one and update: `base` image path, total slide count, slide content

## Coding Style & Naming Conventions

- 4-space indentation.
- Spaces around `=`, including keyword arguments and defaults.
- Import order: standard library, third-party, then local modules.
- `logger = logging.getLogger(__name__)` immediately after imports.
- English comments and docstrings with parameter descriptions.
- Prefer `Path` over raw string concatenation for file paths.
- Verb-led snake_case for generator scripts (e.g., `build_rich_pptx.py`).
- Artifact-type folder names: `images/`, `preview/`, `rendered/`.

## Testing Guidelines

No automated test suite. Validate by:

1. Running the PPTX generator script — check no tracebacks and expected file exists.
2. Opening `index.html` in browser — verify all slides render, images load, keyboard navigation works.
3. Opening `.pptx` — verify slide count, image placement, and text content.

## Commit & Pull Request Guidelines

Conventional Commits: `feat:`, `refactor:`, `fix:`, etc. Keep messages short and scoped to one artifact family or script change.

Pull requests should include:

- a short summary of what changed;
- the exact regeneration command used;
- affected output paths;
- screenshots or preview images for slide or layout changes.

Do not commit large intermediate files unless they are part of the intended deliverable.
