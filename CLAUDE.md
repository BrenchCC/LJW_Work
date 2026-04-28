# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This repo is a **self-learning research workspace** for industrial machine vision defect detection (电池划痕检测). Each work batch is organized by date (`YYYY-MM-DD/`), containing `sources/` (PDF papers) and `results/` (generated outputs). The core workflow is: read papers → write analysis reports → build presentation decks (HTML + PPTX).

The domain problem: a battery factory has defect labels for 5 product models, but when a 6th model arrives with no scratch labels, direct transfer fails due to domain shift. The fundamental issue is that **product appearance subspace dominates defect semantic subspace**.

## Conda Environment

All Python execution uses the `pdf_trans` conda environment:

```bash
conda run -n pdf_trans python <SCRIPT_PATH>
```

## Key Workflow: Paper → Report → PPT

### 1. Paper Analysis & Report Generation

Reports are markdown files at `results/<paper-id>/report.md` with images in `results/<paper-id>/images/`. Reports follow a consistent structure: background, method, experiments, application to our task, and recommendations by audience level.

### 2. HTML PPT Deck

Located at `results/html_ppt/`. Built using the `html-ppt-skill` skill:

- `index.html` — Main slide deck (scoped CSS class `.tpl-<name>-report`)
- `style.css` — Deck-specific styles (grid layouts, cards, pipelines, formula cards)
- `assets/` — Copied from `~/.claude/skills/html-ppt-skill/assets/` (themes, runtime.js, fonts, animations)
- Images referenced via relative paths to `../<paper-id>/images/`

When creating a new HTML deck:
1. Copy `assets/` from the skill directory
2. Mirror the style of the most recent existing deck (e.g., same `.tpl-*` scoping, component patterns)
3. Include `<aside class="notes">` on every slide for speaker scripts (150–300 words each)
4. Use `corporate-clean` as default theme; press T in browser to cycle themes
5. Keyboard: S = presenter mode, F = fullscreen, ← → = navigate

### 3. PPTX Generation

Each `html_ppt/` directory contains a `build_rich_pptx.py` script that generates a matching PowerPoint file:

```bash
conda run -n pdf_trans python 2026-04-28/results/html_ppt/build_rich_pptx.py
```

The script uses `python-pptx` + `Pillow` with a consistent visual system:
- Slide size: 13.33" × 7.5" (widescreen)
- Color palette: `NAVY(8,35,62)`, `BLUE(29,78,216)`, `GREEN(15,118,110)`, `ORANGE(217,119,6)`, `RED(185,28,28)`, `MUTED(71,85,105)`, `LIGHT(244,247,251)`
- Helper functions: `add_header()`, `add_card()`, `add_panel()`, `add_table()`, `add_badge()`, `add_image_fit()`, `add_bullets()`
- Image paths reference `results/<paper-id>/images/` relative to repo root

When creating a new `build_rich_pptx.py`, copy the structure from the most recent one and update: `base` image path, total slide count, and all slide content.

## Repository Structure Pattern

```
YYYY-MM-DD/
├── sources/                    # Input PDFs and reference material
└── results/
    ├── <arxiv-id>/             # Paper analysis outputs
    │   ├── report.md           # Structured analysis report
    │   └── images/             # Extracted figures from the paper
    └── html_ppt/               # Presentation deck
        ├── index.html          # HTML slide deck
        ├── style.css           # Scoped deck styles
        ├── build_rich_pptx.py  # PPTX generator script
        ├── <Name>.pptx         # Generated PowerPoint
        └── assets/             # Themes, runtime, fonts, animations
```

## Coding Conventions

Follow the existing deck builder style:
- 4-space indentation, spaces around `=` (including keyword args)
- Import order: stdlib → third-party → local
- `logger = logging.getLogger(__name__)` after imports
- English comments and docstrings with parameter descriptions
- `Path` over raw string concatenation for file paths
- Verb-led snake_case for generator scripts (e.g., `build_rich_pptx.py`)
- Artifact-type folder names: `images/`, `preview/`, `rendered/`

## Validation

No automated test suite. Validate by:
1. Running the PPTX generator script — check no tracebacks and expected file exists
2. Opening `index.html` in browser — verify all slides render, images load, keyboard navigation works
3. Opening `.pptx` — verify slide count, image placement, and text content

## Commit Style

Conventional Commits: `feat:`, `refactor:`, `fix:`, etc. Keep messages short and scoped to one artifact family.
