# PPTX Generator

[![PyPI version](https://badge.fury.io/py/ai-pptx-generator.svg)](https://pypi.org/project/ai-pptx-generator/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

[繁體中文](README.md) | English

Convert any document, report, or structured content into corporate-grade PPTX presentations.

**Recommended workflow:** Feed your content + the [prompt template](assets/prompt-template.md) to any AI → AI generates `slides.json` → tool produces PPTX. No manual slide data required.

Also works as an AI IDE Skill — just say "make me a presentation" and the AI handles everything.

---

## Table of Contents

- [Features](#features)
- [Quick Start](#quick-start)
- [Installation](#installation)
- [CLI Usage](#cli-usage)
- [Input Formats](#input-formats)
- [Python API](#python-api)
- [Use as AI Skill](#use-as-ai-skill)
- [Slide Types](#slide-types)
- [Pipeline](#pipeline)
- [Project Structure](#project-structure)
- [License](#license)

---

## Features

- **AI-driven**: Includes a [prompt template](assets/prompt-template.md) that works with any AI to auto-generate slide data
- **3 input formats**: JSON (AI-generated / full control), Markdown (most intuitive for manual writing), YAML (human-readable)
- **11 slide types**: Title, Outline, Section, Bullets, Architecture Diagram, Code, Two-Column, Table, Image, KPI Cards, Ending
- **Mermaid diagram rendering**: Automatically converts Mermaid syntax to high-resolution PNG embedded in slides
- **CJK-aware auto-fit**: Mixed Chinese/English content auto-scales fonts to prevent overflow
- **Brand customization**: Brand color, custom fonts, footer, page numbers, watermark
- **Template support**: Apply custom `.pptx` templates with automatic layout detection
- **Quality verification**: Post-generation check ensures every slide has visible content

---

## Quick Start

### Option 1: AI-generated (Recommended)

The easiest approach — no manual slide data needed:

1. Paste the content of [`assets/prompt-template.md`](assets/prompt-template.md) to any AI (ChatGPT, Claude, Gemini, etc.)
2. Provide your source content (documents, reports, notes, verbal instructions)
3. AI generates a `slides.json`
4. Save and run:

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.json --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.json --out output.pptx -v
```

### Option 2: Use in AI IDE (Simplest)

After installing as an AI Skill, just say in the chat:

```
Make a presentation from these meeting notes
```

The AI handles the entire pipeline (outline → slides.json → PPTX).

### Option 3: Write Markdown manually

If you prefer full control, write slides in Markdown:

```bash
pptx-generate --input slides.md --out output.pptx -v
```

---

## Installation

### Option 1: From PyPI (Recommended)

**Windows (PowerShell):**
```powershell
pip install ai-pptx-generator
```

**Linux / macOS:**
```bash
pip install ai-pptx-generator
```

For YAML input support:

**Windows (PowerShell):**
```powershell
pip install "ai-pptx-generator[yaml]"
```

**Linux / macOS:**
```bash
pip install "ai-pptx-generator[yaml]"
```

### Option 2: From GitHub

**Windows (PowerShell):**
```powershell
pip install git+https://github.com/paul0728/pptx-generator.git
```

**Linux / macOS:**
```bash
pip install git+https://github.com/paul0728/pptx-generator.git
```

### Option 3: Development mode

**Windows (PowerShell):**
```powershell
git clone https://github.com/paul0728/pptx-generator.git
cd pptx-generator
pip install -e ".[all]"
```

**Linux / macOS:**
```bash
git clone https://github.com/paul0728/pptx-generator.git
cd pptx-generator
pip install -e ".[all]"
```

After installation, use `pptx-generate` command or `python -m pptx_generator`.

---

## CLI Usage

### Basic — Generate from Markdown

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.md --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.md --out output.pptx -v
```

### Generate from YAML

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.yaml --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.yaml --out output.pptx -v
```

### Generate from JSON

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.json --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.json --out output.pptx -v
```

> `--json slides.json` still works (backward compatible), but `--input` is recommended.

### Full parameters

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.yaml --template my-template.pptx --out report.pptx --brand-color "#007A33" --font "Noto Sans TC" --footer "Company · Confidential" --version-label "2026-Q2 v1.2" --watermark "DRAFT" --page-numbers -v
```

**Linux / macOS:**
```bash
pptx-generate \
    --input slides.yaml \
    --template my-template.pptx \
    --out report.pptx \
    --brand-color "#007A33" \
    --font "Noto Sans TC" \
    --footer "Company · Confidential" \
    --version-label "2026-Q2 v1.2" \
    --watermark "DRAFT" \
    --page-numbers \
    -v
```

| Parameter | Description | Default |
|-----------|-------------|---------|
| `--input` | Slide data path (`.json` / `.yaml` / `.yml` / `.md`) (required) | — |
| `--template` | .pptx template path | Built-in `default-template.pptx` |
| `--out` | Output file path | `output_presentation.pptx` |
| `--brand-color` | Brand color HEX | `#2B579A` |
| `--font` | Override font | Microsoft JhengHei / Calibri |
| `--footer` | Footer text | metadata.title |
| `--version-label` | Version label (bottom center) | metadata.version |
| `--watermark` | Watermark text | — |
| `--page-numbers` | Show page numbers | Off |
| `-v` / `-vv` | Verbose output (INFO / DEBUG) | WARNING |

---

## Input Formats

### JSON (Recommended — AI-generated)

Use the [prompt template (prompt-template.md)](assets/prompt-template.md) with any AI to auto-generate a valid `slides.json`.

The template includes the complete schema, examples for all 11 slide types, and content limit rules. Just provide your raw content and the AI produces ready-to-use JSON.

Full example: [`assets/example-slides.json`](assets/example-slides.json)

### Markdown (Most intuitive for manual writing)

Write slide content in Markdown — the tool auto-converts to presentation structure. No special format to learn.

````markdown
---
title: Project Progress Report
version: 2026-Q2
---

# Project Progress Report
Q2 2026 Review

## Outline
- Background
- Architecture
- Key Results

# Architecture

## Key Results
- Query latency reduced by 42%
  - p95 from 2.3s → 1.3s
- Error rate down to 0.4%

## Connection Cache

```python
@lru_cache(maxsize=32)
def get_conn(dsn): ...
```

## Thank You
Q & A
````

**Markdown mapping rules:**

| Markdown Syntax | Slide Type |
|----------------|------------|
| `# H1` (first) | `title_slide` (cover) |
| `# H1` (subsequent) | `section_slide` (section divider) |
| `## H2` | New slide (`bullet_points` or `code_demo`) |
| Bullet lists `-` / `*` | `bullet_points` points |
| Fenced code blocks | `code_demo` |
| `> Blockquote` | speaker notes |
| YAML front matter | `presentation_metadata` |

Full example: [`assets/example-slides.md`](assets/example-slides.md)

### YAML (Human-readable)

More readable than JSON, supports flat format (no nested `content` required):

```yaml
title: Project Progress Report
version: 2026-Q2

slides:
  - type: title_slide
    title: Project Progress Report
    sub_title: "Q2 2026 Review\n2026-04-21"

  - type: bullet_points
    title: Key Results
    points:
      - Query latency reduced by 42%
      - Error rate down to 0.4%

  - type: code_demo
    title: Connection Cache
    language: python
    code: |
      @lru_cache(maxsize=32)
      def get_conn(dsn: str) -> Connection:
          return psycopg2.connect(dsn)
```

> YAML input requires PyYAML: `pip install pyyaml` or `pip install "ai-pptx-generator[yaml]"`

Full example: [`assets/example-slides.yaml`](assets/example-slides.yaml)

---

## Python API

```python
from pathlib import Path
from pptx_generator import generate, BrandConfig

brand = BrandConfig(
    color="#007A33",
    font="Noto Sans TC",
    footer="My Company",
    page_numbers=True,
)

# Supports .json / .yaml / .md input
output = generate(
    json_path=Path("slides.md"),      # or slides.yaml / slides.json
    template_path=None,                # use built-in template
    output_path=Path("output.pptx"),
    brand=brand,
)
print(f"Presentation generated: {output}")
```

### Markdown Parsing API

```python
from pptx_generator import parse_markdown

data = parse_markdown("""
# My Presentation
Subtitle here

## Key Points
- Point 1
- Point 2
""")

print(data["slides"])  # pass to generate() or modify before generating
```

---

## Use as AI Skill

This tool can be used as an AI IDE Skill. Once installed, the AI handles the entire flow from content analysis to PPTX generation.

### Install via npx (Recommended)

One command to download the skill into your project:

**Windows (PowerShell):**
```powershell
npx degit paul0728/pptx-generator .skills/pptx-generator
```

**Linux / macOS:**
```bash
npx degit paul0728/pptx-generator .skills/pptx-generator
```

Then install Python dependencies:

```bash
pip install python-pptx requests Pillow
```

### Install via git clone

**Windows (PowerShell):**
```powershell
git clone https://github.com/paul0728/pptx-generator.git .skills/pptx-generator
```

**Linux / macOS:**
```bash
git clone https://github.com/paul0728/pptx-generator.git .skills/pptx-generator
```

### Expected directory structure

```
your-project/
├── .skills/
│   └── pptx-generator/
│       ├── SKILL.md                 ← AI reads this to understand capabilities
│       ├── assets/
│       │   ├── default-template.pptx
│       │   ├── prompt-template.md   ← Prompt template for AI-driven generation
│       │   ├── example-slides.json
│       │   ├── example-slides.yaml
│       │   └── example-slides.md
│       ├── scripts/
│       │   └── generate_pptx_template.py
│       └── pptx_generator/
│           └── ...
├── src/
└── ...
```

### Usage in AI IDE

Just say in the chat:

```
Make a presentation from these meeting notes
```

The AI will automatically:
1. Detect your intent → activate pptx-generator Skill
2. Analyze content and plan outline (Phase 0)
3. Generate structured slide data (Phase 1)
4. Render Mermaid diagrams (Phase 2)
5. Assemble PPTX file (Phase 3)
6. Quality verification (Phase 4)

### Trigger phrases

| Language | Examples |
|----------|----------|
| Chinese | 「幫我做簡報」「把這份文件轉成投影片」「做一份進度報告 PPT」 |
| English | "Make a PPT" "Convert this to slides" "Generate a slide deck" |
| With template | "Apply ./template.pptx and make slides" |
| With style | "KPI dashboard style, brand color #007A33" |
| With page limit | "Keep it under 12 slides" |
| With audience | "Executive summary for the boss" |

---

## Slide Types

| Type | Purpose | Content Fields |
|------|---------|---------------|
| `title_slide` | Cover (slide 1) | `title`, `sub_title` |
| `outline_slide` | Table of contents (slide 2) | `title`, `points[]` |
| `section_slide` | Section divider | `title`, `sub_title` |
| `bullet_points` | Bullet list | `title`, `points[]` |
| `architecture_diagram` | Mermaid diagram | `title`, `mermaid_code`, `description` |
| `code_demo` | Code display | `title`, `code`, `language` |
| `two_column` | Side-by-side comparison | `title`, `left{heading, points}`, `right{heading, points}` |
| `table` | Data table | `title`, `headers[]`, `rows[][]` |
| `image_slide` | Image + caption | `title`, `image_path`, `caption` |
| `kpi_slide` | KPI cards (≤6) | `title`, `kpis[{label, value, unit, delta}]` |
| `ending_slide` | Closing slide | `title`, `sub_title` |

---

## Pipeline

```
Phase 0  Outline planning (AI auto-executes using prompt-template.md)
Phase 1  Content analysis → slides.json (AI-generated, or user-provided JSON / YAML / Markdown)
Phase 2  Mermaid diagram parallel rendering (with cache & CJK fallback)
Phase 3  python-pptx assembly (layout lookup → auto-fit → brand chrome)
Phase 4  Quality verification
```

---

## Project Structure

```
pptx-generator/
├── pptx_generator/                  # Python package (pip install)
│   ├── __init__.py
│   ├── __main__.py                  # python -m pptx_generator
│   ├── generator.py                 # Core logic
│   ├── markdown_parser.py           # Markdown → slides conversion
│   └── assets/
│       ├── default-template.pptx
│       └── example-slides.json
├── scripts/
│   └── generate_pptx_template.py    # Skill entry point (backward compat)
├── assets/                           # Examples & resources
│   ├── default-template.pptx
│   ├── example-output.pptx
│   ├── example-slides.json
│   ├── example-slides.yaml
│   ├── example-slides.md
│   └── prompt-template.md            # AI prompt template (core)
├── SKILL.md                          # AI Skill definition
├── pyproject.toml                    # PyPI package config
├── requirements.txt
├── MANIFEST.in
├── LICENSE
├── README.md                         # 繁體中文
└── README.en.md                      # English
```

---

## License

MIT License — see [LICENSE](LICENSE).
