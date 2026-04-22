# PPTX Generator

[![PyPI version](https://img.shields.io/pypi/v/ai-pptx-generator)](https://pypi.org/project/ai-pptx-generator/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

English | [繁體中文](README.zh-TW.md)

Convert any document, report, or structured content into corporate-grade PPTX presentations.

**Recommended workflow:** Feed your content + the [prompt template](assets/prompt-template.md) to any AI → AI generates `slides.json` → tool produces PPTX. No manual slide data required.

Also works as an AI IDE Skill — just say "make me a presentation" and the AI handles everything.

---

## Installation

```bash
# From PyPI (recommended)
pip install ai-pptx-generator

# With YAML support
pip install "ai-pptx-generator[yaml]"

# From GitHub
pip install git+https://github.com/paul0728/pptx-generator.git

# Development mode
git clone https://github.com/paul0728/pptx-generator.git
cd pptx-generator
pip install -e ".[all]"
```

After installation, use `pptx-generate` command or `python -m pptx_generator`.

---

## Quick Start

`--input` supports three formats: `.json`, `.md`, `.yaml`.

### Option 1: AI-generated JSON (Recommended)

1. Paste [`assets/prompt-template.md`](assets/prompt-template.md) to any AI (ChatGPT, Claude, Gemini, etc.)
2. Provide your source content (documents, reports, notes)
3. AI generates `slides.json` — save and run:

```bash
pptx-generate --input slides.json --out output.pptx -v
```

The template includes the complete schema, all 11 slide type examples, and content limit rules. Example: [`assets/example-slides.json`](assets/example-slides.json)

### Option 2: Write Markdown manually

````markdown
---
title: Project Progress Report
version: 2026-Q2
---

# Project Progress Report
Q2 2026 Review

## Key Results
- Query latency reduced by 42%
  - p95 from 2.3s → 1.3s

## Connection Cache

```python
@lru_cache(maxsize=32)
def get_conn(dsn): ...
```
````

```bash
pptx-generate --input slides.md --out output.pptx -v
```

| Markdown Syntax | Slide Type |
|----------------|------------|
| `# H1` (first) | `title_slide` (cover) |
| `# H1` (subsequent) | `section_slide` (section divider) |
| `## H2` | New slide (`bullet_points` or `code_demo`) |
| Bullet lists `-` / `*` | `bullet_points` |
| Fenced code blocks | `code_demo` |
| `> Blockquote` | speaker notes |

> Markdown only supports basic types. For `kpi_slide`, `table`, `two_column` etc., use JSON or YAML.

Full example: [`assets/example-slides.md`](assets/example-slides.md)

### Option 3: YAML

More readable than JSON, supports flat format (no nested `content` required). Requires PyYAML.

```bash
pptx-generate --input slides.yaml --out output.pptx -v
```

Full example: [`assets/example-slides.yaml`](assets/example-slides.yaml)

### Option 4: Use in AI IDE

Install as a Skill, then say "Make a presentation from these meeting notes" — AI handles the entire pipeline. See [Use as AI Skill](#use-as-ai-skill).

---

## CLI Usage

```bash
# Basic (supports .json / .yaml / .md)
pptx-generate --input slides.md --out output.pptx -v

# Full parameters (use \ for line breaks on Linux/macOS; single line on Windows)
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
| `--input` | Slide data path `.json` / `.yaml` / `.md` (required) | — |
| `--template` | .pptx template path | Built-in `default-template.pptx` |
| `--out` | Output path | `output_presentation.pptx` |
| `--brand-color` | Brand color HEX | `#2B579A` |
| `--font` | Override font | Microsoft JhengHei / Calibri |
| `--footer` | Footer text | metadata.title |
| `--version-label` | Version label (bottom center) | metadata.version |
| `--watermark` | Watermark text | — |
| `--page-numbers` | Show page numbers | Off |
| `-v` / `-vv` | Verbose output | WARNING |

> `--json` still works (backward compatible), `--input` is recommended.

---

## Python API

```python
from pathlib import Path
from pptx_generator import generate, BrandConfig

brand = BrandConfig(color="#007A33", font="Noto Sans TC", footer="My Company", page_numbers=True)

output = generate(
    json_path=Path("slides.md"),       # supports .json / .yaml / .md
    template_path=None,                 # None = built-in template
    output_path=Path("output.pptx"),
    brand=brand,
)
```

```python
from pptx_generator import parse_markdown

data = parse_markdown("# Title\nSubtitle\n\n## Key Points\n- Point 1\n- Point 2")
print(data["slides"])  # pass to generate() or modify first
```

---

## Use as AI Skill

Once installed, the AI handles the entire flow from content analysis to PPTX generation (Phase 0–4).

### Install the Skill

Choose based on your AI agent:

**Kiro:**

```bash
git clone https://github.com/paul0728/pptx-generator.git .kiro/skills/pptx-generator
pip install python-pptx requests Pillow
```

**Claude Code / Cursor / Codex etc. (40+ agents):**

```bash
npx skills add paul0728/pptx-generator
pip install python-pptx requests Pillow
```

> [`npx skills`](https://github.com/vercel-labs/skills) installs to project-level by default (e.g. `.claude/skills/`). Add `-g` for global install.

```bash
npx skills list                            # List installed skills
npx skills update                          # Update all skills
npx skills remove pptx-generator           # Remove skill
```

### Directory structure after install

```
your-project/
├── .kiro/skills/pptx-generator/     ← Kiro
│   ├── SKILL.md
│   ├── assets/
│   └── pptx_generator/
│
├── .claude/skills/pptx-generator/   ← Claude Code (via npx skills)
│   └── ...
```

### Trigger phrases

| Language | Examples |
|----------|----------|
| Chinese | 「幫我做簡報」「把這份文件轉成投影片」 |
| English | "Make a PPT" "Convert this to slides" |
| Advanced | "Apply ./template.pptx, brand color #007A33, keep under 12 slides" |

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
Phase 1  Content analysis → slides.json (AI-generated / user-provided)
Phase 2  Mermaid diagram parallel rendering (cache + CJK fallback)
Phase 3  python-pptx assembly (layout lookup → auto-fit → brand chrome)
Phase 4  Quality verification
```

---

## Project Structure

```
pptx-generator/
├── pptx_generator/              # Python package
│   ├── generator.py             # Core logic
│   ├── markdown_parser.py       # Markdown → slides
│   └── assets/                  # Built-in template + examples
├── assets/                      # Examples & resources
│   ├── prompt-template.md       # AI prompt template (core)
│   ├── example-slides.*         # JSON / YAML / MD examples
│   └── default-template.pptx
├── SKILL.md                     # AI Skill definition
├── pyproject.toml
├── README.md                    # English
└── README.zh-TW.md             # 繁體中文
```

---

## License

MIT — see [LICENSE](LICENSE).
