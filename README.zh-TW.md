# PPTX Generator

[![PyPI version](https://img.shields.io/pypi/v/ai-pptx-generator)](https://pypi.org/project/ai-pptx-generator/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

繁體中文 | [English](README.md)

將任何文件、報告或結構化內容轉換為企業級 PPTX 簡報。

**推薦用法：** 把你的內容 + [提示詞樣板](pptx_generator/assets/prompt-template.md) 餵給 AI → AI 產出 `slides.json` → 工具生成 PPTX。不需要手寫任何投影片資料。

也可作為 AI IDE Skill，直接說「幫我做簡報」就能自動完成全部流程。

---

## 安裝

```bash
# 從 PyPI（推薦）
pip install ai-pptx-generator

# 含 YAML 支援
pip install "ai-pptx-generator[yaml]"

# 從 GitHub
pip install git+https://github.com/paul0728/pptx-generator.git

# 開發模式
git clone https://github.com/paul0728/pptx-generator.git
cd pptx-generator
pip install -e ".[all]"
```

安裝後可使用 `pptx-generate` 指令或 `python -m pptx_generator`。

---

## 快速開始

`--input` 支援三種格式：`.json`、`.md`、`.yaml`。

### 方式一：AI 自動產出 JSON（推薦）

1. 將 [`pptx_generator/assets/prompt-template.md`](pptx_generator/assets/prompt-template.md) 貼給任何 AI（ChatGPT、Claude、Gemini 等）
2. 提供你的需求描述或原始內容
3. AI 產出 `slides.json`，存檔後執行：

```bash
pptx-generate --input slides.json --out output.pptx -v
```

樣板包含完整 schema、11 種投影片類型範例、內容量限制規則。範例：[`pptx_generator/assets/example-slides.json`](pptx_generator/assets/example-slides.json)

### 方式二：手動撰寫 Markdown

````markdown
---
title: 專案進度報告
version: 2026-Q2
---

# 專案進度報告
Q2 2026 Review

## 關鍵成果
- 查詢延遲下降 42%
  - p95 由 2.3s → 1.3s

## 連線快取實作

```python
@lru_cache(maxsize=32)
def get_conn(dsn): ...
```
````

```bash
pptx-generate --input slides.md --out output.pptx -v
```

| Markdown 語法 | 投影片類型 |
|--------------|-----------|
| `# H1`（第一個） | `title_slide`（封面） |
| `# H1`（後續） | `section_slide`（章節分隔） |
| `## H2` | 新投影片（`bullet_points` 或 `code_demo`） |
| 項目符號 `-` / `*` | `bullet_points` |
| 程式碼區塊 | `code_demo` |
| `> 引用` | speaker notes |

> Markdown 僅支援基本類型。需要 `kpi_slide`、`table`、`two_column` 等進階類型時，請用 JSON 或 YAML。

完整範例：[`pptx_generator/assets/example-slides.md`](pptx_generator/assets/example-slides.md)

### 方式三：YAML

比 JSON 更易讀寫，支援扁平格式（不需要巢狀 `content`）。需安裝 PyYAML。

```bash
pptx-generate --input slides.yaml --out output.pptx -v
```

完整範例：[`pptx_generator/assets/example-slides.yaml`](pptx_generator/assets/example-slides.yaml)

### 方式四：在 AI IDE 中使用

安裝為 Skill 後，直接在聊天視窗說「幫我把這份會議記錄做成簡報」，AI 自動完成全部流程。詳見 [作為 AI Skill 使用](#作為-ai-skill-使用)。

---

## CLI 使用方式

```bash
# 基本用法（支援 .json / .yaml / .md）
pptx-generate --input slides.md --out output.pptx -v

# 完整參數（Linux / macOS 可用 \ 換行，Windows 請寫成一行）
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

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `--input` | 投影片資料路徑 `.json` / `.yaml` / `.md`（必要） | — |
| `--template` | .pptx 模板路徑 | 內建 `default-template.pptx` |
| `--out` | 輸出路徑 | `output_presentation.pptx` |
| `--brand-color` | 品牌色 HEX | `#2B579A` |
| `--font` | 覆蓋字體 | 微軟正黑體 / Calibri |
| `--footer` | 頁腳文字 | metadata.title |
| `--version-label` | 版本標籤（底部中央） | metadata.version |
| `--watermark` | 浮水印文字 | — |
| `--page-numbers` | 顯示頁碼 | 關閉 |
| `-v` / `-vv` | 詳細輸出 | WARNING |

> `--json` 仍可使用（向後相容），建議改用 `--input`。

---

## Python API

```python
from pathlib import Path
from pptx_generator import generate, BrandConfig

brand = BrandConfig(color="#007A33", font="Noto Sans TC", footer="My Company", page_numbers=True)

output = generate(
    json_path=Path("slides.md"),       # 支援 .json / .yaml / .md
    template_path=None,                 # None = 使用內建模板
    output_path=Path("output.pptx"),
    brand=brand,
)
```

```python
from pptx_generator import parse_markdown

data = parse_markdown("# Title\nSubtitle\n\n## Key Points\n- Point 1\n- Point 2")
print(data["slides"])  # 可傳給 generate() 或自行修改
```

---

## 作為 AI Skill 使用

安裝後，AI 可自動完成從內容分析到 PPTX 生成的完整流程（Phase 0–4）。

### 安裝 Skill

依你使用的 AI agent 選擇：

**Kiro：**

```bash
# Linux / macOS
git clone --depth 1 https://github.com/paul0728/pptx-generator.git .kiro/skills/pptx-generator
cd .kiro/skills/pptx-generator && rm -rf .git .github .gitattributes .gitignore LICENSE MANIFEST.in pyproject.toml README.md README.zh-TW.md requirements.txt && cd -
pip install python-pptx requests Pillow
```

```powershell
# Windows (PowerShell)
git clone --depth 1 https://github.com/paul0728/pptx-generator.git .kiro/skills/pptx-generator
Remove-Item -Force .kiro/skills/pptx-generator/.git, .kiro/skills/pptx-generator/.github, .kiro/skills/pptx-generator/.gitattributes, .kiro/skills/pptx-generator/.gitignore, .kiro/skills/pptx-generator/LICENSE, .kiro/skills/pptx-generator/MANIFEST.in, .kiro/skills/pptx-generator/pyproject.toml, .kiro/skills/pptx-generator/README.md, .kiro/skills/pptx-generator/README.zh-TW.md, .kiro/skills/pptx-generator/requirements.txt -Recurse -ErrorAction SilentlyContinue
pip install python-pptx requests Pillow
```

**Claude Code / Cursor / Codex 等（40+ agents）：**

```bash
npx skills add paul0728/pptx-generator
pip install python-pptx requests Pillow
```

> [`npx skills`](https://github.com/vercel-labs/skills) 預設安裝到 project-level（如 `.claude/skills/`）。加 `-g` 為全域安裝。

```bash
npx skills list                            # 列出已安裝 skills
npx skills update                          # 更新所有 skills
npx skills remove pptx-generator           # 移除 skill
```

### 安裝後的目錄結構

```
your-project/
├── .kiro/skills/pptx-generator/     ← Kiro
│   ├── SKILL.md
│   └── pptx_generator/
│       ├── generator.py
│       └── assets/
│
├── .claude/skills/pptx-generator/   ← Claude Code（npx skills 建立）
│   └── ...
```

### 觸發方式

| 語言 | 範例 |
|------|------|
| 中文 | 「幫我做簡報」「把這份文件轉成投影片」 |
| 英文 | "Make a PPT" "Convert this to slides" |
| 進階 | 「套用 ./template.pptx，品牌色 #007A33，控制在 12 頁」 |

---

## 投影片類型

| type | 用途 | content 欄位 |
|------|------|-------------|
| `title_slide` | 封面（第 1 頁） | `title`, `sub_title` |
| `outline_slide` | 大綱（第 2 頁） | `title`, `points[]` |
| `section_slide` | 章節分隔 | `title`, `sub_title` |
| `bullet_points` | 條列內容 | `title`, `points[]` |
| `architecture_diagram` | Mermaid 架構圖 | `title`, `mermaid_code`, `description` |
| `code_demo` | 程式碼展示 | `title`, `code`, `language` |
| `two_column` | 左右對比 | `title`, `left{heading, points}`, `right{heading, points}` |
| `table` | 資料表格 | `title`, `headers[]`, `rows[][]` |
| `image_slide` | 圖片 + 說明 | `title`, `image_path`, `caption` |
| `kpi_slide` | KPI 卡片（≤6） | `title`, `kpis[{label, value, unit, delta}]` |
| `ending_slide` | 結尾頁 | `title`, `sub_title` |

---

## Pipeline

```
Phase 0  規劃大綱（AI 根據 prompt-template.md 自動執行）
Phase 1  內容解析 → slides.json（AI 產出 / 使用者提供）
Phase 2  Mermaid 圖表平行渲染（快取 + CJK fallback）
Phase 3  python-pptx 組裝（Layout 查找 → auto-fit → 品牌 chrome）
Phase 4  品質驗證
```

---

## 專案結構

```
pptx-generator/
├── pptx_generator/              # Python 套件
│   ├── generator.py             # 核心邏輯
│   ├── markdown_parser.py       # Markdown → slides
│   └── assets/                  # 內建模板 + 範例
│       ├── prompt-template.md   # AI 提示詞樣板（核心）
│       ├── example-slides.*     # JSON / YAML / MD 範例
│       └── default-template.pptx
├── SKILL.md                     # AI Skill 定義
├── pyproject.toml
├── README.md                    # English
└── README.zh-TW.md             # 繁體中文
```

---

## License

MIT — 詳見 [LICENSE](LICENSE)。
