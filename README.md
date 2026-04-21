# PPTX Generator

[![PyPI version](https://badge.fury.io/py/ai-pptx-generator.svg)](https://pypi.org/project/ai-pptx-generator/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

將任何文件、報告或結構化內容轉換為企業級 PPTX 簡報。

支援三種輸入格式：**Markdown**、**YAML**、**JSON** — 不需要手寫 JSON，用 Markdown 就能生成簡報。

也可作為 AI IDE Skill 使用，讓 AI 自動完成從內容分析到 PPTX 生成的完整流程。

---

## 目錄

- [功能特色](#功能特色)
- [安裝方式](#安裝方式)
- [CLI 使用方式](#cli-使用方式)
- [輸入格式](#輸入格式)
- [Python API](#python-api)
- [作為 AI Skill 使用](#作為-ai-skill-使用)
- [投影片類型](#投影片類型)
- [Pipeline 流程](#pipeline-流程)
- [專案結構](#專案結構)
- [License](#license)

---

## 功能特色

- **三種輸入格式**：Markdown（最直覺）、YAML（可讀性佳）、JSON（完整控制）
- **11 種投影片類型**：封面、大綱、章節、條列、架構圖、程式碼、雙欄對比、表格、圖片、KPI 卡片、結尾
- **Mermaid 圖表渲染**：自動將 Mermaid 語法轉為高解析度 PNG 嵌入投影片
- **CJK 感知自動排版**：中英文混合內容自動縮放字體，確保不溢出
- **品牌客製化**：支援品牌色、自訂字體、頁腳、頁碼、浮水印
- **模板支援**：可套用自訂 `.pptx` 模板，自動偵測 Layout
- **品質驗證**：生成後自動檢查每頁是否有可見內容

---

## 安裝方式

### 方式一：從 PyPI 安裝（推薦）

**Windows (PowerShell):**
```powershell
pip install ai-pptx-generator
```

**Linux / macOS:**
```bash
pip install ai-pptx-generator
```

如需 YAML 輸入支援：

**Windows (PowerShell):**
```powershell
pip install "ai-pptx-generator[yaml]"
```

**Linux / macOS:**
```bash
pip install "ai-pptx-generator[yaml]"
```

### 方式二：從 GitHub 安裝

**Windows (PowerShell):**
```powershell
pip install git+https://github.com/paul0728/pptx-generator.git
```

**Linux / macOS:**
```bash
pip install git+https://github.com/paul0728/pptx-generator.git
```

### 方式三：開發模式安裝

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

安裝後即可使用 `pptx-generate` 指令，或 `python -m pptx_generator`。

---

## CLI 使用方式

### 基本用法 — 從 Markdown 生成

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.md --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.md --out output.pptx -v
```

### 從 YAML 生成

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.yaml --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.yaml --out output.pptx -v
```

### 從 JSON 生成（向後相容）

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.json --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.json --out output.pptx -v
```

> `--json slides.json` 仍可使用（向後相容），但建議改用 `--input`。

### 完整參數

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

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `--input` | 投影片資料路徑（`.json` / `.yaml` / `.yml` / `.md`）（必要） | — |
| `--template` | .pptx 模板路徑 | 內建 `default-template.pptx` |
| `--out` | 輸出檔案路徑 | `output_presentation.pptx` |
| `--brand-color` | 品牌色 HEX | `#2B579A` |
| `--font` | 覆蓋字體 | 微軟正黑體 / Calibri |
| `--footer` | 頁腳文字 | metadata.title |
| `--version-label` | 版本標籤（底部中央） | metadata.version |
| `--watermark` | 浮水印文字 | — |
| `--page-numbers` | 顯示頁碼 | 關閉 |
| `-v` / `-vv` | 詳細輸出 (INFO / DEBUG) | WARNING |

---

## 輸入格式

### Markdown（推薦 — 最直覺）

用 Markdown 撰寫投影片內容，工具自動轉換為簡報結構。不需要學任何特殊格式。

```markdown
---
title: 專案進度報告
version: 2026-Q2
---

# 專案進度報告
Q2 2026 Review

## 大綱
- 專案背景
- 系統架構
- 關鍵成果

# 系統架構

## 關鍵成果
- 查詢延遲下降 42%
  - p95 由 2.3s → 1.3s
- 錯誤率下降至 0.4%

## 連線快取實作

```python
@lru_cache(maxsize=32)
def get_conn(dsn): ...
```

## Thank You
Q & A
```

**Markdown 對應規則：**

| Markdown 語法 | 投影片類型 |
|--------------|-----------|
| `# H1`（第一個） | `title_slide`（封面） |
| `# H1`（後續） | `section_slide`（章節分隔） |
| `## H2` | 新投影片（`bullet_points` 或 `code_demo`） |
| 項目符號 `-` / `*` | `bullet_points` 的 points |
| 程式碼區塊 ` ``` ` | `code_demo` |
| `> 引用` | speaker notes |
| YAML front matter | `presentation_metadata` |

完整範例：[`assets/example-slides.md`](assets/example-slides.md)

### YAML（可讀性佳）

比 JSON 更易讀寫，支援扁平格式（不需要巢狀 `content`）：

```yaml
title: 專案進度報告
version: 2026-Q2

slides:
  - type: title_slide
    title: 專案進度報告
    sub_title: "Q2 2026 Review\n2026-04-21"

  - type: bullet_points
    title: 關鍵成果
    points:
      - 查詢延遲下降 42%
      - 錯誤率下降至 0.4%

  - type: code_demo
    title: 連線快取實作
    language: python
    code: |
      @lru_cache(maxsize=32)
      def get_conn(dsn: str) -> Connection:
          return psycopg2.connect(dsn)
```

> YAML 輸入需要安裝 PyYAML：`pip install pyyaml` 或 `pip install "ai-pptx-generator[yaml]"`

完整範例：[`assets/example-slides.yaml`](assets/example-slides.yaml)

### JSON（完整控制）

適合程式化生成或需要精確控制每個欄位的場景。

完整範例：[`assets/example-slides.json`](assets/example-slides.json)

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

# 支援 .json / .yaml / .md 輸入
output = generate(
    json_path=Path("slides.md"),      # 或 slides.yaml / slides.json
    template_path=None,                # 使用內建模板
    output_path=Path("output.pptx"),
    brand=brand,
)
print(f"簡報已生成：{output}")
```

### Markdown 解析 API

```python
from pptx_generator import parse_markdown

data = parse_markdown("""
# My Presentation
Subtitle here

## Key Points
- Point 1
- Point 2
""")

print(data["slides"])  # 可直接傳給 generate 或自行修改後再生成
```

---

## 作為 AI Skill 使用

本工具可作為 AI IDE 的 Skill 使用。安裝後，AI 可自動完成從內容分析到 PPTX 生成的完整流程。

### 安裝步驟

#### 步驟 1：Clone 到專案的 `.skills/` 目錄

**Windows (PowerShell):**
```powershell
git clone https://github.com/paul0728/pptx-generator.git .skills/pptx-generator
```

**Linux / macOS:**
```bash
git clone https://github.com/paul0728/pptx-generator.git .skills/pptx-generator
```

或者手動複製：

**Windows (PowerShell):**
```powershell
New-Item -ItemType Directory -Path .skills\pptx-generator -Force
Copy-Item -Recurse path\to\pptx-generator\* .skills\pptx-generator\
```

**Linux / macOS:**
```bash
mkdir -p .skills/pptx-generator
cp -r /path/to/pptx-generator/* .skills/pptx-generator/
```

最終目錄結構應為：

```
your-project/
├── .skills/
│   └── pptx-generator/
│       ├── SKILL.md                 ← AI 讀取此檔案理解能力
│       ├── scripts/
│       │   └── generate_pptx_template.py
│       ├── assets/
│       │   ├── default-template.pptx
│       │   ├── example-output.pptx
│       │   ├── example-slides.json
│       │   ├── example-slides.yaml
│       │   └── example-slides.md
│       └── pptx_generator/          ← Python 套件
│           └── ...
├── src/
└── ...
```

#### 步驟 2：安裝 Python 依賴

**Windows (PowerShell):**
```powershell
pip install python-pptx requests Pillow
```

**Linux / macOS:**
```bash
pip install python-pptx requests Pillow
```

或者直接安裝套件（依賴會自動安裝）：

**Windows (PowerShell):**
```powershell
pip install ai-pptx-generator
```

**Linux / macOS:**
```bash
pip install ai-pptx-generator
```

#### 步驟 3：在 AI IDE 中使用

在聊天視窗中直接說：

```
幫我把這份會議記錄做成簡報
```

AI 會自動：
1. 辨識到你要做簡報 → 啟動 pptx-generator Skill
2. 分析你的內容，規劃大綱（Phase 0）
3. 生成結構化的投影片資料（Phase 1）
4. 渲染 Mermaid 圖表（Phase 2）
5. 組裝 PPTX 檔案（Phase 3）
6. 品質驗證（Phase 4）

### 觸發方式

以下任何說法都會自動觸發此 Skill：

| 語言 | 範例 |
|------|------|
| 中文 | 「幫我做簡報」「把這份文件轉成投影片」「做一份進度報告 PPT」 |
| 英文 | "Make a PPT" "Convert this to slides" "Generate a slide deck" |
| 指定模板 | 「套用 ./template.pptx 做簡報」 |
| 指定風格 | 「做成 KPI dashboard 風格，品牌色 #007A33」 |
| 指定頁數 | 「控制在 12 頁」 |
| 指定受眾 | 「給老闆看的精簡版」 |

### SKILL.md 的角色

`SKILL.md` 是 AI Skill 的核心設定檔，它告訴 AI：

1. **何時觸發**：front-matter 中的 `name` 和 `description` 讓 AI 判斷使用者意圖
2. **如何執行**：Pipeline 定義（Phase 0-4）指導 AI 的工作流程
3. **格式規範**：投影片資料 schema、內容量限制、排版規則
4. **錯誤處理**：各種 fallback 策略

你可以根據需求修改 `SKILL.md` 來調整 AI 的行為。

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

## Pipeline 流程

```
Phase 0  規劃大綱（AI Skill 模式，AI 自動執行）
Phase 1  內容解析 → 投影片資料（Markdown / YAML / JSON）
Phase 2  Mermaid 圖表平行渲染（含快取與 CJK fallback）
Phase 3  python-pptx 組裝（Layout 查找 → auto-fit → 品牌 chrome）
Phase 4  品質驗證
```

---

## 專案結構

```
pptx-generator/
├── pptx_generator/                  # Python 套件（pip install 用）
│   ├── __init__.py
│   ├── __main__.py                  # python -m pptx_generator
│   ├── generator.py                 # 核心邏輯
│   ├── markdown_parser.py           # Markdown → slides 轉換
│   └── assets/
│       ├── default-template.pptx
│       └── example-slides.json
├── scripts/
│   └── generate_pptx_template.py    # Skill 用的入口（向後相容）
├── assets/                           # 範例與資源檔
│   ├── default-template.pptx
│   ├── example-output.pptx
│   ├── example-slides.json
│   ├── example-slides.yaml
│   └── example-slides.md
├── SKILL.md                          # AI Skill 定義
├── pyproject.toml                    # PyPI 套件設定
├── requirements.txt
├── MANIFEST.in
├── LICENSE
└── README.md
```

---

## License

MIT License — 詳見 [LICENSE](LICENSE)。
