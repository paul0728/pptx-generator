# PPTX Generator

[![PyPI version](https://badge.fury.io/py/ai-pptx-generator.svg)](https://pypi.org/project/ai-pptx-generator/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

將任何文件、報告或結構化內容轉換為企業級 PPTX 簡報。

支援兩種使用方式：
- **CLI / Python 套件** — `pip install` 後直接用指令生成簡報
- **Kiro AI Skill** — 在 Kiro IDE 中用自然語言驅動，AI 自動完成整個流程

---

## 目錄

- [功能特色](#功能特色)
- [安裝方式](#安裝方式)
- [CLI 使用方式](#cli-使用方式)
- [Python API 使用方式](#python-api-使用方式)
- [作為 Kiro Skill 使用](#作為-kiro-skill-使用)
- [slides.json 格式](#slidesjson-格式)
- [投影片類型](#投影片類型)
- [Pipeline 流程](#pipeline-流程)
- [專案結構](#專案結構)
- [License](#license)

---

## 功能特色

- **11 種投影片類型**：封面、大綱、章節、條列、架構圖、程式碼、雙欄對比、表格、圖片、KPI 卡片、結尾
- **Mermaid 圖表渲染**：自動將 Mermaid 語法轉為高解析度 PNG 嵌入投影片
- **CJK 感知自動排版**：中英文混合內容自動縮放字體，確保不溢出
- **品牌客製化**：支援品牌色、自訂字體、頁腳、頁碼、浮水印
- **模板支援**：可套用自訂 `.pptx` 模板，自動偵測 Layout
- **品質驗證**：生成後自動檢查每頁是否有可見內容

---

## 安裝方式

### 方式一：從 PyPI 安裝（推薦）

```bash
pip install ai-pptx-generator
```

### 方式二：從 GitHub 安裝

```bash
pip install git+https://github.com/paul0728/pptx-generator.git
```

### 方式三：開發模式安裝

```bash
git clone https://github.com/paul0728/pptx-generator.git
cd pptx-generator
pip install -e .
```

安裝後即可使用 `pptx-generate` 指令，或 `python -m pptx_generator`。

---

## CLI 使用方式

### 基本用法

```bash
pptx-generate --json slides.json --out output.pptx -v
```

### 完整參數

```bash
pptx-generate \
    --json slides.json \
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
| `--json` | slides.json 路徑（必要） | — |
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

## Python API 使用方式

```python
from pathlib import Path
from pptx_generator import generate, BrandConfig

brand = BrandConfig(
    color="#007A33",
    font="Noto Sans TC",
    footer="My Company",
    page_numbers=True,
)

output = generate(
    json_path=Path("slides.json"),
    template_path=None,       # 使用內建模板
    output_path=Path("output.pptx"),
    brand=brand,
)
print(f"簡報已生成：{output}")
```

---

## 作為 Kiro Skill 使用

這是本工具最強大的用法。安裝為 Kiro Skill 後，你只需要用自然語言告訴 AI「幫我做簡報」，AI 就會自動完成從內容分析到 PPTX 生成的完整流程。

### 什麼是 Kiro Skill？

[Kiro](https://kiro.dev) 是一款 AI 驅動的 IDE。Skill 是 Kiro 的擴充能力模組，透過 `SKILL.md` 檔案定義觸發條件與執行流程，讓 AI 在對話中自動辨識使用者意圖並啟動對應的工作流程。

### 安裝步驟

#### 步驟 1：Clone 到專案的 `.skills/` 目錄

在你的專案根目錄下執行：

```bash
git clone https://github.com/paul0728/pptx-generator.git .skills/pptx-generator
```

或者手動複製：

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
│       │   └── example-slides.json
│       └── pptx_generator/          ← Python 套件（可選）
│           └── ...
├── src/
└── ...
```

#### 步驟 2：安裝 Python 依賴

```bash
pip install python-pptx requests Pillow
```

或者如果你已經 `pip install ai-pptx-generator`，依賴會自動安裝。

#### 步驟 3：在 Kiro 中使用

打開 Kiro，在聊天視窗中直接說：

```
幫我把這份會議記錄做成簡報
```

AI 會自動：
1. 辨識到你要做簡報 → 啟動 pptx-generator Skill
2. 分析你的內容，規劃大綱（Phase 0）
3. 生成結構化的 `slides.json`（Phase 1）
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

### 進階用法

#### 提供既有 slides.json

如果你已經有 `slides.json`，可以跳過 AI 規劃階段：

```
用這份 slides.json 幫我生成簡報
```

AI 會跳過 Phase 0 和 Phase 1，直接從 Phase 2 開始。

#### 自訂模板

```
幫我做簡報，套用 ./my-template.pptx，品牌色 #007A33，字體用 Noto Sans TC
```

#### 搭配其他 Skill

pptx-generator 可以與其他 Skill 串接。例如：
- 先用資料分析 Skill 產出報告 → 再用 pptx-generator 轉成簡報
- 先用會議記錄 Skill 整理重點 → 再用 pptx-generator 做成投影片

### SKILL.md 的角色

`SKILL.md` 是 Kiro Skill 的核心設定檔，它告訴 AI：

1. **何時觸發**：front-matter 中的 `name` 和 `description` 讓 AI 判斷使用者意圖
2. **如何執行**：Pipeline 定義（Phase 0-4）指導 AI 的工作流程
3. **格式規範**：slides.json schema、內容量限制、排版規則
4. **錯誤處理**：各種 fallback 策略

你可以根據需求修改 `SKILL.md` 來調整 AI 的行為，例如：
- 修改預設語言
- 調整內容量上限
- 新增自訂的投影片類型
- 變更觸發條件

---

## slides.json 格式

```json
{
  "presentation_metadata": {
    "title": "簡報標題",
    "version": "2026-04-21"
  },
  "slides": [
    {
      "id": 1,
      "type": "title_slide",
      "content": {
        "title": "專案報告",
        "sub_title": "Q2 2026",
        "notes": "Speaker notes（可選）"
      }
    }
  ]
}
```

完整範例請參考 [`assets/example-slides.json`](assets/example-slides.json)。

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
Phase 0  規劃大綱（Kiro Skill 模式，AI 自動執行）
Phase 1  內容解析 → slides.json
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
│   └── assets/
│       ├── default-template.pptx
│       └── example-slides.json
├── scripts/
│   └── generate_pptx_template.py    # Skill 用的入口（向後相容）
├── assets/                           # Skill 用的資源檔
│   ├── default-template.pptx
│   ├── example-output.pptx
│   └── example-slides.json
├── SKILL.md                          # Kiro Skill 定義
├── pyproject.toml                    # PyPI 套件設定
├── requirements.txt
├── MANIFEST.in
├── LICENSE
└── README.md
```

---

## License

MIT License — 詳見 [LICENSE](LICENSE)。
