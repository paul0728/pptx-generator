---
name: pptx-generator
description: >
  Generates corporate-compliant PPTX presentations from any report, document,
  or structured content. Pipeline: outline planning → structured data →
  Mermaid diagram rendering → python-pptx assembly → quality verification.
  Use this skill whenever the user wants to turn any text, report, analysis,
  meeting notes, or system document into a polished slide deck. Trigger even if
  the user just says "make a PPT", "convert this to slides", "幫我做簡報",
  "把這份文件轉成投影片", or any similar phrasing — regardless of topic or domain.
---

# PPTX Presentation Generator

將任何文件或報告轉換為可直接用於公司報告的 PPTX 簡報。

---

## 推薦使用流程

**使用者不需要手寫投影片資料。** 推薦流程如下：

```
使用者提供需求描述 / 原始內容
        ↓
AI 根據「提示詞樣板」自動產出 slides.json
        ↓
pptx-generate 生成 PPTX 簡報
```

### 提示詞樣板 (Prompt Template)

本工具提供一份標準化的提示詞樣板：**[pptx_generator/assets/prompt-template.md](pptx_generator/assets/prompt-template.md)**

此樣板包含：
- 完整的 slides.json schema
- 所有 11 種投影片類型的範例
- 內容量限制規則
- 排版規則
- 完整的輸出範例

**使用方式：**
1. 將 `prompt-template.md` 的內容貼給任何 AI（ChatGPT、Claude、Gemini 等）
2. 接著提供你的需求描述或原始內容（文件、報告、筆記、口頭指示皆可）
3. AI 產出 `slides.json`
4. 執行 `pptx-generate --input slides.json --out output.pptx -v`

**在 AI IDE 中更簡單：** 直接說「幫我做簡報」，AI 會自動讀取此樣板並完成全部流程。

---

## 使用者輸入

### 必要

| 輸入 | 說明 |
|------|------|
| **內容來源** | 報告、文件、需求描述、口頭指示、URL — 任何形式皆可 |

### 可選

使用者可以用自然語言提供以下任何資訊，AI 自動理解並調整：

| 輸入 | 範例 | 未提供時的行為 |
|------|------|---------------|
| 模板檔案 | 「套用 ./template.pptx」 | 優先用 `pptx_generator/assets/default-template.pptx`；不存在則用空白簡報 |
| 目標受眾 | 「給老闆看」「給 IT 團隊」 | 依內容性質自動判斷 |
| 語言 | 「用英文」「中英混合」 | 跟隨內容來源語言；混合時預設繁體中文 |
| 頁數 | 「控制在 15 頁」 | 依內容量自動估算 |
| 風格 | 「圖表為主」「精簡摘要」 | 依內容性質平衡 |
| 輸出路徑 | 「存到 docs/report.pptx」 | `output_presentation.pptx` |
| 品牌色 | 「用 #007A33」 | 有模板繼承模板；無模板用 `#2B579A` |
| 字體 | 「用 Noto Sans TC」 | 中文 `微軟正黑體`，程式碼 `Consolas` |
| 既有投影片資料 | 「用這份 slides.json / .yaml / .md」 | 從 Phase 0 開始完整規劃 |

**使用者不需要知道這些參數。** 直接說「幫我做簡報」就能跑。

---

## 支援的輸入格式

工具支援三種輸入格式。推薦讓 AI 根據 prompt template 產出 JSON，但也可手動撰寫：

### 1. JSON（推薦 — AI 產出 + 完整控制）

由 AI 根據 [prompt-template.md](pptx_generator/assets/prompt-template.md) 自動產出，或手動撰寫。

- 範例：[pptx_generator/assets/example-slides.json](pptx_generator/assets/example-slides.json)

### 2. Markdown（手動撰寫最直覺）

用 Markdown 撰寫投影片內容，工具自動轉換為簡報結構：

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

## 關鍵成果
- 查詢延遲下降 42%
  - p95 由 2.3s → 1.3s
```

Markdown 對應規則：
- `# H1` → 封面（第一個）或章節分隔（後續）
- `## H2` → 新投影片
- 項目符號 → bullet_points
- 程式碼區塊 → code_demo
- `> 引用` → speaker notes

> 注意：Markdown 格式僅支援 `title_slide`、`section_slide`、`bullet_points`、`code_demo`。
> 需要 `kpi_slide`、`table`、`two_column` 等進階類型時，請使用 JSON 或 YAML。

### 3. YAML（可讀性佳）

比 JSON 更易讀寫，支援扁平格式（不需要巢狀 `content`）：

```yaml
title: 專案進度報告
version: 2026-Q2

slides:
  - type: title_slide
    title: 專案進度報告
    sub_title: Q2 2026 Review

  - type: bullet_points
    title: 關鍵成果
    points:
      - 查詢延遲下降 42%
      - 錯誤率下降至 0.4%
```

> YAML 輸入需要安裝 PyYAML：`pip install pyyaml`

---

## Pipeline

```
Phase 0  規劃大綱（AI 根據 prompt-template.md 自動執行）
Phase 1  內容解析 → slides.json（AI 產出，或使用者提供 JSON / YAML / Markdown）
Phase 2  Mermaid 圖表渲染
Phase 3  python-pptx 組裝
Phase 4  品質驗證
```

- **Phase 0 + 1 由 AI 自動完成**：AI 讀取 [pptx_generator/assets/prompt-template.md](pptx_generator/assets/prompt-template.md) 中的 schema 與規則，根據使用者提供的內容自動產出 `slides.json`。
- 若使用者已提供投影片資料檔（JSON / YAML / Markdown），跳過 Phase 0 和 Phase 1，直接進入 Phase 2。
- 範例檔案：[pptx_generator/assets/example-slides.json](pptx_generator/assets/example-slides.json) / [pptx_generator/assets/example-slides.yaml](pptx_generator/assets/example-slides.yaml) / [pptx_generator/assets/example-slides.md](pptx_generator/assets/example-slides.md)
- 提示詞樣板：[pptx_generator/assets/prompt-template.md](pptx_generator/assets/prompt-template.md)
- 依賴：`python-pptx`、`requests`、`Pillow`（YAML 輸入另需 `pyyaml`）

### Quickstart

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.md --out output.pptx -v
```

**Linux / macOS:**
```bash
pptx-generate --input slides.md --out output.pptx -v
```

完整參數範例：

**Windows (PowerShell):**
```powershell
pptx-generate --input slides.yaml --template assets/default-template.pptx --out output.pptx --brand-color "#007A33" --font "Noto Sans TC" --footer "Company · Confidential" --version-label "2026-Q2 v1.2" --watermark "DRAFT" --page-numbers -v
```

**Linux / macOS:**
```bash
pptx-generate \
    --input slides.yaml \
    --template assets/default-template.pptx \
    --out output.pptx \
    --brand-color "#007A33" --font "Noto Sans TC" \
    --footer "Company · Confidential" --version-label "2026-Q2 v1.2" \
    --watermark "DRAFT" --page-numbers \
    -v
```

> `--json` 仍可使用（向後相容），但建議改用 `--input`。

未指定 `--template` 時，優先使用 `pptx_generator/assets/default-template.pptx`，不存在則用空白簡報。
未指定 `--footer` / `--version-label` 時，會自動使用 `presentation_metadata.title` / `.version`。

---

## Phase 0 — 大綱規劃

> **此階段由 AI 自動執行。** AI 讀取 [pptx_generator/assets/prompt-template.md](pptx_generator/assets/prompt-template.md) 中的規則來規劃大綱。

1. 根據內容量與需求，估算合理頁數。
2. 列出每一頁的標題與 slide type。
3. **大綱固定為第 2 頁**（封面之後），使用 `outline_slide`。
4. **每頁只講一個主題**。內容太多時拆成多頁（如「需求清單 1/2」「需求清單 2/2」）。
5. 大綱項目超過 10 個時，考慮合併章節或分兩頁大綱。

---

## Phase 1 — 內容解析 → slides.json

> **此階段由 AI 自動執行。** AI 根據 [pptx_generator/assets/prompt-template.md](pptx_generator/assets/prompt-template.md) 中的 schema、slide types、內容量限制，將使用者的原始內容轉換為 `slides.json`。

### slides.json Schema

```json
{
  "presentation_metadata": {
    "title": "string",
    "version": "YYYY-MM-DD"
  },
  "slides": [
    {
      "id": 1,
      "type": "<slide type>",
      "content": {
        "title":        "string — 所有類型，≤ 20 中文字（40 英文字元）",
        "sub_title":    "string — title_slide / section_slide",
        "points":       ["string array — bullet_points / outline_slide"],
        "mermaid_code": "string — architecture_diagram",
        "description":  "string — architecture_diagram",
        "code":         "string — code_demo",
        "language":     "string — code_demo",
        "notes":        "string — speaker notes（可選，任何類型）"
      }
    }
  ]
}
```

### Slide Types

| type | 用途 | 固定位置 |
|------|------|---------|
| `title_slide` | 封面 | 第 1 頁 |
| `outline_slide` | 大綱目錄 | 第 2 頁 |
| `section_slide` | 章節分隔 | 每章開頭 |
| `bullet_points` | 條列內容 | — |
| `architecture_diagram` | 流程/架構圖 | — |
| `code_demo` | 程式碼/指令 | — |
| `two_column` | 左右兩欄對比（優缺點、方案 A/B） | — |
| `table` | 資料表格（規格比較、版本歷程） | — |
| `image_slide` | 單張圖片 + 說明 | — |
| `kpi_slide` | KPI 卡片（大數字 + 標籤 + 變動） | — |
| `ending_slide` | 結尾 | 最後一頁 |

#### 額外 type 的 content 欄位

```jsonc
// two_column
{ "title": "...", "left": {"heading": "優勢", "points": ["..."]},
                   "right": {"heading": "風險", "points": ["..."]} }

// table
{ "title": "...", "headers": ["欄1", "欄2"], "rows": [["a", "b"], ["c", "d"]] }

// image_slide
{ "title": "...", "image_path": "path/to/img.png", "caption": "圖說" }

// kpi_slide  (最多 6 個)
{ "title": "...", "kpis": [
    {"label": "延遲 p95", "value": "1.3", "unit": "s", "delta": "▼ 42%"}
]}
```

### 內容量控制（防溢出的第一道防線）

Phase 1 生成資料時就必須控制每頁內容量。Phase 3 的 auto-fit 是第二道防線，不可作為主要依賴。

| 類型 | 上限 | 超過時 |
|------|------|--------|
| 標題 | ≤ 20 中文字（40 英文字元） | 精簡用語 |
| `bullet_points` | ≤ 10 項，每項 ≤ 40 中文字 | 拆成多頁 |
| `code_demo` | ≤ 20 行 | 拆成多頁或精簡 |
| `architecture_diagram` | description ≤ 1 行；mermaid 節點 ≤ 15 | 拆成多張圖 |
| `outline_slide` | ≤ 10 項 | 合併章節或分兩頁 |
| `two_column` | 每欄 ≤ 6 項 | 改用兩張 `bullet_points` |
| `table` | ≤ 12 列 × 5 欄 | 分頁或改用 `kpi_slide` |
| `kpi_slide` | ≤ 6 個 KPI | 拆成多頁 |

### 類型選擇規則

根據內容性質選擇最適合的投影片類型，不要全部用 `bullet_points`：

| 內容性質 | 應使用的類型 | 說明 |
|---------|------------|------|
| 有順序的步驟、流程、pipeline | `architecture_diagram` | 用 Mermaid flowchart 畫流程圖，比條列更直觀 |
| 系統架構、元件關係、資料流 | `architecture_diagram` | 用 Mermaid 呈現元件之間的連接關係 |
| 兩個方案 / 優缺點對比 | `two_column` | 左右並排比較 |
| 數據指標、KPI、成果數字 | `kpi_slide` | 大數字卡片，視覺衝擊力強 |
| 規格比較、版本歷程、多欄資料 | `table` | 結構化表格 |
| 程式碼、指令、設定檔 | `code_demo` | 等寬字體展示 |
| 一般說明、要點列舉 | `bullet_points` | 條列內容 |

> **重要：** 遇到「步驟 1 → 步驟 2 → 步驟 3」這類有順序性的操作流程時，
> 優先使用 `architecture_diagram` 搭配 Mermaid flowchart，而非 `bullet_points`。
> 流程圖能讓觀眾一眼看出步驟之間的先後關係和分支邏輯。

### 排版規則

| 規則 | 說明 |
|------|------|
| 每頁一個主題 | 不可在一頁塞多個不相關的概念 |
| 所有內容不可溢出 | Phase 1 內容量控制 + Phase 3 auto-fit 雙重保障 |
| 副標題換行 | 用 `\n`，不可同時用分隔符又換行 |
| Bullet 標記 | 主項用 `•`、`-` 或編號；子項用 `  -`（兩空格縮排 + 短橫線） |
| Section 主標 | 章節標題文字，禁止數字編號 |
| Section 副標 | 僅補充說明時使用，否則留空 |

---

## Phase 2 — Mermaid 圖表渲染

對每個含 `mermaid_code` 的 slide：

1. Base64 編碼 → `GET https://mermaid.ink/img/<base64>?width=1600`（高解析度）
2. timeout = 30 秒；失敗重試 3 次
3. 以 `sha256(mermaid_code)` 為 key 快取到 `.cache/mermaid/<hash>.png`，同輸入不重打 API
4. 多張圖平行渲染（最多 4 個 worker）
5. 失敗時標記錯誤，不中斷流程

### Mermaid 中文問題

mermaid.ink 伺服器端可能未安裝中文字體，中文節點會顯示為方塊。

處理優先順序：
1. **優先使用英文標籤**，中文說明放在 slide 的描述文字中
2. 若必須用中文，先嘗試渲染
3. 若中文渲染失敗，自動將 CJK 節點標籤替換為 ASCII placeholder 重試
4. 若仍失敗，fallback 為純文字說明的投影片

---

## Phase 3 — PPTX 組裝

### 3-A 檢查模板 Layout

**每次使用新模板必須先執行**，用名稱查找：

```python
from pptx import Presentation

prs = Presentation("template.pptx")
layout_map = {}
for i, layout in enumerate(prs.slide_layouts):
    layout_map[layout.name] = i
    print(i, layout.name)
    for ph in layout.placeholders:
        print(f"  idx={ph.placeholder_format.idx} type={ph.placeholder_format.type}")
```

Layout 對應策略：
1. 精確名稱匹配（如 `layout_map["一般"]`）
2. 子字串匹配（如名稱含 "封面" 的 layout）
3. 都找不到 → fallback 到 index 0 並印出警告

無模板時用 `Presentation()` 預設空白簡報：
- Layout 0 = Title Slide（idx 0=title, 1=subtitle）
- Layout 1 = Title and Content（idx 0=title, 1=body）
- Layout 2 = Section Header
- Layout 5 = Blank

### 3-B 自動排版引擎（Auto-fit）

**所有文字插入都必須經過 auto-fit**。禁止硬編碼字數閾值。

> **注意：** `shape.text = text` 會清除 placeholder 原有的模板文字格式（字體、顏色、對齊）。
> 這是 python-pptx 的限制。auto-fit 會重新設定字體大小，但顏色和對齊會繼承模板的
> paragraph-level 設定。若模板格式全在 run-level，則需要手動補回。

演算法（中英文混合感知）：

```python
from pptx.util import Pt

def auto_fit_text(shape, text: str, max_pt=18, min_pt=7, is_code=False):
    """Insert text and auto-shrink font to fit within placeholder bounds.
    CJK-aware: Chinese chars ≈ 1x font width, ASCII ≈ 0.55x."""
    shape.text = text
    if not text.strip():
        return
    pw = shape.width / 914400   # inches
    ph = shape.height / 914400  # inches

    for pt in range(max_pt, min_pt - 1, -1):
        char_w = pt / 72  # 1 CJK char width in inches
        total_lines = 0
        for line in text.split("\n"):
            line_width = sum(
                char_w if ord(c) > 127 else char_w * 0.55
                for c in line
            )
            total_lines += max(1, int(line_width / pw + 0.99))

        line_h = pt / 72 * 1.35
        if total_lines * line_h <= ph:
            for p in shape.text_frame.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(pt)
                    if is_code:
                        r.font.name = "Consolas"
            return

    # Fallback: min_pt still doesn't fit — apply min_pt and warn
    print(f"  ⚠️ Content overflows at min {min_pt}pt, consider splitting this slide")
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(min_pt)
            if is_code:
                r.font.name = "Consolas"
```

字體範圍：

| 元素 | max_pt | min_pt | is_code |
|------|--------|--------|---------|
| 標題 | 24 | 12 | No |
| Body / Bullet | 18 | 7 | No |
| 副標題 | 14 | 8 | No |
| 程式碼 | 12 | 6 | **Yes** |

### 3-C 圖片排版

architecture_diagram 的圖片與描述文字不可重疊：

1. Pillow 讀取圖片像素尺寸，**以 aspect ratio 等比縮放**（不依賴 DPI 反推）
2. 等比例縮放至 max 寬 `slide_width - 1"` × 高 3.7"
3. 水平置中：`left = (slide_width - image_width) / 2`
4. 垂直位置：`top = 1.1"`
5. 描述文字**必須用 `add_textbox` 放在圖片正下方**（`top = 1.1 + image_height + 0.15"`），不可用 body placeholder，否則會與圖片重疊
6. 為圖片設定 `descr` / `title`（alt-text）以利無障礙
   - 這是「禁止浮動 shape」原則的唯一例外，因為圖片已佔用 body placeholder 區域

### 3-D 程式碼頁面

code_demo 必須強制等寬字體 `Consolas`，透過 auto-fit 的 `is_code=True` 參數實現。

### 3-E Speaker Notes

若 slide 資料中有 `notes` 欄位：

```python
if c.get("notes"):
    slide.notes_slide.notes_text_frame.text = c["notes"]
```

### 3-F 錯誤處理

| 情境 | 處理 |
|------|------|
| Mermaid 渲染失敗 | 顯示 placeholder 文字，繼續下一張 |
| Layout 找不到 | 精確 → 子字串 → fallback index 0，印出警告 |
| Placeholder idx 不存在 | 跳過，印出警告 |
| auto-fit 到 min_pt 仍溢出 | 套用 min_pt，印出警告建議拆頁 |
| 儲存時檔案被鎖定 | 自動加 `-v2` 後綴重試，最多 3 次 |

### 3-G 品牌色 / 字體 / 頁腳（全域 chrome）

生成完所有 slide 後，統一套用：

| 項目 | 來源 | 說明 |
|------|------|------|
| **品牌色** | `--brand-color #HEX`（預設 `#2B579A`） | 用於 KPI 數值、KPI 卡片框線、表格標題底色 |
| **字體** | `--font "Noto Sans TC"` | 覆蓋所有非程式碼 run；程式碼強制保留 Consolas |
| **頁腳** | `--footer "..."` 或 `presentation_metadata.title` | 左下角，9pt 灰字 |
| **版本/日期** | `--version-label "..."` 或 `presentation_metadata.version` | 下方中央 |
| **頁碼** | `--page-numbers` | 右下角 `n / total` |
| **浮水印** | `--watermark "CONFIDENTIAL"` | 45° 淺灰大字，置於底層 |

封面 (`title_slide`) 與結尾 (`ending_slide`) 自動略過 footer/頁碼，避免干擾。

### 3-H 完成

1. 儲存 `.pptx`（Windows 檔案鎖定時自動加 `-v2` 後綴重試）
2. 清理 temp diagram files

---

## Phase 4 — 品質驗證

生成後自動檢查：

```python
prs = Presentation(output_path)
issues = []
for i, slide in enumerate(prs.slides):
    has_content = any(
        hasattr(s, "text") and s.text.strip() for s in slide.shapes
    )
    if not has_content:
        issues.append(f"Slide {i+1}: no visible content")

expected = len(data["slides"])
actual = len(prs.slides)
if actual != expected:
    issues.append(f"Expected {expected} slides, got {actual}")

if issues:
    print("⚠️ Quality issues:")
    for issue in issues:
        print(f"  - {issue}")
else:
    print("✅ All slides passed quality check")
```

---

## 無模板時的預設樣式

| 元素 | 顏色 | 字體 |
|------|------|------|
| 標題 | `#1A1A1A` | 微軟正黑體 / Calibri |
| 內文 | `#333333` | 微軟正黑體 / Calibri |
| 程式碼 | `#333333` | **Consolas**（強制） |
| 強調 / Section 背景 | `#2B579A` | — |
| Section 文字 | `#FFFFFF` | — |

---

## Trigger Examples

> "幫我把這份會議記錄做成簡報。"
> "把這份文件轉成投影片，套用 ./template.pptx。"
> "做一份給主管看的專案進度報告，精簡一點。"
> "Generate a slide deck from this doc, keep it visual."
> "幫我做簡報，不用模板，簡單乾淨就好。"
> "把 report.md 轉成客戶提案簡報，控制在 12 頁。"
> "做成 KPI dashboard 風格，品牌色 #007A33，加頁碼。"

## Anti-triggers — 不要啟動這個 skill

- 只要求 **Markdown 大綱** 或 **純文字摘要**（沒提到簡報/投影片/PPT/deck）
- 只問 **Mermaid 語法** 或 **python-pptx API**
- 只要產生 **slides data** 但明確說「不要 pptx」
- 只要求 **修改現有 .pptx 檔案的特定物件**（這個 skill 是從頭生成，不做 in-place 編輯）
