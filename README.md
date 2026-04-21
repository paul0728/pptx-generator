# PPTX Generator

將任何文件、報告或結構化內容轉換為企業級 PPTX 簡報。

## 功能特色

- **多種投影片類型**：封面、大綱、章節、條列、架構圖、程式碼、雙欄對比、表格、圖片、KPI 卡片、結尾
- **Mermaid 圖表渲染**：自動將 Mermaid 語法轉為高解析度 PNG 嵌入投影片
- **CJK 感知自動排版**：中英文混合內容自動縮放字體，確保不溢出
- **品牌客製化**：支援品牌色、自訂字體、頁腳、頁碼、浮水印
- **模板支援**：可套用自訂 `.pptx` 模板，自動偵測 Layout
- **品質驗證**：生成後自動檢查每頁是否有可見內容

## 安裝

```bash
pip install -r requirements.txt
```

### 依賴套件

- `python-pptx` >= 0.6.21
- `requests` >= 2.28.0
- `Pillow` >= 9.0.0

## 快速開始

### 1. 準備 slides.json

參考 `assets/example-slides.json` 建立你的投影片結構：

```json
{
  "presentation_metadata": {
    "title": "我的簡報",
    "version": "2026-04-21"
  },
  "slides": [
    {
      "id": 1,
      "type": "title_slide",
      "content": {
        "title": "專案報告",
        "sub_title": "2026 Q2"
      }
    }
  ]
}
```

### 2. 執行生成

```bash
python scripts/generate_pptx_template.py \
    --json slides.json \
    --out output.pptx \
    -v
```

### 完整參數

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `--json` | slides.json 路徑（必要） | — |
| `--template` | .pptx 模板路徑 | `assets/default-template.pptx` |
| `--out` | 輸出檔案路徑 | `output_presentation.pptx` |
| `--brand-color` | 品牌色 HEX | `#2B579A` |
| `--font` | 覆蓋字體 | 微軟正黑體 / Calibri |
| `--footer` | 頁腳文字 | metadata.title |
| `--version-label` | 版本標籤（底部中央） | metadata.version |
| `--watermark` | 浮水印文字 | — |
| `--page-numbers` | 顯示頁碼 | 關閉 |
| `-v` / `-vv` | 詳細輸出 (INFO / DEBUG) | WARNING |

## 投影片類型

| type | 用途 |
|------|------|
| `title_slide` | 封面（固定第 1 頁） |
| `outline_slide` | 大綱目錄（固定第 2 頁） |
| `section_slide` | 章節分隔 |
| `bullet_points` | 條列內容 |
| `architecture_diagram` | Mermaid 流程/架構圖 |
| `code_demo` | 程式碼展示 |
| `two_column` | 左右兩欄對比 |
| `table` | 資料表格 |
| `image_slide` | 單張圖片 + 說明 |
| `kpi_slide` | KPI 卡片（最多 6 個） |
| `ending_slide` | 結尾頁 |

## Pipeline 流程

```
Phase 1  載入並驗證 slides.json
Phase 2  Mermaid 圖表平行渲染（含快取與 CJK fallback）
Phase 3  python-pptx 組裝（Layout 查找 → auto-fit → 品牌 chrome）
Phase 4  品質驗證
```

## 專案結構

```
pptx-generator/
├── scripts/
│   └── generate_pptx_template.py   # 主程式
├── assets/
│   ├── default-template.pptx       # 預設模板
│   ├── example-output.pptx         # 範例輸出
│   └── example-slides.json         # 範例 JSON
├── .cache/                          # Mermaid 渲染快取（自動產生）
├── SKILL.md                         # 詳細技術文件
├── requirements.txt
├── .gitignore
└── README.md
```

## 進階用法

### 使用自訂模板

```bash
python scripts/generate_pptx_template.py \
    --json slides.json \
    --template my-corporate-template.pptx \
    --brand-color "#007A33" \
    --font "Noto Sans TC" \
    --footer "Company · Confidential" \
    --page-numbers \
    -v
```

### 加入浮水印

```bash
python scripts/generate_pptx_template.py \
    --json slides.json \
    --watermark "DRAFT" \
    -v
```

## 內容量限制

為確保投影片不溢出，JSON 中的內容應遵守：

| 類型 | 上限 |
|------|------|
| 標題 | ≤ 20 中文字 |
| bullet_points | ≤ 10 項，每項 ≤ 40 字 |
| code_demo | ≤ 20 行 |
| table | ≤ 12 列 × 5 欄 |
| kpi_slide | ≤ 6 個 KPI |

## License

MIT
