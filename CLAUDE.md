# CV Conference Cases - Project Context

## 概述
心血管影像跨科部討論會（CV Conference）案例資料庫。每月一次討論會，每次 2-3 筆案例，涵蓋 MPI / CTA / CAG 檢查結果。

## 資料結構

### 檔案來源
- 資料夾內的 `CV list YYYYMMDD.docx` / `.doc` / `.pdf` 檔案
- 每個檔案包含一個 7 欄 Word 表格：Name/Chart No, Age/Gender, Reason of MPI, Risk Factors, MPI, CTA, CAG

### 資料庫位置
- **Excel 總表**: `CV_Conference_Database.xlsx`（統計研究用）
- **Notion Database**: "CV Conference Cases v3"，Data Source ID: `d3512b4c-5371-4928-a19b-d3ff69b4dd70`
  - Database URL: https://www.notion.so/7bbdd0d027574d2ca2f1aa78a449029e
  - 位於 Notion 的「心血管影像跨科部討論會」頁面下
- **舊的 Notion Database** (v2, 有錯誤資料): https://www.notion.so/89629fc261334d509c753bd48e27beaf — 可刪除
- **舊的 Notion Database** (v1, 有錯誤資料): https://www.notion.so/c5fad0747e304d6a8e0d39787cb29b5d — 可刪除

### 欄位定義
| 欄位 | Excel | Notion | 說明 |
|------|-------|--------|------|
| Conference Date | 日期文字 | Date property | 討論會日期，從檔名提取 |
| Chart No | 文字 | Rich Text | 病歷號（7 碼） |
| Name | 文字 | Title 的一部分 | 姓名（已遮蔽） |
| Age | 數字 | Number | 年齡 |
| Gender | M/F | Select | 性別 |
| Reason of MPI | 文字 | Rich Text | MPI 檢查原因 |
| Risk Factors | 文字 | Multi-Select | HTN, DM, DLP, Smoking, Obesity, Age, Gender |
| MPI Dates | 文字 | Rich Text | MPI 檢查日期（可多個） |
| CTA Dates | 文字 | Rich Text | CTA 檢查日期 |
| CAG Dates | 文字 | Rich Text | CAG 檢查日期 |
| Data Quality | 標記 | Select | Complete / Partial / Needs Review |
| Source File | 文字 | Rich Text | 來源檔名 |

### Data Quality 定義
- **Complete**: 從 .docx 完整提取，7 欄位齊全
- **Partial**: 從 .doc 提取，部分欄位缺失（尤其 Risk Factors）
- **Needs Review**: PDF 來源或提取失敗，需人工核對

## 常見操作

### 新增一次討論會資料（標準流程）

**您只需要做一件事**：把新的 CV list 檔案放進 `input/`，然後告訴 Claude：
> 「有新的 CV list，請處理」

Claude 會自動執行以下步驟：

```bash
# Step 1: 偵測哪些是新檔案（比對 Excel 的 Source File 欄）
python parse_cv_list.py --new

# Step 2: 解析新檔案並更新 Excel
python parse_cv_list.py --new --update-excel

# Step 3: Claude 讀取 notion_import.json，匯入 Notion（append）
# 使用 Notion MCP create-pages
# parent: {"data_source_id": "d3512b4c-5371-4928-a19b-d3ff69b4dd70"}
```

⚠️ **不要用 `--batch --update-excel`**，否則會將所有舊資料重複寫入 Excel。

### 完整重建（從零開始）
```bash
# 重新解析所有檔案並重建 Excel（需先手動刪除舊 Excel）
python parse_cv_list.py --batch --update-excel

# 查看統計
python parse_cv_list.py --stats
```

### Notion 匯入注意事項
- Notion MCP `create-pages` 每次最多 100 筆
- Risk Factors 需以 JSON array 格式傳入，例如 `["HTN", "DM"]`
- Conference Date 使用 expanded format: `date:Conference Date:start` = ISO date
- **絕對不要讓 AI agent 自行生成資料**，必須從 JSON 檔案讀取

## 技術備註
- .doc 檔案使用 `olefile` + UTF-16LE 解碼提取文字（需安裝 `olefile`）
- .docx 使用 `python-docx`
- .pdf 使用 `PyMuPDF (fitz)`
- `CV list 20180829.doc` 格式特殊，目前無法自動解析，需手動處理
- 所有依賴: `python-docx`, `openpyxl`, `olefile`, `PyMuPDF`
