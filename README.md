# CV Conference Case Parser

心血管影像跨科部討論會（CV Conference）案例資料庫自動化工具。
從每月的 CV list 檔案（.docx / .doc / .pdf）解析案例資料，同步至 Notion 資料庫與 Excel 總表。

---

## 快速開始：有新的 CV list 時

**您只需要兩步：**

1. 把新的 CV list 檔案放進 `data/` 資料夾
2. 打開 Claude Code，說：「有新的 CV list，請處理」

Claude 會自動偵測新檔案、解析資料、匯入 Notion、更新 Excel。

---

## 專案結構

```
parse_cv_list/
├── parse_cv_list.py          # 主要解析腳本
├── notion_import.json        # 最近一次的 Notion 匯入暫存檔
├── data/                     # 所有資料檔案
│   ├── CV_Conference_Database.xlsx  # Excel 總表（所有案例）
│   ├── CV list 20180926.doc
│   ├── CV list 20221019.docx
│   ├── CV list 20250114 (1).pdf
│   └── ...
├── CLAUDE.md                 # Claude Code 的操作指引（技術細節）
└── README.md                 # 本檔案
```

---

## 資料庫位置

| 位置 | 說明 | 連結 |
|------|------|------|
| **Excel** | `data/CV_Conference_Database.xlsx` | 本機檔案 |
| **Notion** | CV Conference Cases v3 | [開啟資料庫](https://www.notion.so/7bbdd0d027574d2ca2f1aa78a449029e) |
| **Notion 父頁面** | 心血管影像跨科部討論會 | [開啟頁面](https://www.notion.so/2f39af7a822c80049a61e8ae4aa6729a) |

---

## Notion 編輯注意事項

以下操作**不影響**未來 append：
- ✅ 在個別案例頁面內加筆記或評論
- ✅ 新增 View（篩選、排序、分組）
- ✅ 修改個別欄位的值（人工更正資料）
- ✅ 調整欄位顯示順序

以下操作會**導致 append 失敗或資料錯誤**，請勿進行：
- ❌ **更名欄位**（如把「Chart No」改名）→ 腳本找不到欄位
- ❌ **刪除欄位**（Case、Chart No、Conference Date 等核心欄位）
- ❌ **更改欄位類型**（如把 Date 改成 Text）
- ❌ **刪除 Risk Factors 的選項值**（HTN、DM、DLP 等）
- ❌ **移動或刪除整個資料庫**

---

## 腳本用法

```bash
# 偵測並處理新檔案（最常用）
python parse_cv_list.py --new

# 新檔案 + 同步更新 Excel
python parse_cv_list.py --new --update-excel

# 處理單一指定檔案
python parse_cv_list.py "CV list 20260318.docx"
python parse_cv_list.py "CV list 20260318.docx" --update-excel

# 查看統計摘要
python parse_cv_list.py --stats
```

> ⚠️ 不要用 `--batch --update-excel`（會重複寫入所有舊資料到 Excel）

---

## 如何知道哪些檔案已處理？

Excel 的 **Source File** 欄記錄每筆資料的來源檔名，`--new` 指令會自動比對 `data/` 資料夾與 Excel，找出尚未處理的新檔案。`data/` 裡的舊檔案可以刪除，不影響 Excel 的歷史記錄。

---

## 資料欄位

| 欄位 | 說明 |
|------|------|
| Conference Date | 討論會日期（從檔名提取） |
| Chart No | 病歷號 |
| Name | 姓名（已遮蔽） |
| Age | 年齡 |
| Gender | M / F |
| Reason of MPI | MPI 檢查原因 |
| Risk Factors | HTN / DM / DLP / Smoking / Obesity / Age / Gender |
| MPI / CTA / CAG Dates | 各項檢查日期 |
| Data Quality | Complete / Partial / Needs Review |
| Source File | 來源檔名（用於追蹤哪些檔案已處理） |

### Data Quality 說明

- **Complete**：來自 .docx，7 欄位齊全
- **Partial**：來自 .doc，Risk Factors 通常缺失
- **Needs Review**：來自 PDF，需人工核對

---

## 安裝依賴

```bash
pip install python-docx openpyxl olefile PyMuPDF
```

---

## 已知限制

- `CV list 20180829.doc`：格式特殊，無法自動解析，需手動處理
- `.doc` 檔案（2018–2021 年）：Risk Factors 無法自動提取，標記為 Partial
- PDF 檔案：僅能提取基本欄位，標記為 Needs Review
