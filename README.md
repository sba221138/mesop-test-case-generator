# 🤖 AI 測試案例生成器 (AI Test Case Generator)

這是一個基於 Python 與 Google Gemini AI 的自動化工具，能夠協助 QA 工程師快速將需求文件（Word, PDF, 純文字）轉換為結構化的測試案例（Test Cases），並支援匯出為 CSV 格式以便匯入 Excel 或測試管理系統。

## ✨ 主要功能

- **多格式支援**：支援讀取 `.docx`, `.pdf`, `.txt` 格式的需求文件。
- **AI 智能分析**：使用 Google Gemini 2.5 Flash 模型，自動理解需求邏輯。
- **結構化輸出**：自動生成包含「編號、標題、預置條件、步驟、預期結果」的標準測試表格。
- **CSV 下載**：一鍵匯出帶有中文表頭的 CSV 檔案，完美相容 Excel。
- **現代化 UI**：使用 Google Mesop 框架打造的簡潔網頁介面。
- **安全設計**：支援 `.env` 環境變數管理 API Key，並使用 `.gitignore` 保護敏感資料。

## 🛠️ 技術棧

- **Python**: 核心程式語言
- **Mesop**: Google 推出的 Python UI 框架
- **Google Generative AI**: Gemini 模型介接
- **Pandas**: 表格資料處理與 CSV 輸出
- **pdf2image / Poppler**: PDF 轉圖片處理

## 🚀 快速開始 (Installation)

### 1. 環境準備
請確保你的電腦已安裝 Python 3.10 或以上版本。

### 2. 下載專案
```bash
git clone https://github.com/sba221138/mesop-test-case-generator
cd mesop-test-case-generator
```

### 3. 建立虛擬環境 (推薦)
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# Mac/Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 4. 安裝依賴套件
```bash
pip install -r requirements.txt
```

### 5. 設定環境變數
請在專案根目錄建立一個 .env 檔案，並填入你的 Google API Key：
In your `.env` file context
```txt
API_KEY=你的_GOOGLE_GEMINI_API_KEY
```

> ⚠️ 注意：本專案依賴 Poppler 工具來處理 PDF。請確保專案目錄下有 bin/poppler 資料夾，或已將 Poppler 加入系統環境變數。

https://github.com/oschwartz10612/poppler-windows/releases/tag/v25.12.0-0

## 🏃‍♂️ 如何執行 (Usage)
在終端機執行以下指令啟動網頁伺服器：
``` bash
mesop main.py
```
啟動後，瀏覽器會自動開啟，或請手動前往：http://localhost:32123

## 📂 專案結構
```
mesop-test-case-generator/
├── bin/                 # Poppler 執行檔 (PDF處理用)
├── .env                 # API Key 設定檔 (請勿上傳)
├── .gitignore           # Git 忽略清單
├── main.py              # 主程式入口
├── requirements.txt     # 套件依賴清單
└── README.md            # 專案說明文件
```

## 📝 License (MIT)