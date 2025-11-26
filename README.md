# Excel 格式比對工具

一個簡單易用的 Streamlit 應用，用於比對兩個 Excel 檔案的格式差異。

## 功能

- ✓ 欄位名稱和順序檢查
- ✓ 資料類型檢查
- ✓ 數值格式（小數位數、長度）檢查
- ✓ 空值處理檢查
- ✓ 儲存格格式檢查
- ✓ 資料筆數檢查

## 安裝

### 方式 1：本地運行

1. **克隆或下載專案**
   ```bash
   cd excelcheck
   ```

2. **安裝依賴**
   ```bash
   pip install -r requirements.txt
   ```

3. **運行應用**
   ```bash
   streamlit run app.py
   ```

4. **訪問應用**
   - 應用會自動在瀏覽器中打開，通常是 `http://localhost:8501`

### 方式 2：部署到 Streamlit Cloud

1. **推送到 GitHub**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/yourusername/excelcheck.git
   git push -u origin main
   ```

2. **訪問 Streamlit Cloud**
   - 前往 https://streamlit.io/cloud
   - 使用 GitHub 帳號登錄
   - 點擊「New app」
   - 選擇此倉庫和 `app.py` 作為主文件

## 使用步驟

1. **上傳正確的 Excel 檔案**
   - 點擊左側「正確的檔案」上傳框
   - 選擇作為標準的 Excel 檔案

2. **上傳待檢查的 Excel 檔案**
   - 點擊右側「待檢查的檔案」上傳框
   - 選擇要驗證格式的 Excel 檔案

3. **開始比對**
   - 點擊「🔍 開始比對」按鈕
   - 等待檢查完成

4. **查看結果**
   - 查看頂部的統計摘要
   - 展開各個檢查項目查看詳細結果

## 支持的檔案格式

- `.xls` (Excel 97-2003)
- `.xlsx` (Excel 2007+)

## 系統需求

- Python 3.8+
- pip（Python 包管理器）

## 依賴

- **streamlit** - Web 應用框架
- **pandas** - 資料處理
- **xlrd** - Excel 讀取
- **openpyxl** - Excel 處理

## 常見問題

### Q: 為什麼無法讀取某些 Excel 檔案？
A: 確保檔案不是損壞的，且不被其他程式鎖定。如果是較新的 Excel 格式，請嘗試另存為 `.xlsx`。

### Q: 如何在本地私密部署？
A: 可以在自己的伺服器上運行 `streamlit run app.py`，或使用 Streamlit Community Cloud 進行部署。

### Q: 支持多個工作表嗎？
A: 目前只支持 Excel 檔案的第一個工作表。

## 許可證

MIT License

## 聯絡方式

如有問題或建議，請提出 Issue 或 PR。
