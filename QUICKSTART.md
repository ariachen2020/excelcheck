# 快速開始指南

## 🚀 3 分鐘快速開始

### 第一步：安裝依賴（1 分鐘）

```bash
pip install -r requirements.txt
```

### 第二步：創建示例文件（1 分鐘）

```bash
python create_sample_files.py
```

這會創建兩個示例 Excel 檔案：
- `correct_sample.xlsx` - 正確的檔案
- `test_sample.xlsx` - 帶有格式差異的檔案

### 第三步：運行應用（1 分鐘）

```bash
streamlit run app.py
```

應用會自動在瀏覽器中打開（通常是 `http://localhost:8501`）

---

## 💡 使用示例

1. **上傳檔案**
   - 點擊左側「正確的檔案」→ 選擇 `correct_sample.xlsx`
   - 點擊右側「待檢查的檔案」→ 選擇 `test_sample.xlsx`

2. **開始比對**
   - 點擊「🔍 開始比對」按鈕

3. **查看結果**
   - 檢查頂部的統計摘要
   - 展開各個項目查看詳細結果

### 預期看到的結果

示例檔案會顯示以下差異：
- ❌ 資料類型：年齡欄位類型不同
- ⚠️  欄位順序：部門和薪資位置交換
- ✅ 其他項目都會通過檢查

---

## 📁 項目結構

```
excelcheck/
├── app.py                    # Streamlit 應用（主文件）
├── create_sample_files.py    # 創建示例檔案的腳本
├── requirements.txt          # Python 依賴
├── README.md                 # 詳細說明文檔
├── DEPLOY.md                 # 部署指南
├── QUICKSTART.md             # 本文件
├── .gitignore                # Git 忽略文件
└── .streamlit/
    └── config.toml           # Streamlit 配置
```

---

## 🎯 功能清單

應用會自動檢查：

- ✓ **欄位名稱和順序** - 檢查列名和順序是否相同
- ✓ **資料類型** - 檢查數據類型是否一致
- ✓ **儲存格格式** - 檢查單元格的原始格式
- ✓ **數值精度** - 檢查小數位數和長度
- ✓ **空值處理** - 檢查缺失值的數量
- ✓ **資料筆數** - 檢查行數是否相同

---

## ⚠️ 常見問題

### 我的應用無法打開怎麼辦？

```bash
# 確保安裝了所有依賴
pip install -r requirements.txt

# 嘗試清除 Streamlit 緩存
rm -rf ~/.streamlit/

# 重新運行應用
streamlit run app.py
```

### 上傳的檔案無法讀取？

- 確保檔案格式是 `.xls` 或 `.xlsx`
- 檢查檔案是否被其他程式打開
- 嘗試重新保存檔案

### 如何更改主題配色？

編輯 `.streamlit/config.toml` 文件中的顏色值：

```toml
[theme]
primaryColor = "#FF6B6B"      # 改變主色
backgroundColor = "#FFFFFF"    # 背景色
secondaryBackgroundColor = "#F0F2F6"  # 次級背景色
textColor = "#262730"         # 文本色
```

---

## 📤 下一步

### 本地測試完成後

1. **初始化 Git 倉庫**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   ```

2. **推送到 GitHub**
   ```bash
   git remote add origin https://github.com/yourusername/excelcheck.git
   git branch -M main
   git push -u origin main
   ```

3. **部署到 Streamlit Cloud**
   - 前往 https://streamlit.io/cloud
   - 使用 GitHub 登錄
   - 創建新應用並選擇此倉庫

詳見 [DEPLOY.md](DEPLOY.md)

---

## 💬 获得帮助

- 📖 詳細文檔：[README.md](README.md)
- 🚀 部署指南：[DEPLOY.md](DEPLOY.md)
- 🐛 報告問題：在 GitHub 上提出 Issue

---

## 📝 許可證

MIT License - 可自由使用和修改

---

## 🎉 開始吧！

```bash
# 現在就試試看
streamlit run app.py
```

祝您使用愉快！
