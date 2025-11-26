# 部署指南

## 本地運行

### 快速開始

```bash
# 1. 進入項目目錄
cd excelcheck

# 2. 安裝依賴
pip install -r requirements.txt

# 3. 運行應用
streamlit run app.py
```

應用會自動在瀏覽器中打開，通常是 `http://localhost:8501`

### 使用虛擬環境（推薦）

```bash
# 創建虛擬環境
python -m venv venv

# 激活虛擬環境
# 在 macOS/Linux:
source venv/bin/activate
# 在 Windows:
venv\Scripts\activate

# 安裝依賴
pip install -r requirements.txt

# 運行應用
streamlit run app.py

# 停止應用
# 按 Ctrl+C
```

## 部署到 Streamlit Cloud

### 前置要求
- GitHub 帳號
- 專案已推送到 GitHub

### 部署步驟

1. **準備 GitHub 倉庫**
   ```bash
   cd excelcheck
   git init
   git add .
   git commit -m "Initial commit: Excel format checker tool"
   git branch -M main
   git remote add origin https://github.com/yourusername/excelcheck.git
   git push -u origin main
   ```

2. **訪問 Streamlit Cloud**
   - 前往 https://streamlit.io/cloud
   - 使用 GitHub 帳號登錄

3. **創建新應用**
   - 點擊 "New app" 按鈕
   - 選擇 "GitHub" 作為源
   - 選擇你的 `excelcheck` 倉庫
   - 選擇分支：`main`
   - 選擇主文件路徑：`app.py`
   - 點擊 "Deploy"

4. **配置應用**
   - 應用會自動部署
   - 你將獲得一個公共 URL（例如：`https://excelcheck.streamlit.app`）
   - 可以分享此 URL 給其他用戶

## 部署到其他平台

### Heroku

1. **創建 Procfile**
   ```
   web: streamlit run --logger.level=error --client.showErrorDetails=false app.py
   ```

2. **部署**
   ```bash
   heroku create your-app-name
   git push heroku main
   ```

### Docker

1. **創建 Dockerfile**
   ```dockerfile
   FROM python:3.9-slim

   WORKDIR /app

   COPY requirements.txt .
   RUN pip install -r requirements.txt

   COPY . .

   EXPOSE 8501

   CMD ["streamlit", "run", "app.py"]
   ```

2. **構建並運行**
   ```bash
   docker build -t excelcheck .
   docker run -p 8501:8501 excelcheck
   ```

## 環境配置

### 自定義主題

編輯 `.streamlit/config.toml` 文件以自定義應用外觀：

```toml
[theme]
primaryColor = "#FF6B6B"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"
font = "sans serif"
```

### 性能優化

在部署到 Streamlit Cloud 時，建議在 `.streamlit/config.toml` 中添加：

```toml
[client]
showErrorDetails = true
toolbarMode = "auto"

[logger]
level = "info"

[client.toolbarMode]
mode = "minimal"
```

## 常見問題

### Q: 應用無法找到上傳的檔案
A: 確保你有足夠的磁盤空間，並且檔案沒有被其他程式鎖定。

### Q: 部署到 Streamlit Cloud 後仍然很慢？
A:
- 檢查網絡連接
- 嘗試清除瀏覽器緩存
- 檢查 Streamlit Cloud 的狀態頁面

### Q: 可以限制上傳檔案的大小嗎？
A: 可以在 `.streamlit/config.toml` 中添加：
```toml
[server]
maxUploadSize = 200
```

### Q: 如何使用私密 GitHub 倉庫部署？
A: 在 Streamlit Cloud 中授予必要的 GitHub 權限即可。

## 維護

### 更新依賴
```bash
pip install --upgrade -r requirements.txt
```

### 查看日誌
- **本地**：在終端中查看
- **Streamlit Cloud**：在應用設置中查看

## 安全建議

1. 不要在代碼中提交敏感信息
2. 如果部署私密應用，使用身份驗證
3. 定期更新依賴以修復安全漏洞
4. 檢查上傳檔案的大小限制

## 支援

如有問題，請提出 GitHub Issue 或檢查：
- Streamlit 文檔：https://docs.streamlit.io
- Streamlit Cloud 文檔：https://docs.streamlit.io/streamlit-cloud
