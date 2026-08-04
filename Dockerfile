# 採用 Python 3.12 slim 基礎映像檔
FROM python:3.12-slim

# 安裝 LibreOffice 與中文字型 (確保雲端環境 Word/Excel 轉檔無中文亂碼)
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# 複製需求套件並安裝
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製應用程式所有內容
COPY . .

# 設定環境變數
ENV HOST=0.0.0.0
ENV PORT=8080

EXPOSE 8080

# 啟動 Web 轉檔伺服器
CMD ["python", "app.py"]
