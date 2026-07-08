# 社内サーバー / クラウド（Cloud Run・Azure 等）で動かすためのコンテナ定義
FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

# コンテナの死活監視
HEALTHCHECK CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://localhost:8501/_stcore/health').status==200 else 1)"

ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.port=8501", "--server.address=0.0.0.0", \
            "--server.headless=true"]
