#!/usr/bin/env bash
# Mac / Linux 用の起動スクリプト（Windows の方は start.bat を使ってください）
set -e
cd "$(dirname "$0")"

echo "============================================================"
echo "  ファイル軽量化アプリ を起動します"
echo "============================================================"
echo
echo "[1/2] 必要な部品を確認しています...（初回だけ数分かかります）"
python3 -m pip install --quiet --disable-pip-version-check -r requirements.txt

echo
echo "[2/2] アプリを起動します。"
echo "  ● このウィンドウは開いたままにしてください"
echo "  ● 同じ社内ネットワークの人は、下に出る Network URL を開くと使えます"
echo "  ● 終わるとき: Ctrl + C"
echo
python3 -m streamlit run app.py --server.address=0.0.0.0 --server.port=8501
