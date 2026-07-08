@echo off
chcp 65001 >nul
cd /d "%~dp0"
title ファイル軽量化アプリ
echo ============================================================
echo   ファイル軽量化アプリ を起動します
echo ============================================================
echo.
echo [1/2] 必要な部品を確認しています...（初回だけ数分かかります）
python -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 (
  echo.
  echo [エラー] Python が見つからないか、部品の準備に失敗しました。
  echo         README.md の「準備」を見て、Python を入れ直してください。
  echo.
  pause
  exit /b 1
)
echo.
echo [2/2] アプリを起動します。
echo.
echo   ● この黒い画面は「開いたまま」にしてください（閉じるとアプリが止まります）
echo   ● 少し待つとブラウザが自動で開きます
echo   ● 同じ社内ネットワークの人は、下に表示される "Network URL"
echo     （http://192.168.x.x:8501 のような住所）をブラウザに入れると使えます
echo   ● 終わるとき: この画面で Ctrl + C を押すか、画面を閉じてください
echo.
python -m streamlit run app.py --server.address=0.0.0.0 --server.port=8501
echo.
echo アプリを終了しました。
pause
