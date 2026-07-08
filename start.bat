@echo off
cd /d "%~dp0"
title ファイル軽量化アプリ

echo ============================================================
echo   ファイル軽量化アプリ を起動します
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 goto NOPYTHON

echo 必要な部品を準備しています...
echo （初回だけ数分かかります。文字が流れている間は正常です。お待ちください）
echo.
python -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto PIPERROR

echo.
echo ------------------------------------------------------------
echo  アプリを起動します。この画面は「開いたまま」にしてください。
echo   ・少し待つとブラウザが自動で開きます
echo   ・社内の人には、下に出る Network URL を伝えてください
echo   ・終了する時は、この画面で Ctrl + C を押します
echo ------------------------------------------------------------
echo.
python -m streamlit run app.py --server.address=0.0.0.0 --server.port=8501

echo.
echo アプリを終了しました。何かキーを押すと閉じます。
pause
exit /b 0

:NOPYTHON
echo [エラー] Python が見つかりません。
echo.
echo   Python をインストールしてください:
echo     1. https://www.python.org/downloads/ を開く
echo     2. 「Download Python」からインストーラを実行
echo     3. 最初の画面で「Add python.exe to PATH」に必ずチェック
echo     4. パソコンを再起動して、もう一度 start.bat を実行
echo.
echo この画面は閉じません。確認したら何かキーを押してください。
pause
exit /b 1

:PIPERROR
echo.
echo [エラー] 部品の準備に失敗しました。
echo   ・インターネットに接続されているか確認してください
echo   ・もう一度 start.bat を実行してみてください
echo   ・それでもダメなら、上に出ているメッセージを控えて担当者にお知らせください
echo.
pause
exit /b 1
