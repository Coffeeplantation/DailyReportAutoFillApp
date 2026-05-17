3@echo off
chcp 65001 > nul
echo ===================================================
echo  日報自動入力アプリ  ビルドスクリプト
echo ===================================================
echo.

echo [1/2] 必要パッケージをインストール中...
pip install pyinstaller openpyxl jpholiday
if %errorlevel% neq 0 (
    echo パッケージのインストールに失敗しました。
    pause
    exit /b 1
)

echo.
echo [2/2] exe ファイルをビルド中...
pyinstaller --onefile --windowed --name "日報自動入力アプリ" daily_report_app.py
if %errorlevel% neq 0 (
    echo ビルドに失敗しました。
    pause
    exit /b 1
)

echo.
echo ===================================================
echo  完了！
echo  dist\日報自動入力アプリ.exe を配布してください。
echo  settings.json は exe と同じフォルダに自動生成されます。
echo ===================================================
pause
