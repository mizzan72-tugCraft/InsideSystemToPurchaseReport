@echo off
echo 仕入レポート分析プログラムを開始します...
echo.

REM 仮想環境をアクティベート
call venv\Scripts\activate.bat

REM Pythonプログラムを実行
python purchase_analysis_generator.py

REM 仮想環境を非アクティベート
deactivate

echo.
echo プログラムが終了しました。
pause

REM 2025-08-29 17:30:00
