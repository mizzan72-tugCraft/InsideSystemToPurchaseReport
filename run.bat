@echo off
echo 仕入レポート生成プログラムを開始します...
echo.

REM 仮想環境をアクティベート
call venv\Scripts\activate.bat

REM プログラムを実行
python purchase_report_generator.py

echo.
echo プログラムの実行が完了しました。
pause
