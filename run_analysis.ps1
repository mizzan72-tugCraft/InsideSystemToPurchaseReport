Write-Host "仕入レポート分析プログラムを開始します..." -ForegroundColor Green
Write-Host ""

# 仮想環境をアクティベート
& ".\venv\Scripts\Activate.ps1"

# Pythonプログラムを実行
python purchase_analysis_generator.py

# 仮想環境を非アクティベート
deactivate

Write-Host ""
Write-Host "プログラムが終了しました。" -ForegroundColor Green
Read-Host "Enterキーを押して終了"

# 2025-08-29 17:30:00
