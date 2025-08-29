Write-Host "仕入レポート生成プログラムを開始します..." -ForegroundColor Green
Write-Host ""

# 仮想環境をアクティベート
& ".\venv\Scripts\Activate.ps1"

# プログラムを実行
python purchase_report_generator.py

Write-Host ""
Write-Host "プログラムの実行が完了しました。" -ForegroundColor Green
Read-Host "Enterキーを押して終了"
