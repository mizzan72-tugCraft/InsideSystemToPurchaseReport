# 仕入レポート生成プログラム

社内システムから出力された Excel ファイル（.xls, .xlsx）を成形し、仕入レポートを生成する Python プログラムです。

## 機能

- オリジナルデータの読み込みと文字化け修正
- 成形リストに基づくデータの絞り込み
- 分類置換テーブルによる分類名称の統一
- 分類別・ファイル別の集計機能
- ピボットテーブル形式のレポート生成（予定）

## 必要なファイル

### 入力ファイル（SampleData ディレクトリ内）

- `*_オリジナルデータ.xls` - 社内システムから出力された仕入データ
- `*_成形リスト.xlsx` - 必要な分類コードと置換名称のマッピング
- 分類置換テーブル - プログラム内に定義済み（分類コード -> 置換名称）

### 出力ファイル

- `purchase_report_YYYYMMDD_HHMMSS.json` - 詳細データ（JSON 形式、分析用に最適化）
- `purchase_summary_YYYYMMDD_HHMMSS.json` - 集計データ（JSON 形式）
- `analysis_results_YYYYMMDD_HHMMSS.json` - 分析結果（JSON 形式）
  N- `purchase_report_YYYYMMDD_HHMMSS.xlsx` - Excel ファイル（画像の列構成に準拠）
  - 分類コード、分類名称、仕入先コード、仕入先、ファイル No.、UNIT、No.、品名、メーカー、材質・型式、数、受入日、単価の列構成

## 環境構築

### 1. 仮想環境の作成とアクティベート

```bash
# 仮想環境の作成
python -m venv venv

# 仮想環境のアクティベート
# Windows (PowerShell)
.\venv\Scripts\Activate.ps1

# Windows (Command Prompt)
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 2. 必要なライブラリのインストール

```bash
pip install -r requirements.txt
```

### 3. プログラムの実行

```bash
python purchase_report_generator.py
```

## 使用ライブラリ

- `pandas==2.1.4` - データ処理
- `openpyxl==3.1.2` - Excel ファイル（.xlsx）の読み込み
- `xlrd==2.0.1` - Excel ファイル（.xls）の読み込み
- `numpy==1.24.3` - 数値計算

## 処理フロー

1. **オリジナルデータの読み込み** - 社内システムから出力された Excel ファイルを読み込み
2. **成形リストの読み込み** - 必要な分類コードと置換名称のマッピングを読み込み
3. **分類置換テーブルの適用** - 内部定義された分類コード -> 置換名称マッピングを適用
4. **成形リストの処理** - 有効な分類コードを抽出
5. **データの絞り込み** - 成形リストに含まれる分類コードのみを抽出
6. **集計処理** - 分類別・ファイル別の集計を実行
7. **データ出力** - JSON 形式と Excel 形式でデータを出力
8. **データ分析** - 分類別・仕入先別・月別の集計分析
9. **レポート生成** - 画像の列構成に準拠した Excel レポート生成

## 注意事項

- 仮想環境を使用することで、システム全体の Python 環境に影響を与えることなく、プロジェクト固有のライブラリを管理できます
- プログラムを実行する前に、必ず仮想環境をアクティベートしてください
- 入力ファイルは`SampleData`ディレクトリ内に配置してください
- 出力ファイルは`ReportOutput`ディレクトリに自動生成されます
- JSON ファイルは分析・グラフ作成・AI 予測に最適化されています
- データ分析ユーティリティ（`data_analyzer.py`）で簡単に分析可能です

## トラブルシューティング

### PowerShell の実行ポリシーエラー

```bash
# 管理者権限でPowerShellを開き、以下を実行
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 仮想環境のアクティベートができない場合

```bash
# 代替コマンド
venv\Scripts\activate
```

## 開発者向け情報

このプログラムは、Google Apps Script（GAS）の知識を持つユーザー向けに設計されています。
Python の`pandas`は、GAS の`SpreadsheetApp`に相当する機能を提供し、より効率的なデータ処理が可能です。
