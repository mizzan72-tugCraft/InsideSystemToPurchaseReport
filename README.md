# 仕入レポート生成・分析プログラム

社内システムから出力された Excel ファイル（.xls, .xlsx）を成形し、仕入レポートを生成・分析する Python プログラムです。

## 機能

### 基本機能（purchase_report_generator.py）

- オリジナルデータの読み込みと文字化け修正
- 分類置換テーブルによる分類名称の統一
- 分類別・ファイル別の集計機能
- JSON 形式と Excel 形式でのデータ出力

### 分析機能（purchase_analysis_generator.py）

- ファイル No.別の詳細仕入分析
- カテゴリ → 仕入先 → 商品の階層別価格順ランク付け
- ブラウザ表示用 HTML レポート生成
- 縦長リストに最適化された表示

## 必要なファイル

### 入力ファイル（SampleData ディレクトリ内）

- `*_オリジナルデータ.xls` - 社内システムから出力された仕入データ
- `*_成形リスト.xlsx` - 必要な分類コードと置換名称のマッピング
- 分類置換テーブル - プログラム内に定義済み（分類コード -> 置換名称）

### 出力ファイル

#### 基本機能の出力

- `purchase_report_YYYYMMDD_HHMMSS.json` - 詳細データ（JSON 形式、分析用に最適化）
- `purchase_summary_YYYYMMDD_HHMMSS.json` - 集計データ（JSON 形式）
- `purchase_report_YYYYMMDD_HHMMSS.xlsx` - Excel ファイル（指定列構成に準拠）
  - 分類コード、分類名称、仕入先コード、仕入先、ファイル No.、UNIT、No.、品名、メーカー、材質・型式、数、受入日、単価の列構成

#### 分析機能の出力

- `purchase_analysis_[ファイルNo.]_YYYYMMDD_HHMMSS.html` - HTML レポート（ブラウザ表示用）
  - カテゴリ → 仕入先 → 商品の階層別表示
  - 価格順ランク付け
  - レスポンシブデザイン対応

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

#### 基本機能の実行

```bash
python purchase_report_generator.py
```

#### 分析機能の実行

```bash
python purchase_analysis_generator.py
```

#### バッチファイルでの実行（Windows）

```bash
# 基本機能
run_report.bat

# 分析機能
run_analysis.bat
```

#### PowerShell スクリプトでの実行（Windows）

```bash
# 基本機能
.\run_report.ps1

# 分析機能
.\run_analysis.ps1
```

## 使用ライブラリ

- `pandas==2.1.4` - データ処理
- `openpyxl==3.1.2` - Excel ファイル（.xlsx）の読み込み
- `xlrd==2.0.1` - Excel ファイル（.xls）の読み込み
- `numpy==1.24.3` - 数値計算

## 処理フロー

### 基本機能（purchase_report_generator.py）

1. **オリジナルデータの読み込み** - 社内システムから出力された Excel ファイルを読み込み
2. **分類置換テーブルの適用** - 内部定義された分類コード -> 置換名称マッピングを適用
3. **データ処理** - 全データを処理
4. **集計処理** - 分類別・ファイル別の集計を実行
5. **データ出力** - JSON 形式と Excel 形式でデータを出力

### 分析機能（purchase_analysis_generator.py）

1. **オリジナルデータの読み込み** - 社内システムから出力された Excel ファイルを読み込み
2. **分類置換テーブルの適用** - 内部定義された分類コード -> 置換名称マッピングを適用
3. **ファイル No.一覧の取得** - データに含まれるすべてのファイル No.を抽出
4. **ファイル No.選択** - ユーザーが分析対象のファイル No.を選択
5. **詳細分析** - カテゴリ → 仕入先 → 商品の階層別価格順ランク付け
6. **HTML レポート生成** - ブラウザ表示用の HTML レポートを生成
7. **ブラウザ表示** - 自動的にブラウザでレポートを表示

## 新機能の詳細

### ファイル No.別詳細分析機能

- **目的**: 特定のファイル No.について、仕入データを詳細に分析
- **階層構造**: カテゴリ → 仕入先 → 商品の 3 段階で整理
- **価格順ランク付け**: 各階層で仕入れ価格（単価 × 数量）の高い順に表示
- **表示形式**: ブラウザで見やすい HTML レポート
- **特徴**:
  - ユニット NO.と部品番号を含む詳細情報
  - レスポンシブデザインでスマートフォンでも閲覧可能
  - 価格情報の視覚的強調表示

## 注意事項

- 仮想環境を使用することで、システム全体の Python 環境に影響を与えることなく、プロジェクト固有のライブラリを管理できます
- プログラムを実行する前に、必ず仮想環境をアクティベートしてください
- 入力ファイルは`SampleData`ディレクトリ内に配置してください
- 出力ファイルは`ReportOutput`ディレクトリに自動生成されます
- JSON ファイルは分析・グラフ作成・AI 予測に最適化されています
- データ分析ユーティリティ（`data_analyzer.py`）で簡単に分析可能です
- HTML レポートは自動的にブラウザで開かれます

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
