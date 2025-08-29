#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
仕入レポート生成プログラム
社内システムから出力されたExcelファイルを成形し、仕入レポートを生成する
"""

import pandas as pd
import os
import json
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

# 分類置換テーブル（分類コード -> 置換名称）
CATEGORY_MAPPING = {
    '02': 'E:盤組',
    '03': 'E:配線',
    '04': 'E:調整',
    '05': 'E:配線',
    '06': 'M:設計',
    '07': 'M:製作',
    '08': 'M:組立',
    '09': 'M:組立',
    '10': 'M:組立',
    '11': 'E:部品',
    '12': 'E:部品',
    '13': 'E:部品',
    '14': 'M:一式',
    '15': 'M:購入',
    '16': 'M:材料',
    '17': 'M:製作',
    '18': 'M:一式',
    '19': '-',
    '20': 'Others:',
    '100': 'S:旅費',
    '101': 'E:旅費',
    '102': 'M:旅費',
    '103': 'S:旅費',
    '104': 'S:旅費'
}

# Excel出力パターンの列定義とマッピング情報
EXCEL_OUTPUT_COLUMNS = [
    {
        'column': 'A',
        'title': '分類コード',
        'description': '分類コード（整数として表示）',
        'source_keywords': ['分類', 'コード', '分類ｺｰﾄﾞ'],
        'data_type': 'int',
        'transformation': 'safe_int_convert_category'
    },
    {
        'column': 'B',
        'title': '分類名称',
        'description': '分類名称（置換後）',
        'source_keywords': ['分類名称'],
        'data_type': 'string',
        'transformation': 'category_mapping'
    },
    {
        'column': 'C',
        'title': '仕入先コード',
        'description': '仕入先コード',
        'source_keywords': ['仕入先ｺｰﾄﾞ'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'D',
        'title': '仕入先',
        'description': '仕入先名称',
        'source_keywords': ['仕入先略称'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'E',
        'title': 'ファイルNo.',
        'description': 'ファイル番号',
        'source_keywords': ['ﾌｧｲﾙNO', 'ファイル', 'NO'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'F',
        'title': 'UNIT',
        'description': 'ユニット番号',
        'source_keywords': ['ﾕﾆｯﾄNO', 'UNIT', 'ユニット'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'G',
        'title': 'No.',
        'description': '部品番号（整数として表示）',
        'source_keywords': ['部品番号'],
        'data_type': 'int',
        'transformation': 'safe_int_convert'
    },
    {
        'column': 'H',
        'title': '品名',
        'description': '品目名称',
        'source_keywords': ['品目名称', '品名'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'I',
        'title': 'メーカー',
        'description': 'メーカー名',
        'source_keywords': ['ﾒｰｶｰ名', 'メーカー'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'J',
        'title': '材質・型式',
        'description': '材質・型式',
        'source_keywords': ['材質・型式', '材質', '型式'],
        'data_type': 'string',
        'transformation': 'direct_copy'
    },
    {
        'column': 'K',
        'title': '数',
        'description': '受入数量（単位なし、整数）',
        'source_keywords': ['受入数量', '数量', '数'],
        'data_type': 'int',
        'transformation': 'quantity_convert'
    },
    {
        'column': 'L',
        'title': '受入日',
        'description': '受入日（納入日の日付データをそのまま）',
        'source_keywords': ['納入日', '受入日'],
        'data_type': 'date',
        'transformation': 'direct_copy'
    },
    {
        'column': 'M',
        'title': '単価',
        'description': '受入単価（数値として表示）',
        'source_keywords': ['受入単価', '単価'],
        'data_type': 'float',
        'transformation': 'price_convert'
    }
]

# 変換ロジックの説明
TRANSFORMATION_LOGIC = {
    'safe_int_convert_category': '数値に変換可能なもののみ変換、それ以外は0',
    'category_mapping': '分類コードを2桁にゼロパディングし、CATEGORY_MAPPINGで置換',
    'direct_copy': '元データをそのままコピー',
    'safe_int_convert': '数値に変換可能なもののみ変換、それ以外は0（部品番号用）',
    'quantity_convert': 'NaNを0に変換し、整数として表示',
    'price_convert': 'NaNを0に変換し、数値として表示'
}

class PurchaseReportGenerator:
    """仕入レポート生成クラス"""
    
    def __init__(self, output_dir="ReportOutput"):
        """
        初期化
        
        Args:
            output_dir (str): 出力ディレクトリのパス
        """
        self.output_dir = Path(output_dir)
        self.original_data = None
        self.category_mapping = None
        
        # 出力ディレクトリが存在しない場合は作成
        self.output_dir.mkdir(exist_ok=True)
    
    def select_file_dialog(self, title="ファイルを選択してください", file_types=None):
        """
        ファイル選択ダイアログを表示
        
        Args:
            title (str): ダイアログのタイトル
            file_types (list): ファイルタイプのリスト [("説明", "拡張子"), ...]
        
        Returns:
            str: 選択されたファイルのパス、キャンセルされた場合はNone
        """
        # Tkinterのルートウィンドウを作成（非表示）
        root = tk.Tk()
        root.withdraw()  # メインウィンドウを非表示
        
        # ファイル選択ダイアログを表示
        if file_types is None:
            file_types = [
                ("Excelファイル", "*.xlsx *.xls"),
                ("すべてのファイル", "*.*")
            ]
        
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=file_types,
            initialdir=os.getcwd()
        )
        
        # ルートウィンドウを破棄
        root.destroy()
        
        return file_path if file_path else None
    
    def load_original_data(self, file_path=None):
        """
        オリジナルデータを読み込む
        
        Args:
            file_path (str, optional): 読み込むファイルのパス。Noneの場合はダイアログで選択
        
        Returns:
            pandas.DataFrame: 読み込んだデータ
        """
        if file_path is None:
            # ファイル選択ダイアログを表示
            file_path = self.select_file_dialog(
                title="オリジナルデータファイルを選択してください",
                file_types=[
                    ("Excelファイル", "*.xlsx *.xls"),
                    ("すべてのファイル", "*.*")
                ]
            )
            
            if file_path is None:
                raise ValueError("ファイルが選択されませんでした")
        
        file_path = Path(file_path)
        
        print(f"オリジナルデータを読み込み中: {file_path}")
        
        try:
            # ファイル拡張子に応じて読み込み方法を変更
            if file_path.suffix.lower() == '.xls':
                # .xlsファイルの場合
                self.original_data = pd.read_excel(file_path, engine='xlrd')
                # 列名の文字化けを修正
                self.original_data.columns = self.original_data.columns.str.encode('latin1').str.decode('shift_jis', errors='ignore')
                
                # 文字列データの文字化けを修正
                for col in self.original_data.select_dtypes(include=['object']).columns:
                    self.original_data[col] = self.original_data[col].astype(str).str.encode('latin1').str.decode('shift_jis', errors='ignore')
            else:
                # .xlsxファイルの場合
                self.original_data = pd.read_excel(file_path, engine='openpyxl')
            
            print(f"データ読み込み完了: {len(self.original_data)}行")
            print(f"列名: {list(self.original_data.columns)}")
            
            return self.original_data
            
        except Exception as e:
            print(f"データ読み込みエラー: {e}")
            raise
    

    
    def load_category_mapping(self, filename=None):
        """
        分類置換テーブルを読み込む（非推奨 - 内部定義を使用）
        
        Args:
            filename (str, optional): 読み込むファイル名。Noneの場合は自動検索
        
        Returns:
            dict: 分類置換マッピング
        """
        print("分類置換テーブルは内部定義を使用します")
        return CATEGORY_MAPPING
    

    
    def apply_category_mapping(self, data):
        """
        分類置換テーブルを適用して分類名称を置換
        
        Args:
            data (pandas.DataFrame): 対象データ
        
        Returns:
            pandas.DataFrame: 分類名称が置換されたデータ
        """
        print("分類置換テーブルを適用中...")
        
        # 分類コードを元に置換名称を適用
        data_copy = data.copy()
        if '分類ｺｰﾄﾞ' in data_copy.columns:
            # 分類コードを文字列に変換して2桁にゼロパディング
            data_copy['分類名称_置換後'] = data_copy['分類ｺｰﾄﾞ'].astype(str).str.zfill(2).map(CATEGORY_MAPPING).fillna(data_copy['分類名称'])
            print(f"分類名称の置換完了: {len(CATEGORY_MAPPING)}件のマッピングを適用")
        
        return data_copy
    
    def filter_data(self, data):
        """
        データを処理（全データを処理）
        
        Args:
            data (pandas.DataFrame): 対象データ
        
        Returns:
            pandas.DataFrame: 処理されたデータ
        """
        print("データ処理中...")
        
        # 全データを処理
        filtered_data = data.copy()
        
        print(f"処理前: {len(data)}行")
        print(f"処理後: {len(filtered_data)}行")
        
        return filtered_data
    
    def export_data_to_json(self, data, filename=None):
        """
        データをJSONファイルに出力（分析用に最適化）
        
        Args:
            data (pandas.DataFrame): 出力するデータ
            filename (str, optional): 出力ファイル名。Noneの場合は自動生成
        
        Returns:
            str: 出力されたファイルのパス
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"purchase_report_{timestamp}.json"
        
        file_path = self.output_dir / filename
        
        # データ型情報を取得
        dtype_info = {}
        for col in data.columns:
            dtype_info[col] = str(data[col].dtype)
        
        # 基本統計情報を計算
        numeric_columns = data.select_dtypes(include=['number']).columns
        statistics = {}
        for col in numeric_columns:
            statistics[col] = {
                'mean': float(data[col].mean()) if not data[col].isna().all() else None,
                'std': float(data[col].std()) if not data[col].isna().all() else None,
                'min': float(data[col].min()) if not data[col].isna().all() else None,
                'max': float(data[col].max()) if not data[col].isna().all() else None,
                'count': int(data[col].count())
            }
        
        # カテゴリ変数の基本情報
        categorical_columns = data.select_dtypes(include=['object']).columns
        categorical_info = {}
        for col in categorical_columns:
            value_counts = data[col].value_counts().head(10).to_dict()
            categorical_info[col] = {
                'unique_count': int(data[col].nunique()),
                'top_values': {str(k): int(v) for k, v in value_counts.items()}
            }
        
        # DataFrameをJSON形式に変換
        json_data = {
            'metadata': {
                'generated_at': datetime.now().isoformat(),
                'total_records': len(data),
                'columns': list(data.columns),
                'data_types': dtype_info,
                'file_no': data['ﾌｧｲﾙNO'].iloc[0] if len(data) > 0 else None
            },
            'statistics': {
                'numeric_columns': statistics,
                'categorical_columns': categorical_info
            },
            'data': data.to_dict('records')
        }
        
        # JSONファイルに出力
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        
        print(f"JSONファイルを出力しました: {file_path}")
        return str(file_path)
    
    def export_data_to_csv(self, data, filename=None):
        """
        データをCSVファイルに出力
        
        Args:
            data (pandas.DataFrame): 出力するデータ
            filename (str, optional): 出力ファイル名。Noneの場合は自動生成
        
        Returns:
            str: 出力されたファイルのパス
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"purchase_report_{timestamp}.csv"
        
        file_path = self.output_dir / filename
        
        # CSVファイルに出力（UTF-8 BOM付きでExcel対応）
        data.to_csv(file_path, index=False, encoding='utf-8-sig')
        
        print(f"CSVファイルを出力しました: {file_path}")
        return str(file_path)
    
    def export_summary_to_json(self, category_summary, file_summary, filename=None):
        """
        集計データをJSONファイルに出力
        
        Args:
            category_summary (pandas.DataFrame): 分類別集計データ
            file_summary (pandas.DataFrame): ファイル別集計データ
            filename (str, optional): 出力ファイル名。Noneの場合は自動生成
        
        Returns:
            str: 出力されたファイルのパス
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"purchase_summary_{timestamp}.json"
        
        file_path = self.output_dir / filename
        
        # 集計データをJSON形式に変換
        summary_data = {
            'metadata': {
                'generated_at': datetime.now().isoformat(),
                'category_count': len(category_summary),
                'file_count': len(file_summary)
            },
            'category_summary': category_summary.to_dict('records'),
            'file_summary': file_summary.to_dict('records')
        }
        
        # JSONファイルに出力
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(summary_data, f, ensure_ascii=False, indent=2)
        
        print(f"集計JSONファイルを出力しました: {file_path}")
        return str(file_path)
    
    def export_to_excel_format(self, filtered_data, category_summary, file_summary, filename=None):
        """
        画像の列構成に合わせてExcelファイルに出力
        
        Args:
            filtered_data (pandas.DataFrame): フィルタリングされたデータ
            category_summary (pandas.DataFrame): 分類別集計データ（未使用）
            file_summary (pandas.DataFrame): ファイル別集計データ（未使用）
            filename (str, optional): 出力ファイル名。Noneの場合は自動生成
        
        Returns:
            str: 出力されたファイルのパス
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"purchase_report_{timestamp}.xlsx"
        
        file_path = self.output_dir / filename
        
        # 画像の列構成に合わせてデータを整形
        formatted_data = self._format_data_for_excel(filtered_data)
        
        # Excelファイルに出力
        formatted_data.to_excel(file_path, index=False, sheet_name='20250825_オリジナルデータ')
        
        print(f"Excelファイルを出力しました: {file_path}")
        return str(file_path)
    
    def _find_column_by_keywords(self, df, keywords):
        """
        キーワードに基づいて列名を検索する
        
        Args:
            df (pandas.DataFrame): 検索対象のDataFrame
            keywords (list): 検索キーワードのリスト
        
        Returns:
            str: 見つかった列名、見つからない場合はNone
        """
        # 完全一致を優先
        for col in df.columns:
            col_str = str(col)
            for keyword in keywords:
                if keyword == col_str:
                    return col
        
        # 部分一致を試行
        for col in df.columns:
            col_lower = str(col).lower()
            for keyword in keywords:
                if keyword.lower() in col_lower:
                    return col
        
        return None

    def _format_data_for_excel(self, filtered_data):
        """
        EXCEL_OUTPUT_COLUMNSの定義に基づいてデータを整形
        
        Args:
            filtered_data (pandas.DataFrame): フィルタリングされたデータ
        
        Returns:
            pandas.DataFrame: 整形されたデータ
        """
        print("Excel出力形式にデータを整形中...")
        formatted_data = pd.DataFrame()
        
        # 各列の定義に基づいてデータを処理
        for col_def in EXCEL_OUTPUT_COLUMNS:
            column_title = col_def['title']
            source_keywords = col_def['source_keywords']
            transformation = col_def['transformation']
            data_type = col_def['data_type']
            
            print(f"列 {col_def['column']} ({column_title}): {col_def['description']}")
            
            # ソース列を検索
            source_col = self._find_column_by_keywords(filtered_data, source_keywords)
            
            if source_col:
                print(f"  ソース列: {source_col}")
                
                # 変換ロジックを適用
                if transformation == 'safe_int_convert_category':
                    def safe_int_convert_category(x):
                        try:
                            if pd.isna(x):
                                return 0
                            return int(float(x))
                        except (ValueError, TypeError):
                            return 0
                    
                    formatted_data[column_title] = filtered_data[source_col].apply(safe_int_convert_category)
                
                elif transformation == 'category_mapping':
                    # 分類コードを文字列に変換して2桁にゼロパディングし、分類置換テーブルを適用
                    def safe_int_convert_category(x):
                        try:
                            if pd.isna(x):
                                return 0
                            return int(float(x))
                        except (ValueError, TypeError):
                            return 0
                    
                    # 分類コード列を検索（A列と同じソースを使用）
                    category_code_col = self._find_column_by_keywords(filtered_data, ['分類', 'コード', '分類ｺｰﾄﾞ'])
                    if category_code_col:
                        category_codes = filtered_data[category_code_col].apply(safe_int_convert_category).astype(str).str.zfill(2)
                        formatted_data[column_title] = category_codes.map(CATEGORY_MAPPING).fillna('')
                        print(f"    分類コード列 '{category_code_col}' を使用して分類名称を生成")
                    else:
                        print(f"    警告: 分類コード列が見つかりません")
                        formatted_data[column_title] = ''
                
                elif transformation == 'direct_copy':
                    formatted_data[column_title] = filtered_data[source_col].fillna('')
                
                elif transformation == 'safe_int_convert':
                    def safe_int_convert(x):
                        try:
                            if pd.isna(x):
                                return 0
                            return int(float(x))
                        except (ValueError, TypeError):
                            return 0
                    
                    formatted_data[column_title] = filtered_data[source_col].apply(safe_int_convert)
                
                elif transformation == 'quantity_convert':
                    formatted_data[column_title] = filtered_data[source_col].fillna(0).astype(int)
                
                elif transformation == 'price_convert':
                    formatted_data[column_title] = filtered_data[source_col].fillna(0)
                
                else:
                    print(f"  警告: 未定義の変換ロジック '{transformation}' を使用")
                    formatted_data[column_title] = filtered_data[source_col].fillna('')
            
            else:
                print(f"  警告: ソース列が見つかりません。キーワード: {source_keywords}")
                # デフォルト値を設定
                if data_type == 'int':
                    formatted_data[column_title] = 0
                elif data_type == 'float':
                    formatted_data[column_title] = 0.0
                else:
                    formatted_data[column_title] = ''
        
        print(f"データ整形完了: {len(formatted_data)}行、{len(formatted_data.columns)}列")
        return formatted_data
    
    def display_data_info(self):
        """データの基本情報を表示"""
        if self.original_data is None:
            print("オリジナルデータが読み込まれていません")
            return
        
        print("\n=== オリジナルデータ情報 ===")
        print(f"行数: {len(self.original_data)}")
        print(f"列数: {len(self.original_data.columns)}")
        print(f"列名: {list(self.original_data.columns)}")
        print("\n最初の5行:")
        print(self.original_data.head())
        print("\nデータ型:")
        print(self.original_data.dtypes)
        
        print("\n=== 分類置換テーブル情報 ===")
        print(f"内部定義マッピング数: {len(CATEGORY_MAPPING)}")
        print("分類コード -> 置換名称:")
        for code, name in sorted(CATEGORY_MAPPING.items()):
            print(f"  {code} -> {name}")
        
        print("\n=== Excel出力列定義情報 ===")
        print(f"出力列数: {len(EXCEL_OUTPUT_COLUMNS)}")
        print("列定義:")
        for col_def in EXCEL_OUTPUT_COLUMNS:
            print(f"  {col_def['column']}列: {col_def['title']}")
            print(f"    説明: {col_def['description']}")
            print(f"    ソースキーワード: {col_def['source_keywords']}")
            print(f"    データ型: {col_def['data_type']}")
            print(f"    変換ロジック: {col_def['transformation']} ({TRANSFORMATION_LOGIC.get(col_def['transformation'], '未定義')})")
            print()
        
        print("=== 変換ロジック詳細 ===")
        for logic_name, description in TRANSFORMATION_LOGIC.items():
            print(f"  {logic_name}: {description}")

def main():
    """メイン関数"""
    print("仕入レポート生成プログラムを開始します")
    print("ファイル選択ダイアログが表示されます。")
    
    # レポート生成器のインスタンスを作成
    generator = PurchaseReportGenerator()
    
    try:
        # オリジナルデータを読み込み
        print("\n=== ステップ1: オリジナルデータの読み込み ===")
        original_data = generator.load_original_data()
        
        # 分類置換テーブル情報を表示
        print("\n=== ステップ2: 分類置換テーブル情報 ===")
        generator.load_category_mapping()
        
        # データ情報を表示
        print("\n=== ステップ3: データ情報の表示 ===")
        generator.display_data_info()
        
        # 分類置換テーブルを適用
        print("\n=== ステップ4: 分類置換テーブルの適用 ===")
        processed_data = generator.apply_category_mapping(original_data)
        
        # データを処理
        print("\n=== ステップ5: データの処理 ===")
        filtered_data = generator.filter_data(processed_data)
        
        # 処理結果を表示
        print("\n=== ステップ6: 処理結果の表示 ===")
        print(f"処理後のデータ行数: {len(filtered_data)}")
        
        # 分類別の集計
        print("\n=== 分類別集計 ===")
        category_summary = filtered_data.groupby(['分類ｺｰﾄﾞ', '分類名称_置換後'])['受入金額'].agg(['count', 'sum']).reset_index()
        category_summary.columns = ['分類コード', '分類名称（置換後）', '件数', '合計金額']
        print(category_summary)
        
        # ファイル別の集計
        print("\n=== ファイル別集計 ===")
        file_summary = filtered_data.groupby('ﾌｧｲﾙNO')['受入金額'].agg(['count', 'sum']).reset_index()
        file_summary.columns = ['ファイルNO', '件数', '合計金額']
        print(file_summary)
        
        print("\n処理後のデータ（最初の10行）:")
        print(filtered_data[['分類ｺｰﾄﾞ', '分類名称', '分類名称_置換後', 'ﾌｧｲﾙNO', '受入金額']].head(10))
        
        # データ出力
        print("\n=== ステップ7: データ出力 ===")
        
        # 詳細データをJSONで出力（分析用に最適化）
        json_file = generator.export_data_to_json(filtered_data)
        
        # 集計データをJSONで出力
        summary_file = generator.export_summary_to_json(category_summary, file_summary)
        
        # Excelファイルを出力（指定フォーマット）
        excel_file = generator.export_to_excel_format(filtered_data, category_summary, file_summary)
        
        print("\n=== 処理完了 ===")
        print("オリジナルデータから分類置換テーブルの適用が完了しました。")
        print("データが以下のファイルに出力されました：")
        print(f"  - 詳細データ（JSON）: {json_file}")
        print(f"  - 集計データ（JSON）: {summary_file}")
        print(f"  - Excelファイル: {excel_file}")
        print("次のステップでは、このデータを基に分析・グラフ作成・AI予測を行います。")
        
        # 完了メッセージを表示
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("処理完了", f"仕入レポートの生成が完了しました。\n\n出力ファイル:\n{json_file}\n{summary_file}\n{excel_file}")
        root.destroy()
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()
        
        # エラーメッセージを表示
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("エラー", f"処理中にエラーが発生しました:\n{e}")
        root.destroy()

if __name__ == "__main__":
    main()
