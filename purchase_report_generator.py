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

class PurchaseReportGenerator:
    """仕入レポート生成クラス"""
    
    def __init__(self, sample_data_dir="SampleData", output_dir="ReportOutput"):
        """
        初期化
        
        Args:
            sample_data_dir (str): サンプルデータのディレクトリパス
            output_dir (str): 出力ディレクトリのパス
        """
        self.sample_data_dir = Path(sample_data_dir)
        self.output_dir = Path(output_dir)
        self.original_data = None
        self.molding_list = None
        self.category_mapping = None
        
        # 出力ディレクトリが存在しない場合は作成
        self.output_dir.mkdir(exist_ok=True)
        
    def load_original_data(self, filename=None):
        """
        オリジナルデータを読み込む
        
        Args:
            filename (str, optional): 読み込むファイル名。Noneの場合は自動検索
        
        Returns:
            pandas.DataFrame: 読み込んだデータ
        """
        if filename is None:
            # オリジナルデータファイルを自動検索
            original_files = list(self.sample_data_dir.glob("*_オリジナルデータ.*"))
            if not original_files:
                raise FileNotFoundError("オリジナルデータファイルが見つかりません")
            file_path = original_files[0]
        else:
            file_path = self.sample_data_dir / filename
        
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
    
    def load_molding_list(self, filename=None):
        """
        成形リストを読み込む
        
        Args:
            filename (str, optional): 読み込むファイル名。Noneの場合は自動検索
        
        Returns:
            pandas.DataFrame: 読み込んだデータ
        """
        if filename is None:
            # 成形リストファイルを自動検索
            molding_files = list(self.sample_data_dir.glob("*_成形リスト.*"))
            if not molding_files:
                raise FileNotFoundError("成形リストファイルが見つかりません")
            file_path = molding_files[0]
        else:
            file_path = self.sample_data_dir / filename
        
        print(f"成形リストを読み込み中: {file_path}")
        
        try:
            # ファイル拡張子に応じて読み込み方法を変更
            if file_path.suffix.lower() == '.xls':
                # .xlsファイルの場合
                self.molding_list = pd.read_excel(file_path, engine='xlrd')
                # 列名の文字化けを修正
                self.molding_list.columns = self.molding_list.columns.str.encode('latin1').str.decode('shift_jis', errors='ignore')
                
                # 文字列データの文字化けを修正
                for col in self.molding_list.select_dtypes(include=['object']).columns:
                    self.molding_list[col] = self.molding_list[col].astype(str).str.encode('latin1').str.decode('shift_jis', errors='ignore')
            else:
                # .xlsxファイルの場合
                self.molding_list = pd.read_excel(file_path, engine='openpyxl')
            
            print(f"成形リスト読み込み完了: {len(self.molding_list)}行")
            print(f"列名: {list(self.molding_list.columns)}")
            
            return self.molding_list
            
        except Exception as e:
            print(f"成形リスト読み込みエラー: {e}")
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
    
    def process_molding_list(self):
        """
        成形リストを処理して、必要な列とフィルタ条件を抽出
        
        Returns:
            dict: 成形リストの処理結果
        """
        if self.molding_list is None:
            raise ValueError("成形リストが読み込まれていません")
        
        print("成形リストを処理中...")
        
        # 成形リストの構造を確認
        print("成形リストの列名:", list(self.molding_list.columns))
        print("成形リストの最初の10行:")
        print(self.molding_list.head(10))
        
        # 成形リストから必要な情報を抽出
        # 分類コードと置換名称の列を探す
        category_code_col = None
        replacement_name_col = None
        
        # 各列の内容を確認して分類コードと置換名称の列を特定
        for col in self.molding_list.columns:
            col_values = self.molding_list[col].astype(str)
            if any('分類' in val and 'コード' in val for val in col_values):
                category_code_col = col
            elif any('置換' in val and '名称' in val for val in col_values):
                replacement_name_col = col
        
        print(f"分類コード列: {category_code_col}")
        print(f"置換名称列: {replacement_name_col}")
        
        if category_code_col is None or replacement_name_col is None:
            # 列名が特定できない場合は、データの内容から推測
            for i, row in self.molding_list.iterrows():
                row_str = str(row.values)
                if '分類' in row_str and 'コード' in row_str:
                    # この行がヘッダー行
                    print(f"ヘッダー行を発見: 行 {i}")
                    print(f"ヘッダー行の内容: {row.values}")
                    
                    # 分類コードと置換名称の列インデックスを特定
                    for j, val in enumerate(row.values):
                        if '分類' in str(val) and 'コード' in str(val):
                            category_code_col = self.molding_list.columns[j]
                        elif '置換' in str(val) and '名称' in str(val):
                            replacement_name_col = self.molding_list.columns[j]
                    
                    # ヘッダー行以降のデータを取得
                    molding_data = self.molding_list.iloc[i+1:].copy()
                    break
        
        if category_code_col is None or replacement_name_col is None:
            raise ValueError("成形リストで分類コードまたは置換名称の列が見つかりません")
        
        # 分類コードと置換名称のマッピングを作成
        category_mapping = {}
        
        # 成形リストの各行を確認して、有効な分類コードと置換名称を抽出
        for i, row in self.molding_list.iterrows():
            category_code = row[category_code_col]
            replacement_name = row[replacement_name_col]
            
            # 分類コードが数値で、置換名称が有効な場合のみ処理
            if pd.notna(category_code) and pd.notna(replacement_name):
                try:
                    # 分類コードを数値として確認
                    category_code_int = int(category_code)
                    category_code_str = str(category_code_int).zfill(2)
                    
                    # 置換名称が有効な文字列の場合のみ追加
                    replacement_str = str(replacement_name).strip()
                    if replacement_str and replacement_str != 'nan' and replacement_str != '―':
                        # 既に存在する場合は上書きしない
                        if category_code_str not in category_mapping:
                            category_mapping[category_code_str] = replacement_str
                            print(f"分類コード {category_code_str} -> {replacement_str}")
                except (ValueError, TypeError):
                    # 分類コードが数値でない場合はスキップ
                    continue
        
        print(f"有効な分類マッピング: {category_mapping}")
        
        # すべての分類コードを含める（成形リストの制限を無効化）
        print("注意: 成形リストの制限を無効化し、すべてのデータを処理します")
        
        return {
            'category_mapping': category_mapping,
            'molding_data': self.molding_list
        }
    
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
    
    def filter_data_by_molding_list(self, data, molding_info):
        """
        成形リストに基づいてデータを絞り込み（現在は全データを処理）
        
        Args:
            data (pandas.DataFrame): 対象データ
            molding_info (dict): 成形リストの処理結果
        
        Returns:
            pandas.DataFrame: 絞り込まれたデータ
        """
        print("データ処理中...")
        
        # 現在は全データを処理（成形リストの制限を無効化）
        filtered_data = data.copy()
        
        print(f"処理前: {len(data)}行")
        print(f"処理後: {len(filtered_data)}行")
        print("注意: 成形リストの制限を無効化し、すべてのデータを処理しています")
        
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
    
    def _format_data_for_excel(self, filtered_data):
        """
        画像の列構成に合わせてデータを整形
        
        Args:
            filtered_data (pandas.DataFrame): フィルタリングされたデータ
        
        Returns:
            pandas.DataFrame: 整形されたデータ
        """
        # 画像の列構成に合わせてデータを整形
        formatted_data = pd.DataFrame()
        
        # A列: 分類コード
        formatted_data['分類コード'] = filtered_data['分類ｺｰﾄﾞ'].astype(str).str.zfill(2)
        
        # B列: 分類名称
        formatted_data['分類名称'] = filtered_data['分類名称_置換後']
        
        # C列: 仕入先コード
        formatted_data['仕入先コード'] = filtered_data['仕入先ｺｰﾄﾞ']
        
        # D列: 仕入先
        formatted_data['仕入先'] = filtered_data['仕入先略称']
        
        # E列: ファイルNo.
        formatted_data['ファイルNo.'] = filtered_data['ﾌｧｲﾙNO']
        
        # F列: UNIT
        formatted_data['UNIT'] = filtered_data['ﾕﾆｯﾄNO'].fillna('')
        
        # G列: No. (部品番号) - 整数として表示
        formatted_data['No.'] = filtered_data['部品番号'].fillna(0).astype(int)
        
        # H列: 品名 (品目名称)
        formatted_data['品名'] = filtered_data['品目名称']
        
        # I列: メーカー
        formatted_data['メーカー'] = filtered_data['ﾒｰｶｰ名']
        
        # J列: 材質・型式
        formatted_data['材質・型式'] = filtered_data['材質・型式']
        
        # K列: 数 (受入数量のみ、単位なし)
        formatted_data['数'] = filtered_data['受入数量'].fillna(0).astype(int)
        
        # L列: 受入日 (納入日の日付データをそのまま)
        formatted_data['受入日'] = filtered_data['納入日']
        
        # M列: 単価 (受入単価) - 数値として表示
        formatted_data['単価'] = filtered_data['受入単価'].fillna(0)
        
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
        
        if self.molding_list is not None:
            print("\n=== 成形リスト情報 ===")
            print(f"行数: {len(self.molding_list)}")
            print(f"列数: {len(self.molding_list.columns)}")
            print(f"列名: {list(self.molding_list.columns)}")
            print("\n最初の5行:")
            print(self.molding_list.head())
        
        print("\n=== 分類置換テーブル情報 ===")
        print(f"内部定義マッピング数: {len(CATEGORY_MAPPING)}")
        print("分類コード -> 置換名称:")
        for code, name in sorted(CATEGORY_MAPPING.items()):
            print(f"  {code} -> {name}")

def main():
    """メイン関数"""
    print("仕入レポート生成プログラムを開始します")
    
    # レポート生成器のインスタンスを作成
    generator = PurchaseReportGenerator()
    
    try:
        # オリジナルデータを読み込み
        print("\n=== ステップ1: オリジナルデータの読み込み ===")
        original_data = generator.load_original_data()
        
        # 成形リストを読み込み
        print("\n=== ステップ2: 成形リストの読み込み ===")
        molding_list = generator.load_molding_list()
        
        # 分類置換テーブル情報を表示
        print("\n=== ステップ3: 分類置換テーブル情報 ===")
        generator.load_category_mapping()
        
        # データ情報を表示
        print("\n=== ステップ4: データ情報の表示 ===")
        generator.display_data_info()
        
        # 成形リストを処理
        print("\n=== ステップ5: 成形リストの処理 ===")
        molding_info = generator.process_molding_list()
        
        # 分類置換テーブルを適用
        print("\n=== ステップ6: 分類置換テーブルの適用 ===")
        processed_data = generator.apply_category_mapping(original_data)
        
        # 成形リストに基づいてデータを絞り込み
        print("\n=== ステップ7: データの絞り込み ===")
        filtered_data = generator.filter_data_by_molding_list(processed_data, molding_info)
        
        # 処理結果を表示
        print("\n=== ステップ8: 処理結果の表示 ===")
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
        print("\n=== ステップ9: データ出力 ===")
        
        # 詳細データをJSONで出力（分析用に最適化）
        json_file = generator.export_data_to_json(filtered_data)
        
        # 集計データをJSONで出力
        summary_file = generator.export_summary_to_json(category_summary, file_summary)
        
        # Excelファイルを出力（成形リストフォーマット）
        excel_file = generator.export_to_excel_format(filtered_data, category_summary, file_summary)
        
        print("\n=== 処理完了 ===")
        print("オリジナルデータから成形リストに基づく絞り込みと分類置換テーブルの適用が完了しました。")
        print("データが以下のファイルに出力されました：")
        print(f"  - 詳細データ（JSON）: {json_file}")
        print(f"  - 集計データ（JSON）: {summary_file}")
        print(f"  - Excelファイル: {excel_file}")
        print("次のステップでは、このデータを基に分析・グラフ作成・AI予測を行います。")
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
