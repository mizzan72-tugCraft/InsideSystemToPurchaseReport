#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
データ分析ユーティリティ
JSONファイルを読み込んで分析・グラフ作成・AI予測に使用する
"""

import json
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

class DataAnalyzer:
    """データ分析クラス"""
    
    def __init__(self, json_file_path):
        """
        初期化
        
        Args:
            json_file_path (str): JSONファイルのパス
        """
        self.json_file_path = Path(json_file_path)
        self.data = None
        self.metadata = None
        self.statistics = None
        self.df = None
        
        # JSONファイルを読み込み
        self.load_json_data()
    
    def load_json_data(self):
        """JSONファイルを読み込む"""
        try:
            with open(self.json_file_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            self.metadata = json_data.get('metadata', {})
            self.statistics = json_data.get('statistics', {})
            self.data = json_data.get('data', [])
            
            # DataFrameに変換
            self.df = pd.DataFrame(self.data)
            
            print(f"データ読み込み完了: {len(self.df)}行, {len(self.df.columns)}列")
            print(f"ファイルNO: {self.metadata.get('file_no', 'Unknown')}")
            
        except Exception as e:
            print(f"JSONファイル読み込みエラー: {e}")
            raise
    
    def get_basic_info(self):
        """基本情報を取得"""
        return {
            'file_no': self.metadata.get('file_no'),
            'total_records': len(self.df),
            'columns': list(self.df.columns),
            'data_types': self.metadata.get('data_types', {}),
            'generated_at': self.metadata.get('generated_at')
        }
    
    def get_numeric_statistics(self):
        """数値列の統計情報を取得"""
        return self.statistics.get('numeric_columns', {})
    
    def get_categorical_info(self):
        """カテゴリ変数の情報を取得"""
        return self.statistics.get('categorical_columns', {})
    
    def get_dataframe(self):
        """pandas DataFrameを取得"""
        return self.df
    
    def filter_by_category(self, category_name):
        """分類名称でフィルタリング"""
        if '分類名称_置換後' in self.df.columns:
            return self.df[self.df['分類名称_置換後'] == category_name]
        return pd.DataFrame()
    
    def get_category_summary(self):
        """分類別の集計を取得"""
        if '分類名称_置換後' in self.df.columns and '受入金額' in self.df.columns:
            return self.df.groupby('分類名称_置換後')['受入金額'].agg(['count', 'sum', 'mean']).reset_index()
        return pd.DataFrame()
    
    def get_supplier_summary(self):
        """仕入先別の集計を取得"""
        if '仕入先略称' in self.df.columns and '受入金額' in self.df.columns:
            return self.df.groupby('仕入先略称')['受入金額'].agg(['count', 'sum', 'mean']).reset_index()
        return pd.DataFrame()
    
    def get_monthly_summary(self):
        """月別の集計を取得"""
        if '受入日' in self.df.columns and '受入金額' in self.df.columns:
            # 受入日を日付型に変換
            self.df['受入日'] = pd.to_datetime(self.df['受入日'], errors='coerce')
            self.df['受入月'] = self.df['受入日'].dt.strftime('%Y-%m')
            
            return self.df.groupby('受入月')['受入金額'].agg(['count', 'sum', 'mean']).reset_index()
        return pd.DataFrame()
    
    def export_analysis_results(self, output_dir="ReportOutput"):
        """分析結果をJSONで出力"""
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"analysis_results_{timestamp}.json"
        file_path = output_path / filename
        
        # 分析結果をまとめる
        analysis_results = {
            'metadata': {
                'analyzed_at': datetime.now().isoformat(),
                'source_file': str(self.json_file_path),
                'file_no': self.metadata.get('file_no')
            },
            'basic_info': self.get_basic_info(),
            'category_summary': self.get_category_summary().to_dict('records'),
            'supplier_summary': self.get_supplier_summary().to_dict('records'),
            'monthly_summary': self.get_monthly_summary().to_dict('records'),
            'statistics': self.statistics
        }
        
        # JSONファイルに出力
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(analysis_results, f, ensure_ascii=False, indent=2)
        
        print(f"分析結果を出力しました: {file_path}")
        return str(file_path)

def main():
    """テスト用メイン関数"""
    # 最新のJSONファイルを自動検索
    output_dir = Path("ReportOutput")
    json_files = list(output_dir.glob("purchase_report_*.json"))
    
    if not json_files:
        print("JSONファイルが見つかりません")
        return
    
    # 最新のファイルを使用
    latest_file = max(json_files, key=lambda x: x.stat().st_mtime)
    print(f"分析対象ファイル: {latest_file}")
    
    # データ分析を実行
    analyzer = DataAnalyzer(latest_file)
    
    # 基本情報を表示
    print("\n=== 基本情報 ===")
    basic_info = analyzer.get_basic_info()
    print(f"ファイルNO: {basic_info['file_no']}")
    print(f"レコード数: {basic_info['total_records']}")
    print(f"列数: {len(basic_info['columns'])}")
    
    # 分類別集計を表示
    print("\n=== 分類別集計 ===")
    category_summary = analyzer.get_category_summary()
    print(category_summary)
    
    # 仕入先別集計を表示
    print("\n=== 仕入先別集計 ===")
    supplier_summary = analyzer.get_supplier_summary()
    print(supplier_summary.head(10))
    
    # 分析結果を出力
    analyzer.export_analysis_results()

if __name__ == "__main__":
    main()
