#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
出力ファイル確認スクリプト
ReportOutputディレクトリ内のファイルを確認する
"""

import json
import pandas as pd
from pathlib import Path

def check_output_files():
    """出力ファイルを確認"""
    output_dir = Path("ReportOutput")
    
    if not output_dir.exists():
        print("ReportOutputディレクトリが見つかりません")
        return
    
    print("=== ReportOutputディレクトリの内容 ===")
    
    # ディレクトリ内のファイルを一覧表示
    files = list(output_dir.glob("*"))
    for file in files:
        if file.is_file():
            print(f"ファイル: {file.name} ({file.stat().st_size:,} bytes)")
    
    print("\n=== 集計ファイルの内容確認 ===")
    
    # 集計JSONファイルを確認
    summary_files = list(output_dir.glob("*summary*.json"))
    for file in summary_files:
        print(f"\n--- {file.name} ---")
        with open(file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        print(f"生成日時: {data['metadata']['generated_at']}")
        print(f"分類数: {data['metadata']['category_count']}")
        print(f"ファイル数: {data['metadata']['file_count']}")
        
        print("\n分類別集計:")
        for category in data['category_summary']:
            print(f"  {category['分類名称（置換後）']}: {category['件数']}件, {category['合計金額']:,}円")
        
        print("\nファイル別集計:")
        for file_summary in data['file_summary']:
            print(f"  {file_summary['ファイルNO']}: {file_summary['件数']}件, {file_summary['合計金額']:,}円")
    
    print("\n=== 詳細データファイルの内容確認 ===")
    
    # CSVファイルを確認
    csv_files = list(output_dir.glob("*.csv"))
    for file in csv_files:
        print(f"\n--- {file.name} ---")
        df = pd.read_csv(file, encoding='utf-8-sig')
        print(f"行数: {len(df)}")
        print(f"列数: {len(df.columns)}")
        print(f"列名: {list(df.columns)}")
        
        # 分類別の件数を表示
        if '分類名称_置換後' in df.columns:
            print("\n分類別件数:")
            category_counts = df['分類名称_置換後'].value_counts()
            for category, count in category_counts.items():
                print(f"  {category}: {count}件")
        
        # 合計金額を表示
        if '受入金額' in df.columns:
            total_amount = df['受入金額'].sum()
            print(f"\n合計金額: {total_amount:,}円")

if __name__ == "__main__":
    check_output_files()
