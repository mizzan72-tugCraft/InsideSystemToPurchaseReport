#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
仕入レポート分析プログラム
社内システムから出力されたExcelファイルを読み込み、ファイルNo.別の詳細分析を生成する
"""

import pandas as pd
import os
import json
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

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

class PurchaseAnalysisGenerator:
    """仕入レポート分析クラス"""
    
    def __init__(self, output_dir="ReportOutput"):
        """
        初期化
        
        Args:
            output_dir (str): 出力ディレクトリのパス
        """
        self.output_dir = Path(output_dir)
        self.original_data = None
        
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
    
    def select_file_no_dialog(self, file_numbers):
        """
        ファイルNo.選択ダイアログを表示
        
        Args:
            file_numbers (list): 選択可能なファイルNo.のリスト
        
        Returns:
            str: 選択されたファイルNo.、キャンセルされた場合はNone
        """
        # Tkinterのルートウィンドウを作成
        root = tk.Tk()
        root.title("ファイルNo.を選択してください")
        root.geometry("600x400")
        
        # 中央に配置
        root.eval('tk::PlaceWindow . center')
        
        # ラベル
        label = tk.Label(root, text="分析対象のファイルNo.を選択してください:", font=("Arial", 12))
        label.pack(pady=10)
        
        # リストボックス
        listbox_frame = tk.Frame(root)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        listbox = tk.Listbox(listbox_frame, font=("Arial", 11), selectmode=tk.SINGLE)
        scrollbar = tk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # ファイルNo.をリストボックスに追加
        for file_no in sorted(file_numbers):
            listbox.insert(tk.END, file_no)
        
        selected_file_no = [None]  # リストで参照渡し
        
        def on_select():
            selection = listbox.curselection()
            if selection:
                selected_file_no[0] = listbox.get(selection[0])
                root.destroy()
        
        def on_cancel():
            root.destroy()
        
        # ボタンフレーム
        button_frame = tk.Frame(root)
        button_frame.pack(pady=10)
        
        select_button = tk.Button(button_frame, text="選択", command=on_select, font=("Arial", 11))
        select_button.pack(side=tk.LEFT, padx=5)
        
        cancel_button = tk.Button(button_frame, text="キャンセル", command=on_cancel, font=("Arial", 11))
        cancel_button.pack(side=tk.LEFT, padx=5)
        
        # Enterキーで選択
        root.bind('<Return>', lambda e: on_select())
        # Escapeキーでキャンセル
        root.bind('<Escape>', lambda e: on_cancel())
        
        # ダブルクリックで選択
        listbox.bind('<Double-Button-1>', lambda e: on_select())
        
        # ダイアログを表示
        root.focus_set()
        root.grab_set()  # モーダルダイアログ
        root.wait_window()
        
        return selected_file_no[0]
    
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
            # 分類コードを安全に整数に変換してから2桁にゼロパディング
            def safe_int_convert_category(x):
                try:
                    if pd.isna(x):
                        return 0
                    return int(float(x))
                except (ValueError, TypeError):
                    return 0
            
            # 分類コードを整数に変換してから2桁にゼロパディング
            category_codes = data_copy['分類ｺｰﾄﾞ'].apply(safe_int_convert_category).astype(str).str.zfill(2)
            data_copy['分類名称_置換後'] = category_codes.map(CATEGORY_MAPPING).fillna(data_copy['分類名称'])
            print(f"分類名称の置換完了: {len(CATEGORY_MAPPING)}件のマッピングを適用")
            
            # デバッグ情報を表示
            print("分類コード変換例:")
            for i in range(min(5, len(data_copy))):
                original_code = data_copy.iloc[i]['分類ｺｰﾄﾞ']
                converted_code = category_codes.iloc[i]
                mapped_name = data_copy.iloc[i]['分類名称_置換後']
                print(f"  元: {original_code} -> 変換後: {converted_code} -> マッピング: {mapped_name}")
        
        return data_copy
    
    def get_unique_file_numbers(self, data):
        """
        データに含まれるユニークなファイルNo.を取得
        
        Args:
            data (pandas.DataFrame): 対象データ
        
        Returns:
            list: ユニークなファイルNo.のリスト
        """
        if 'ﾌｧｲﾙNO' in data.columns:
            file_numbers = data['ﾌｧｲﾙNO'].dropna().unique().tolist()
            return sorted(file_numbers)
        else:
            return []
    
    def analyze_file_purchases(self, data, file_no):
        """
        指定されたファイルNo.の仕入データを分析
        
        Args:
            data (pandas.DataFrame): 対象データ
            file_no (str): 分析対象のファイルNo.
        
        Returns:
            dict: 分析結果
        """
        print(f"ファイルNo. {file_no} の仕入データを分析中...")
        
        # 指定されたファイルNo.のデータを抽出
        file_data = data[data['ﾌｧｲﾙNO'] == file_no].copy()
        
        if len(file_data) == 0:
            print(f"ファイルNo. {file_no} のデータが見つかりません")
            return None
        
        print(f"分析対象データ: {len(file_data)}行")
        
        # 仕入れ価格を計算（単価 × 数量）
        file_data['仕入れ価格'] = file_data['受入単価'] * file_data['受入数量']
        
        # 階層別に分析
        analysis_result = {
            'file_no': file_no,
            'total_records': len(file_data),
            'total_amount': file_data['仕入れ価格'].sum(),
            'categories': {}
        }
        
        # カテゴリ別にグループ化
        for category, category_group in file_data.groupby('分類名称_置換後'):
            category_total = category_group['仕入れ価格'].sum()
            category_data = {
                'total_amount': category_total,
                'record_count': len(category_group),
                'suppliers': {}
            }
            
            # 仕入先別にグループ化
            for supplier, supplier_group in category_group.groupby('仕入先略称'):
                supplier_total = supplier_group['仕入れ価格'].sum()
                supplier_data = {
                    'total_amount': supplier_total,
                    'record_count': len(supplier_group),
                    'products': []
                }
                
                # 商品別にグループ化（ユニットNO.と部品番号を含む）
                for _, product_row in supplier_group.iterrows():
                     # ユニットNO.の処理（小数点付き数字を文字列に変換）
                     unit_no_raw = product_row.get('ﾕﾆｯﾄNO', '')
                     if pd.isna(unit_no_raw) or unit_no_raw == '' or str(unit_no_raw).lower() == 'nan':
                         unit_no = '-'
                     else:
                         try:
                             # 数値の場合は整数に変換してから文字列化
                             unit_no_int = int(float(unit_no_raw))
                             unit_no = f"{unit_no_int:02d}unit"
                         except (ValueError, TypeError):
                             unit_no = str(unit_no_raw)
                     
                     # 部品番号の処理
                     part_no_raw = product_row.get('部品番号', '')
                     if pd.isna(part_no_raw) or part_no_raw == '' or str(part_no_raw).lower() == 'nan':
                         part_no = '-'
                     else:
                         part_no = str(part_no_raw)
                     
                     # その他の項目の処理（nan値を'-'に変換）
                     def clean_value(value):
                         if pd.isna(value) or value == '' or str(value).lower() == 'nan':
                             return '-'
                         return str(value)
                     
                     product_info = {
                         'unit_no': unit_no,
                         'part_no': part_no,
                         'product_name': clean_value(product_row.get('品目名称', '')),
                         'quantity': product_row.get('受入数量', 0),
                         'unit_price': product_row.get('受入単価', 0),
                         'total_price': product_row.get('仕入れ価格', 0),
                         'manufacturer': clean_value(product_row.get('ﾒｰｶｰ名', '')),
                         'material_type': clean_value(product_row.get('材質・型式', '')),
                         'receive_date': clean_value(product_row.get('納入日', ''))
                     }
                     supplier_data['products'].append(product_info)
                
                # 商品を価格順にソート
                supplier_data['products'].sort(key=lambda x: x['total_price'], reverse=True)
                supplier_data['products'] = supplier_data['products']
                
                supplier_data['products'] = supplier_data['products']
                category_data['suppliers'][supplier] = supplier_data
            
            # 仕入先を価格順にソート
            sorted_suppliers = sorted(
                category_data['suppliers'].items(),
                key=lambda x: x[1]['total_amount'],
                reverse=True
            )
            category_data['suppliers'] = dict(sorted_suppliers)
            
            analysis_result['categories'][category] = category_data
        
        # カテゴリを価格順にソート
        sorted_categories = sorted(
            analysis_result['categories'].items(),
            key=lambda x: x[1]['total_amount'],
            reverse=True
        )
        analysis_result['categories'] = dict(sorted_categories)
        
        print(f"分析完了: {len(analysis_result['categories'])}カテゴリ")
        return analysis_result
    
    def generate_html_report(self, analysis_result, filename=None):
        """
        分析結果をHTMLレポートとして生成
        
        Args:
            analysis_result (dict): 分析結果
            filename (str, optional): 出力ファイル名。Noneの場合は自動生成
        
        Returns:
            str: 出力されたファイルのパス
        """
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"purchase_analysis_{analysis_result['file_no']}_{timestamp}.html"
        
        file_path = self.output_dir / filename
        
        # HTMLテンプレート
        html_content = f"""
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>仕入レポート分析 - {analysis_result['file_no']}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        .header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0;
            font-size: 2.5em;
            font-weight: 300;
        }}
        .header .subtitle {{
            margin-top: 10px;
            font-size: 1.2em;
            opacity: 0.9;
        }}
        .summary {{
            padding: 20px;
            background-color: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
        }}
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }}
        .summary-item {{
            text-align: center;
            padding: 15px;
            background-color: white;
            border-radius: 6px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }}
        .summary-item .value {{
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
        }}
        .summary-item .label {{
            color: #6c757d;
            margin-top: 5px;
        }}
        .content {{
            padding: 20px;
        }}
        .category {{
            margin-bottom: 30px;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            overflow: hidden;
        }}
        .category-header {{
            background-color: #e9ecef;
            padding: 15px 20px;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .category-title {{
            font-size: 1.3em;
            font-weight: bold;
            color: #495057;
        }}
        .category-summary {{
            color: #6c757d;
            font-size: 0.9em;
        }}
        .supplier {{
            margin: 10px;
            border: 1px solid #dee2e6;
            border-radius: 6px;
            overflow: hidden;
        }}
        .supplier-header {{
            background-color: #f8f9fa;
            padding: 12px 15px;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .supplier-title {{
            font-weight: bold;
            color: #495057;
        }}
        .supplier-summary {{
            color: #6c757d;
            font-size: 0.9em;
        }}
        .products-table {{
            width: 100%;
            border-collapse: collapse;
        }}
        .products-table th {{
            background-color: #f8f9fa;
            padding: 10px;
            text-align: left;
            border-bottom: 1px solid #dee2e6;
            font-weight: bold;
            color: #495057;
        }}
        .products-table td {{
            padding: 8px 10px;
            border-bottom: 1px solid #dee2e6;
            vertical-align: top;
        }}
        .products-table tr:hover {{
            background-color: #f8f9fa;
        }}
        .price {{
            font-weight: bold;
            color: #dc3545;
        }}
        .quantity {{
            color: #6c757d;
        }}
        .unit-info {{
            font-size: 0.9em;
            color: #6c757d;
        }}
        .footer {{
            padding: 20px;
            text-align: center;
            color: #6c757d;
            border-top: 1px solid #dee2e6;
            background-color: #f8f9fa;
        }}
        @media (max-width: 768px) {{
            .summary-grid {{
                grid-template-columns: 1fr;
            }}
            .products-table {{
                font-size: 0.9em;
            }}
            .products-table th,
            .products-table td {{
                padding: 5px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>仕入レポート分析</h1>
            <div class="subtitle">ファイルNo.: {analysis_result['file_no']}</div>
        </div>
        
        <div class="summary">
            <h2>概要</h2>
            <div class="summary-grid">
                <div class="summary-item">
                    <div class="value">{analysis_result['total_records']:,}</div>
                    <div class="label">総レコード数</div>
                </div>
                <div class="summary-item">
                    <div class="value">¥{analysis_result['total_amount']:,.0f}</div>
                    <div class="label">総仕入金額</div>
                </div>
                <div class="summary-item">
                    <div class="value">{len(analysis_result['categories'])}</div>
                    <div class="label">カテゴリ数</div>
                </div>
            </div>
        </div>
        
        <div class="content">
"""
        
        # カテゴリ別の内容を生成
        for category, category_data in analysis_result['categories'].items():
            html_content += f"""
            <div class="category">
                <div class="category-header">
                    <div class="category-title">{category}</div>
                    <div class="category-summary">
                        {category_data['record_count']}件 / ¥{category_data['total_amount']:,.0f}
                    </div>
                </div>
"""
            
            # 仕入先別の内容を生成
            for supplier, supplier_data in category_data['suppliers'].items():
                html_content += f"""
                <div class="supplier">
                    <div class="supplier-header">
                        <div class="supplier-title">{supplier}</div>
                        <div class="supplier-summary">
                            {supplier_data['record_count']}件 / ¥{supplier_data['total_amount']:,.0f}
                        </div>
                    </div>
                    <table class="products-table">
                        <thead>
                            <tr>
                                <th>ユニットNO.</th>
                                <th>部品番号</th>
                                <th>品名</th>
                                <th>メーカー</th>
                                <th>材質・型式</th>
                                <th>数量</th>
                                <th>単価</th>
                                <th>仕入金額</th>
                                <th>受入日</th>
                            </tr>
                        </thead>
                        <tbody>
"""
                
                # 商品別の内容を生成
                for product in supplier_data['products']:
                    html_content += f"""
                        <tr>
                            <td class="unit-info">{product['unit_no']}</td>
                            <td class="unit-info">{product['part_no']}</td>
                            <td>{product['product_name']}</td>
                            <td>{product['manufacturer']}</td>
                            <td>{product['material_type']}</td>
                            <td class="quantity">{product['quantity']:,.0f}</td>
                            <td class="price">¥{product['unit_price']:,.0f}</td>
                            <td class="price">¥{product['total_price']:,.0f}</td>
                            <td>{product['receive_date']}</td>
                        </tr>
"""
                
                html_content += """
                        </tbody>
                    </table>
                </div>
"""
            
            html_content += """
            </div>
"""
        
        # フッター
        html_content += f"""
        </div>
        
        <div class="footer">
            <p>生成日時: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}</p>
        </div>
    </div>
</body>
</html>
"""
        
        # HTMLファイルに出力
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"HTMLレポートを出力しました: {file_path}")
        return str(file_path)

def main():
    """メイン関数"""
    print("仕入レポート分析プログラムを開始します")
    print("ファイル選択ダイアログが表示されます。")
    
    # 分析生成器のインスタンスを作成
    analyzer = PurchaseAnalysisGenerator()
    
    try:
        # オリジナルデータを読み込み
        print("\n=== ステップ1: オリジナルデータの読み込み ===")
        original_data = analyzer.load_original_data()
        
        # 分類置換テーブルを適用
        print("\n=== ステップ2: 分類置換テーブルの適用 ===")
        processed_data = analyzer.apply_category_mapping(original_data)
        
        # ファイルNo.の一覧を取得
        print("\n=== ステップ3: ファイルNo.一覧の取得 ===")
        file_numbers = analyzer.get_unique_file_numbers(processed_data)
        
        if not file_numbers:
            raise ValueError("データにファイルNo.が見つかりません")
        
        print(f"検出されたファイルNo.数: {len(file_numbers)}")
        print(f"ファイルNo.一覧: {file_numbers[:10]}{'...' if len(file_numbers) > 10 else ''}")
        
        # ファイルNo.選択ダイアログを表示
        print("\n=== ステップ4: ファイルNo.の選択 ===")
        selected_file_no = analyzer.select_file_no_dialog(file_numbers)
        
        if selected_file_no is None:
            print("ファイルNo.が選択されませんでした")
            return
        
        print(f"選択されたファイルNo.: {selected_file_no}")
        
        # 選択されたファイルNo.の分析を実行
        print("\n=== ステップ5: 仕入データの分析 ===")
        analysis_result = analyzer.analyze_file_purchases(processed_data, selected_file_no)
        
        if analysis_result is None:
            raise ValueError(f"ファイルNo. {selected_file_no} の分析に失敗しました")
        
        # HTMLレポートを生成
        print("\n=== ステップ6: HTMLレポートの生成 ===")
        html_file = analyzer.generate_html_report(analysis_result)
        
        # ブラウザで表示
        print("\n=== ステップ7: ブラウザでの表示 ===")
        print(f"ブラウザでレポートを開いています: {html_file}")
        webbrowser.open(f'file://{Path(html_file).absolute()}')
        
        print("\n=== 処理完了 ===")
        print(f"ファイルNo. {selected_file_no} の仕入レポート分析が完了しました。")
        print(f"HTMLレポート: {html_file}")
        
        # 完了メッセージを表示
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("処理完了", f"仕入レポート分析が完了しました。\n\nファイルNo.: {selected_file_no}\nHTMLレポート: {html_file}")
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

# 2025-08-29 17:30:00
