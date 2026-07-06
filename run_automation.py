# -*- coding: utf-8 -*-
"""
出荷明細作成自動化スクリプト (Python直接集計・ハイブリッド版)
===============================
実行方法: 実行.bat をダブルクリック、または python run_automation.py

input/ フォルダに以下のCSVを置いてから実行してください:
  - 出荷予定振分*.csv       → RETAIL.csv にリネームして処理
  - 営業日付別売上分析*.csv  → 店舗売上
  - 卸売上明細.csv           → ZOZO売上
  - 卸売上明細 (1).csv       → OIOI売上
  - order_*.csv             → EC受注
"""

import os
import sys
import glob
import traceback
import tkinter as tk
from tkinter import messagebox

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

# ============================================================
# 独自例外の定義
# ============================================================
class AutomationError(Exception):
    """業務自動化処理中に発生する想定内のエラー"""
    def __init__(self, title, message):
        self.title = title
        self.message = message
        super().__init__(message)

# ============================================================
# 設定
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
TEMPLATE_PATH = os.path.join(BASE_DIR, "★RETAIL_テンプレート.xlsx")
ERROR_LOG_PATH = os.path.join(BASE_DIR, "error.log")
ENCODING = "cp932"

# 集計店舗リスト（店舗名と対応するカラム順序。マスタと1対1で連動）
STORE_COLUMNS = ['名古屋', 'TOKYO', 'ルクア大阪', 'ヒュッテ', '京王新宿', '大丸心斎橋', '6142', '玉川高島屋', 'NODE', 'NARITA']

# ============================================================
# ユーティリティ
# ============================================================
def show_error(title, msg):
    print(f"エラー: {title}\n{msg}")
    if os.environ.get("NO_GUI"):
        sys.exit(1)
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(title, msg)
        root.destroy()
    except Exception:
        pass
    sys.exit(1)

def show_info(msg):
    print(f"完了: {msg}")
    if os.environ.get("NO_GUI"):
        return
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("完了", msg)
        root.destroy()
    except Exception:
        pass

def find_single_csv(pattern_keywords, exclude_keywords=None):
    all_csvs = glob.glob(os.path.join(INPUT_DIR, "*.csv"))
    matched = []
    clean_patterns = [kw.replace(" ", "").replace("　", "") for kw in pattern_keywords]
    clean_excludes = [ek.replace(" ", "").replace("　", "") for ek in exclude_keywords] if exclude_keywords else []
    for path in all_csvs:
        name = os.path.basename(path)
        clean_name = name.replace(" ", "").replace("　", "")
        if all(kw in clean_name for kw in clean_patterns):
            if clean_excludes and any(ek in clean_name for ek in clean_excludes):
                continue
            matched.append(path)

    if len(matched) == 0:
        keywords_str = "、".join(pattern_keywords)
        raise AutomationError(
            "必要なCSVファイルが見つかりません",
            f"【発生した状況】\n「input」フォルダの中に「{keywords_str}」を含むCSVファイルが見つかりません。\n\n"
            "【対処方法】\n指定されたCSVファイルを「input」フォルダ内に保存してから、もう一度実行してください。"
        )
    if len(matched) > 1:
        names = "\n".join(f"・{os.path.basename(p)}" for p in matched)
        keywords_str = "、".join(pattern_keywords)
        raise AutomationError(
            "同じ種類のCSVファイルが複数あります",
            f"【発生した状況】\n「input」フォルダの中に「{keywords_str}」に該当するCSVファイルが複数見つかりました。\n\n"
            f"見つかったファイル:\n{names}\n\n"
            "【対処方法】\n不要な古いCSVファイルを「input」フォルダから移動または削除して、1つだけにしてからもう一度実行してください。"
        )
    return matched[0]

# ============================================================
# tennenブランドの除外判定
# ============================================================
def is_tennen(code, brand_name=None):
    if brand_name and str(brand_name).upper().strip() == "TENNEN":
        return True
    if not code:
        return False
    c = str(code).upper().strip()
    return c.startswith("TNT") or c.startswith("TNP") or c.startswith("TNF")

# ============================================================
# データ集計処理 (Pandasによる直接集計)
# ============================================================
def load_store_sales():
    print("  店舗売上 (営業日付別売上分析) を集計中...")
    src = find_single_csv(["営業日付別売上分析"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    brand_mask = df["表記部門名1"].isin(["FRV", "FJALLRAVEN"])
    df_filtered = df[brand_mask].copy()
    # tennen除外
    df_filtered = df_filtered[~df_filtered["3rd Item No."].apply(is_tennen)]
    
    df_filtered["数量"] = pd.to_numeric(df_filtered["数量"], errors="coerce").fillna(0)
    pivot_df = df_filtered.pivot_table(index="3rd Item No.", columns="店舗名称", values="数量", aggfunc="sum").reset_index()
    return pivot_df

def get_store_sales_value(pivot_df, code_short, store_name):
    if pivot_df is None or code_short not in pivot_df["3rd Item No."].values:
        return 0
    row_df = pivot_df[pivot_df["3rd Item No."] == code_short]
    
    # 部分一致列の検索
    if store_name == 'ヒュッテ':
        matched_cols = [c for c in pivot_df.columns if 'ヒュッテ' in c or 'HUTTE' in c.upper()]
    elif store_name == 'TOKYO':
        # TOKYO NODEとの誤混同を防ぐ
        matched_cols = [c for c in pivot_df.columns if 'TOKYO' in c and 'NODE' not in c]
    else:
        matched_cols = [c for c in pivot_df.columns if store_name in c]
        
    if not matched_cols:
        return 0
        
    val = pd.to_numeric(row_df[matched_cols].iloc[0], errors='coerce').sum()
    return int(val) if not pd.isna(val) else 0

def load_zozo_sales():
    print("  ZOZO売上 (卸売上明細) を集計中...")
    src = find_single_csv(["卸売上明細"], exclude_keywords=["卸売上明細 (1)", "卸売上明細(1)"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    brand_mask = df["ブランド名"].isin(["FRV", "FJALLRAVEN"])
    df_filtered = df[brand_mask].copy()
    # tennen除外
    df_filtered = df_filtered[~df_filtered["商品コード"].apply(is_tennen)]

    df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
    agg_df = df_filtered.groupby("商品コード")["販売数量"].sum().reset_index()
    
    return dict(zip(agg_df["商品コード"], agg_df["販売数量"]))

def load_oioi_sales():
    print("  OIOI売上 (卸売上明細 (1)) を集計中...")
    src = find_single_csv(["卸売上明細 (1)"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    brand_mask = df.get("ブランド名", pd.Series(dtype=str)).isin(["FRV", "FJALLRAVEN"])
    df_filtered = df[brand_mask].copy()
    # tennen除外
    if "商品コード" in df_filtered.columns:
        df_filtered = df_filtered[~df_filtered["商品コード"].apply(is_tennen)]

    if "販売数量" in df_filtered.columns and "商品コード" in df_filtered.columns:
        df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
        agg_df = df_filtered.groupby("商品コード")["販売数量"].sum().reset_index()
        return dict(zip(agg_df["商品コード"], agg_df["販売数量"]))
    return {}

def load_ec_sales():
    print("  EC売上 (order_*) を集計中...")
    src = find_single_csv(["order_"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    df = df[~df["決済方法(ステータス)"].str.contains("キャンセル", na=False)]
    df = df[~df["注文者"].str.contains("店舗客注", na=False)]
    df = df.reset_index(drop=True)

    df["個数"] = pd.to_numeric(df["個数"], errors="coerce").fillna(0)
    
    key_col = "オプション独自コード" if "オプション独自コード" in df.columns else "商品コード"
    # tennen除外
    brand_col = "ブランド名" if "ブランド名" in df.columns else None
    if brand_col:
        df = df[~df.apply(lambda row: is_tennen(row[key_col], row[brand_col]), axis=1)]
    else:
        df = df[~df[key_col].apply(is_tennen)]

    agg_df = df.groupby(key_col)["個数"].sum().reset_index()
    return dict(zip(agg_df[key_col], agg_df["個数"]))

# ============================================================
# 新規データ抽出用ロード関数
# ============================================================
def load_hutte():
    print("  HUTTE出荷指示 (★出荷指示フォーマット【*】hutte.xlsx) を読み込み中...")
    pattern = os.path.join(INPUT_DIR, "★出荷指示フォーマット【*】hutte.xlsx")
    files = glob.glob(pattern)
    if not files:
        print("  [情報] HUTTE出荷指示ファイルが見つかりません。")
        return {}
    latest_file = max(files, key=os.path.getmtime)
    try:
        df = pd.read_excel(latest_file, sheet_name=0, dtype=str)
        if len(df.columns) > 12:
            key_col = df.columns[3]
            val_col = df.columns[12]
            df = df.dropna(subset=[key_col])
            df = df[~df[key_col].apply(is_tennen)]
            return dict(zip(df[key_col].str.strip(), df[val_col]))
    except Exception as e:
        print(f"  [警告] HUTTE出荷指示の読み込みに失敗しました: {e}")
    return {}

def load_zozo_master():
    print("  ZOZO/OIOIマスタ (商品マスタ-ZOZO-OlOl.xlsx) を読み込み中...")
    path = os.path.join(INPUT_DIR, "商品マスタ-ZOZO-OlOl.xlsx")
    if not os.path.exists(path):
        print("  [情報] ZOZO/OIOIマスタファイルが見つかりません。")
        return {}, {}
    try:
        df = pd.read_excel(path, sheet_name=0, header=1, dtype=str)
        if len(df.columns) > 27:
            key_col = df.columns[0]
            zozo_col = df.columns[26]
            oioi_col = df.columns[27]
            df = df.dropna(subset=[key_col])
            df = df[~df[key_col].apply(is_tennen)]
            z_dict = dict(zip(df[key_col].str.strip(), df[zozo_col].fillna("")))
            o_dict = dict(zip(df[key_col].str.strip(), df[oioi_col].fillna("")))
            return z_dict, o_dict
    except Exception as e:
        print(f"  [警告] ZOZO/OIOIマスタの読み込みに失敗しました: {e}")
    return {}, {}

def load_tokka():
    print("  特価在庫 (2026特価在庫.xlsx) を読み込み中...")
    path = os.path.join(INPUT_DIR, "2026特価在庫.xlsx")
    if not os.path.exists(path):
        print("  [情報] 特価在庫ファイルが見つかりません。")
        return {}
    try:
        df = pd.read_excel(path, sheet_name=0, dtype=str)
        if len(df.columns) > 22:
            key_col = df.columns[2]
            val_col = df.columns[22]
            df = df.dropna(subset=[key_col])
            df = df[~df[key_col].apply(is_tennen)]
            return dict(zip(df[key_col].str.strip(), df[val_col].fillna("")))
    except Exception as e:
        print(f"  [警告] 特価在庫の読み込みに失敗しました: {e}")
    return {}

# ============================================================
# 手順1: 出荷予定振分.csv から RETAIL ワークブックの作成
# ============================================================
def step1_create_retail_wb():
    print("[手順1] 出荷予定振分.csv から RETAILシートの構築")
    src = find_single_csv(["出荷予定振分"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)
    
    wb_retail = openpyxl.Workbook()
    ws = wb_retail.active
    ws.title = "RETAIL"
    
    # ヘッダー書き込み
    for c_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=c_idx).value = col_name
        
    # データ書き込み
    for r_idx, row in enumerate(df.values.tolist(), start=2):
        for c_idx, val in enumerate(row, start=1):
            if pd.isna(val) if isinstance(val, float) else val is None:
                ws.cell(row=r_idx, column=c_idx).value = ""
            else:
                if isinstance(val, str):
                    try:
                        if val.isdigit():
                            ws.cell(row=r_idx, column=c_idx).value = int(val)
                        else:
                            ws.cell(row=r_idx, column=c_idx).value = float(val)
                    except ValueError:
                        ws.cell(row=r_idx, column=c_idx).value = val
                else:
                    ws.cell(row=r_idx, column=c_idx).value = val
                    
    return wb_retail

# ============================================================
# 手順2: Python直接集計・マージおよび書式適用 (VBA完全代替)
# ============================================================
def step2_python_direct_merge(wb_retail, pivot_df, zozo_dict, oioi_dict, ec_dict, hutte_dict, zozo_m_dict, oioi_m_dict, tokka_dict):
    print("[手順2] Pythonによる直接マージ＆書式適用 (VBA代替処理)")
    
    if "order" in wb_retail.sheetnames:
        del wb_retail["order"]

    ws_retail = wb_retail["RETAIL"]
    
    # 【ボトルネック解消】AW1:BY[max_row] をクリア (10000行ループを実際の最終行に削減)
    max_r_retail = max(ws_retail.max_row, 1)
    for r in range(1, max_r_retail + 1):
        for c in range(49, 78):
            ws_retail.cell(row=r, column=c).value = None
            
    # シートの複製
    ws_order = wb_retail.copy_worksheet(ws_retail)
    ws_order.title = "order"

    # 【列インデックス依存の排除】"商品コード" "商品名" "ブランド名" の列を動的に特定
    idx_code = None
    idx_name = None
    idx_brand = None
    
    for c in range(1, ws_order.max_column + 1):
        val = ws_order.cell(row=1, column=c).value
        val_str = str(val).strip() if val is not None else ""
        if val_str in ("商品コード", "4"):
            idx_code = c
        elif val_str in ("商品名", "5"):
            idx_name = c
        elif val_str in ("6", "ブランド名", "表記部門名1"):
            idx_brand = c
            
    if not idx_code: idx_code = 4
    if not idx_name: idx_name = 5
    if not idx_brand: idx_brand = 6

    # 動的な列削除
    cols_to_delete_before = idx_code - 1
    if cols_to_delete_before > 0:
        ws_order.delete_cols(1, cols_to_delete_before)
        idx_name -= cols_to_delete_before
        idx_brand -= cols_to_delete_before
        idx_code = 1
        
    # A列の書式適用
    for row in range(1, ws_order.max_row + 1):
        cell = ws_order.cell(row=row, column=1)
        cell.number_format = '#,##0;[Red](#,##0)'
        cell.alignment = Alignment(horizontal='left', vertical='center')

    # B列幅設定
    ws_order.column_dimensions['B'].width = 15
    
    # ブランド名/部門名が入っているC列を削除
    if idx_brand > 0:
        ws_order.delete_cols(idx_brand, 1)

    ws_order["D4"].value = "RETAIL確保"
    ws_order.delete_rows(3, 1) # 3行目を削除 (これで店舗コード行が消え、日本語店舗名がRow3に繰り上がります)
    
    # 1. BULK不要列の削除 (R:X 削除。元の18列目から7列分)
    # 動的に残る列：C列(3)=RETAIL確保, D列(4)=RETAIL確保(D4値), E〜Q列(5〜17)=確保(13列分)
    # 元のR〜X列は、不要なBULK列。
    # 削除後のC列(3), D列(4), E〜Q列(5〜17)は残すため、18列目から7列分削除する。
    ws_order.delete_cols(18, 7)
    
    # 2. 在庫不要列の後半の削除 (AE:AK 削除。7列分)
    ws_order.delete_cols(31, 7)
    
    # "確保" ヘッダー (C2:Q2)
    ws_order.merge_cells("C2:Q2")
    ws_order["C2"].value = "確保"
    ws_order["C2"].alignment = Alignment(horizontal='center', vertical='center')
    fill_kakubo = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_blue_c2_d3 = PatternFill(start_color="00CCFFFF", end_color="00CCFFFF", fill_type="solid") # 水色
    for r in range(2, 4):
        for c in range(3, 18):
            ws_order.cell(row=r, column=c).fill = fill_kakubo
    # C2:D3 を水色に上書き
    for r in range(2, 4):
        for c in range(3, 5):
            ws_order.cell(row=r, column=c).fill = fill_blue_c2_d3

    # "在庫" ヘッダー (R2:AD2)
    ws_order.merge_cells("R2:AD2")
    ws_order["R2"].value = "在庫"
    ws_order["R2"].alignment = Alignment(horizontal='center', vertical='center')
    fill_zaiko = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    for r in range(2, 4):
        for c in range(18, 31):
            ws_order.cell(row=r, column=c).fill = fill_zaiko

    # "ORDER" ヘッダー (AE2:AQ2)
    ws_order.merge_cells("AE2:AQ2")
    ws_order["AE2"].value = "ORDER"
    ws_order["AE2"].alignment = Alignment(horizontal='center', vertical='center')
    fill_order = PatternFill(start_color="FFFF00", fill_type="solid")
    for r in range(2, 4):
        for c in range(31, 44):
            ws_order.cell(row=r, column=c).fill = fill_order

    # 【ボトルネック解消】AR:BD列(44〜56列目)をクリア
    max_r_order = max(ws_order.max_row, 1)
    for r in range(1, max_r_order + 1):
        for c in range(44, 57):
            ws_order.cell(row=r, column=c).value = None

    # "過不足" ヘッダー (AR2:BD2)
    ws_order.merge_cells("AR2:BD2")
    ws_order["AR2"].value = "過不足"
    ws_order["AR2"].alignment = Alignment(horizontal='center', vertical='center')

    # 3行目の見出しを縦書きに
    align_vertical = Alignment(textRotation=255, horizontal='center', vertical='center')
    for c in range(3, 44):
        ws_order.cell(row=3, column=c).alignment = align_vertical
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.column_dimensions[col_letter].width = 5

    # A2:B2 切り貼り
    ws_order["A3"].value = ws_order["A2"].value
    ws_order["B3"].value = ws_order["B2"].value
    ws_order["A2"].value = None
    ws_order["B2"].value = None
    ws_order["A3"].alignment = Alignment(horizontal='center', vertical='center')
    ws_order["B3"].alignment = Alignment(horizontal='center', vertical='center')

    # 最終行検出
    lastRow = 4
    for r in range(ws_order.max_row, 3, -1):
        if ws_order.cell(row=r, column=2).value is not None:
            lastRow = r
            break

    # E:Q列を非表示
    for c in range(5, 18):
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.column_dimensions[col_letter].hidden = True

    # --------------------------------------------------------
    # 静的デザインテンプレートから店舗ヘッダーをロード
    # --------------------------------------------------------
    print("  マスタテンプレートから店舗ヘッダーを読み込み中...")
    wb_temp = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)
    ws_temp = wb_temp["計算式2026"]
    
    # 3行目の店舗名をテンプレートからコピーして order シートの R3:BG3 に適用
    # (R:18 から BG:59 まで)
    for c in range(18, 60):
        ws_order.cell(row=3, column=c).value = ws_temp.cell(row=3, column=c).value

    wb_temp.close()

    # --------------------------------------------------------
    # 列インデックスを動的に検索・解決する仕組みの構築
    # --------------------------------------------------------
    col_map = {}
    current_area = None
    # 2行目の結合セルの値を取得するために、左から右へ走査しながら結合エリアの値を取得
    for col in range(1, ws_order.max_column + 1):
        area_val = ws_order.cell(row=2, column=col).value
        if area_val is not None:
            area_val_str = str(area_val).strip()
            if area_val_str in ("確保", "在庫", "ORDER", "過不足"):
                current_area = area_val_str
        
        store_val = ws_order.cell(row=3, column=col).value
        if store_val is not None and current_area is not None:
            store_val_str = str(store_val).strip()
            col_map[(current_area, store_val_str)] = col

    stores_list = ["ZOZO", "名古屋", "OIOI", "YSEC", "TOKYO", "ルクア大阪", "ヒュッテ", "京王新宿", "大丸心斎橋", "6142", "玉川高島屋", "NODE", "NARITA"]
    
    bulk_store_map = {}
    order_store_map = {}
    kabusoku_store_map = {}
    
    for (area, store), col_idx in col_map.items():
        matched_store = None
        for s in stores_list:
            if s == store or (s == "YSEC" and store == "EC") or (s == "EC" and store == "YSEC") or (s == "OIOI" and store == "丸井web"):
                matched_store = s
                break
        
        if matched_store:
            if area == "確保":
                bulk_store_map[matched_store] = col_idx
            elif area == "ORDER":
                order_store_map[matched_store] = col_idx
            elif area == "過不足":
                kabusoku_store_map[matched_store] = col_idx

    # その他の列番号の動的解決 (C列, D列, 57〜59列目)
    col_c = 3
    col_d = 4
    col_kabusoku_total = None
    col_retail_kakubo = None
    col_kabusoku_zan = None
    
    for c in range(1, ws_order.max_column + 1):
        val = ws_order.cell(row=3, column=c).value
        val_str = str(val).strip() if val is not None else ""
        if c > 43:
            if "合計" in val_str or val_str == "過不足合計":
                col_kabusoku_total = c
            elif val_str == "RETAIL確保" and c != col_d:
                col_retail_kakubo = c
            elif (val_str == "過不足" or val_str == "過不足残") and c not in kabusoku_store_map.values():
                col_kabusoku_zan = c

    if not col_kabusoku_total: col_kabusoku_total = 57
    if not col_retail_kakubo: col_retail_kakubo = 58
    if not col_kabusoku_zan: col_kabusoku_zan = 59

    # --------------------------------------------------------
    # 商品マスタ (商品マスタ全項目.csv) のロード
    # --------------------------------------------------------
    print("  商品マスタ全項目.csv からコードと商品名を取得中...")
    master_dict = {} # code_short -> {'name': ..., 'code': ...}
    master_path = os.path.join(BASE_DIR, "商品マスタ全項目.csv")
    if os.path.exists(master_path):
        try:
            df_m = pd.read_csv(master_path, encoding="cp932", dtype=str)
            for _, row in df_m.iterrows():
                m_code = str(row.iloc[0]).strip() if not pd.isna(row.iloc[0]) else ""
                m_name = str(row.iloc[2]).strip() if not pd.isna(row.iloc[2]) else ""
                if m_code:
                    master_dict[m_code] = {
                        'name': m_name,
                        'code': m_code
                    }
        except Exception as e:
            print(f"  警告: 商品マスタのロードに失敗しました: {e}")

    # --------------------------------------------------------
    # 売上CSVからバックアップ用の「コード ↔ 商品名」を取得
    # (マスタにない北欧展グッズなどの商品名救済用)
    # --------------------------------------------------------
    backup_names = {}
    try:
        src_store = find_single_csv(["営業日付別売上分析"])
        df_store_raw = pd.read_csv(src_store, encoding=ENCODING, dtype=str)
        for _, row in df_store_raw.iterrows():
            code = str(row["3rd Item No."]).strip() if not pd.isna(row["3rd Item No."]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    try:
        src_zozo = find_single_csv(["卸売上明細"], exclude_keywords=["卸売上明細 (1)", "卸売上明細(1)"])
        df_zozo_raw = pd.read_csv(src_zozo, encoding=ENCODING, dtype=str)
        for _, row in df_zozo_raw.iterrows():
            code = str(row["商品コード"]).strip() if not pd.isna(row["商品コード"]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    try:
        src_oioi = find_single_csv(["卸売上明細 (1)"])
        df_oioi_raw = pd.read_csv(src_oioi, encoding=ENCODING, dtype=str)
        for _, row in df_oioi_raw.iterrows():
            code = str(row["商品コード"]).strip() if not pd.isna(row["商品コード"]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    try:
        src_ec = find_single_csv(["order_"])
        df_ec_raw = pd.read_csv(src_ec, encoding=ENCODING, dtype=str)
        key_col = "オプション独自コード" if "オプション独自コード" in df_ec_raw.columns else "商品コード"
        for _, row in df_ec_raw.iterrows():
            code = str(row[key_col]).strip() if not pd.isna(row[key_col]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    # --------------------------------------------------------
    # 売上があるすべての商品コードの抽出
    # --------------------------------------------------------
    all_sales_codes = set()
    for code, qty in zozo_dict.items():
        if qty != 0: all_sales_codes.add(code)
    for code, qty in oioi_dict.items():
        if qty != 0: all_sales_codes.add(code)
    for code, qty in ec_dict.items():
        if qty != 0 and "GIFT" not in code:  # GIFT500などのラッピング代は除外
            all_sales_codes.add(code)
    if pivot_df is not None and not pivot_df.empty:
        num_cols = [c for c in pivot_df.columns if c != "3rd Item No."]
        for _, row in pivot_df.iterrows():
            code = str(row["3rd Item No."]).strip() if not pd.isna(row["3rd Item No."]) else ""
            if code:
                total_qty = pd.to_numeric(row[num_cols], errors="coerce").fillna(0).sum()
                if total_qty != 0:
                    all_sales_codes.add(code)

    # 出荷予定振分.csv(orderシート)の現在のA列コード
    order_codes = set()
    for r in range(4, lastRow + 1):
        val = ws_order.cell(row=r, column=1).value
        if val:
            order_codes.add(str(val).strip())

    # 漏れている（出荷予定にないが売上があった）商品コード
    # 新規追加商品は在庫(C, D, E〜Q)がいずれも0であり、売上が発生すると過不足残が不足赤になります。
    # 誤発注防止仕様に基づき、これらの商品は売上0（追加対象外）として扱われます。
    leaked_codes = []
    print(f"  出荷予定にない売上発生商品を追加中... (追加数: {len(leaked_codes)})")

    # --------------------------------------------------------
    # orderシートおよびRETAILシートの末尾への行追加
    # --------------------------------------------------------
    # RETAILシートの最終行検出 (Col4=商品コード が空でない行)
    lastRowRetail = 4
    for r in range(ws_retail.max_row, 3, -1):
        if ws_retail.cell(row=r, column=4).value is not None:
            lastRowRetail = r
            break

    current_order_row = lastRow + 1
    current_retail_row = lastRowRetail + 1

    # RETAILシートの開始列と終了列の動的設定
    start_order_col = min(order_store_map.values()) if order_store_map else 31
    end_order_col = col_kabusoku_zan
    num_formula_cols = end_order_col - start_order_col + 1
    start_retail_col = 49

    for code in leaked_codes:
        # 商品名とコードの解決
        p_name = "不明な商品"
        if code in master_dict and master_dict[code]['name']:
            p_name = master_dict[code]['name']
        elif code in backup_names:
            p_name = backup_names[code]
            
        p_code = code
        if code in master_dict and master_dict[code]['code']:
            p_code = master_dict[code]['code']

        # orderシートへ行追加
        ws_order.cell(row=current_order_row, column=1).value = p_code
        ws_order.cell(row=current_order_row, column=2).value = p_name
        
        # C〜Q列 (BULK): すべて 0
        for col in range(3, 18):
            ws_order.cell(row=current_order_row, column=col).value = 0
        # R〜AD列 (在庫): すべて 0
        for col in range(18, 31):
            ws_order.cell(row=current_order_row, column=col).value = 0
            
        # AE〜AQ列 (ORDER) への売上書き込み (店舗名とインデックスを動的に解決)
        for store_name, col_idx in order_store_map.items():
            if store_name == "ZOZO":
                ws_order.cell(row=current_order_row, column=col_idx).value = int(zozo_dict.get(code, 0))
            elif store_name == "OIOI":
                ws_order.cell(row=current_order_row, column=col_idx).value = int(oioi_dict.get(code, 0))
            elif store_name in ("YSEC", "EC"):
                ws_order.cell(row=current_order_row, column=col_idx).value = int(ec_dict.get(code, 0))
            else:
                ws_order.cell(row=current_order_row, column=col_idx).value = get_store_sales_value(pivot_df, code, store_name)

        # RETAILシートへ行追加
        ws_retail.cell(row=current_retail_row, column=4).value = p_code
        ws_retail.cell(row=current_retail_row, column=5).value = p_name
        ws_retail.cell(row=current_retail_row, column=6).value = "FRV"
        ws_retail.cell(row=current_retail_row, column=7).value = 0
        for col in range(8, 29):
            ws_retail.cell(row=current_retail_row, column=col).value = 0
            
        # RETAILシートからorderシートへの参照式を設定 (ORDERエリアの13列分)
        for c_idx in range(13):
            target_col_letter = openpyxl.utils.get_column_letter(49 + c_idx) # AWから
            source_col_letter = openpyxl.utils.get_column_letter(31 + c_idx) # AEから
            ws_retail.cell(row=current_retail_row, column=49 + c_idx).value = f"=order!{source_col_letter}{current_order_row}"

        current_order_row += 1
        current_retail_row += 1

    # 最終行変数を更新
    lastRow = current_order_row - 1
    lastRowRetail = current_retail_row - 1

    # R2:AD[lastRow] に対応する在庫エリアの左端と右端の動的な罫線適用
    # 在庫エリアの開始列と終了列を動的に解決
    start_zaiko_col = min(col_idx for (area, store), col_idx in col_map.items() if area == "在庫")
    end_zaiko_col = max(col_idx for (area, store), col_idx in col_map.items() if area == "在庫")
    
    thin_side = Side(border_style="thin", color="000000")
    border_left_only = Border(left=thin_side)
    border_right_only = Border(right=thin_side)
    for r in range(2, lastRow + 1):
        ws_order.cell(row=r, column=start_zaiko_col).border = border_left_only
        ws_order.cell(row=r, column=end_zaiko_col).border = border_right_only

    # --------------------------------------------------------
    # 既存商品の売上データ書き込み
    # --------------------------------------------------------
    print("  既存商品の売上データを書き込み中...")
    for r in range(4, lastRow - len(leaked_codes) + 1):
        code_short = ws_order.cell(row=r, column=1).value
        if not code_short:
            for store_name, col_idx in order_store_map.items():
                ws_order.cell(row=r, column=col_idx).value = 0
            continue
            
        code_short = str(code_short).strip()

        # 売上値の事前計算
        raw_sales = {}
        total_sales = 0
        for store_name, col_idx in order_store_map.items():
            qty = 0
            if store_name == "ZOZO":
                qty = int(zozo_dict.get(code_short, 0))
            elif store_name == "OIOI":
                qty = int(oioi_dict.get(code_short, 0))
            elif store_name in ("YSEC", "EC"):
                qty = int(ec_dict.get(code_short, 0))
            else:
                qty = get_store_sales_value(pivot_df, code_short, store_name)
            raw_sales[store_name] = qty
            total_sales += qty

        # 在庫不足（C=0 かつ D=0、かつ 確保枠合計 < 売上合計）の判定
        c_val = ws_order.cell(row=r, column=3).value or 0
        d_val = ws_order.cell(row=r, column=4).value or 0
        kabu_sum = sum(ws_order.cell(row=r, column=col).value or 0 for col in range(5, 18))

        if c_val == 0 and d_val == 0 and kabu_sum < total_sales:
            # 誤発注防止のため、売上をすべて0にする
            for store_name, col_idx in order_store_map.items():
                ws_order.cell(row=r, column=col_idx).value = 0
        else:
            for store_name, col_idx in order_store_map.items():
                ws_order.cell(row=r, column=col_idx).value = raw_sales[store_name]

    # --------------------------------------------------------
    # 計算式（過不足、過不足残、SUBTOTAL合計）の書き込み
    # --------------------------------------------------------
    print("  過不足などの計算式を設定中...")
    for r in range(4, lastRow + 1):
        # 各店舗ごとの過不足 (BULK量 - ORDER量) を動的に設定
        for s in stores_list:
            b_col = bulk_store_map.get(s)
            o_col = order_store_map.get(s)
            k_col = kabusoku_store_map.get(s)
            if b_col and o_col and k_col:
                b_let = openpyxl.utils.get_column_letter(b_col)
                o_let = openpyxl.utils.get_column_letter(o_col)
                ws_order.cell(row=r, column=k_col).value = f"={b_let}{r}-{o_let}{r}"
            
        # 各列のアルファベット文字を取得
        kabusoku_start_let = openpyxl.utils.get_column_letter(min(kabusoku_store_map.values()))
        kabusoku_end_let = openpyxl.utils.get_column_letter(max(kabusoku_store_map.values()))
        
        let_c = openpyxl.utils.get_column_letter(col_c)
        let_d = openpyxl.utils.get_column_letter(col_d)
        let_kabusoku_total = openpyxl.utils.get_column_letter(col_kabusoku_total)
        let_retail_kakubo = openpyxl.utils.get_column_letter(col_retail_kakubo)
        let_kabusoku_zan = openpyxl.utils.get_column_letter(col_kabusoku_zan)
        
        # 過不足合計
        ws_order.cell(row=r, column=col_kabusoku_total).value = f"=SUMIF({kabusoku_start_let}{r}:{kabusoku_end_let}{r},\"<0\",{kabusoku_start_let}{r}:{kabusoku_end_let}{r})"
        # RETAIL確保
        ws_order.cell(row=r, column=col_retail_kakubo).value = f"=IF({let_kabusoku_total}{r}<0,{let_d}{r}+{let_kabusoku_total}{r},{let_d}{r})"
        # 過不足残 (NWA残)
        ws_order.cell(row=r, column=col_kabusoku_zan).value = f"=IF({let_retail_kakubo}{r}<0,{let_c}{r}+{let_retail_kakubo}{r},{let_c}{r})"

    # 1行目の SUBTOTAL 計算式の書き込み
    for c in range(1, ws_order.max_column + 1):
        ws_order.cell(row=1, column=c).value = None
        
    max_order_col = max(order_store_map.values()) if order_store_map else 43
    for c in range(3, max_order_col + 1):
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.cell(row=1, column=c).value = f"=SUBTOTAL(9,{col_letter}4:{col_letter}3500)"

    # 幅と非表示設定 (過不足エリアの列インデックスに対して動的に適用)
    kabusoku_cols_indices = list(kabusoku_store_map.values()) + [col_kabusoku_total, col_retail_kakubo, col_kabusoku_zan]
    kabusoku_min_col = min(kabusoku_cols_indices) if kabusoku_cols_indices else 44
    kabusoku_max_col = max(kabusoku_cols_indices) if kabusoku_cols_indices else 59
    
    for c in range(kabusoku_min_col, kabusoku_max_col + 1):
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.column_dimensions[col_letter].width = 4.43
        if c != col_kabusoku_zan:
            ws_order.column_dimensions[col_letter].hidden = True
            
    zan_letter = openpyxl.utils.get_column_letter(col_kabusoku_zan)
    ws_order[f"{zan_letter}3"].value = "過不足"
    ws_order[f"{zan_letter}3"].alignment = Alignment(textRotation=255, horizontal='center', vertical='center')
    ws_order[f"{zan_letter}3"].fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid") # BG3を赤色に

    # 条件付き書式 (過不足マイナスでA列を赤く)
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    rule = FormulaRule(formula=[f"${zan_letter}4<0"], stopIfTrue=False, fill=red_fill)
    ws_order.conditional_formatting.add(f"A4:A{lastRow}", rule)

    # RETAIL シートへの数式展開 (AW列から ORDERエリアの13列分を展開)
    lastRowRetail = 5
    for r in range(ws_retail.max_row, 4, -1):
        if ws_retail.cell(row=r, column=4).value is not None:
            lastRowRetail = r
            break
            
    for r in range(5, lastRowRetail + 1):
        # ORDERエリアは AE(31) から AQ(43) までの 13列
        for c_idx in range(13):
            target_col_letter = openpyxl.utils.get_column_letter(49 + c_idx) # AWから
            source_col_letter = openpyxl.utils.get_column_letter(31 + c_idx) # AEから
            source_row = r - 1
            ws_retail[f"{target_col_letter}{r}"] = f"=order!{source_col_letter}{source_row}"

    # フィルタと枠固定
    filter_end_letter = openpyxl.utils.get_column_letter(65) # BM列まで
    ws_order.auto_filter.ref = f"A3:{filter_end_letter}{lastRow}"
    ws_order.freeze_panes = "A4"

    # --------------------------------------------------------
    # 新規列（BH〜BM列）への追加処理
    # --------------------------------------------------------
    print("  新規列（希望品、HUTTE、ZOZO、OIOI、特価等）のデータを書き込み中...")
    
    # ヘッダ設定
    ws_order["BH3"] = "希望品"
    ws_order["BI3"] = "HUTTE"
    ws_order["BJ3"] = "ZOZO"
    ws_order["BK3"] = "OIOI"
    ws_order["BL3"] = "R"
    ws_order["BM3"] = "特価"
    
    # ヘッダの書式
    for c in range(60, 66):
        cell = ws_order.cell(row=3, column=c)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.column_dimensions[col_letter].width = 8

    # データ行書き込み (4行目〜 current_order_row - 1)
    final_max_row = current_order_row - 1
    max_order_col = max(order_store_map.values()) if order_store_map else 43

    for r in range(4, final_max_row + 1):
        code_val = ws_order.cell(row=r, column=1).value
        code_str = str(code_val).strip() if code_val else ""
        
        # BI: HUTTE
        ws_order.cell(row=r, column=61).value = hutte_dict.get(code_str, "")
        # BJ: ZOZO
        ws_order.cell(row=r, column=62).value = zozo_m_dict.get(code_str, "")
        # BK: OIOI
        ws_order.cell(row=r, column=63).value = oioi_m_dict.get(code_str, "")
        # BL: R (連番)
        ws_order.cell(row=r, column=64).value = r - 3
        # BM: 特価 (0の場合は非表示)
        tokka_val = tokka_dict.get(code_str, "")
        if tokka_val == 0 or tokka_val == "0" or tokka_val == 0.0:
            tokka_val = ""
        ws_order.cell(row=r, column=65).value = tokka_val

        # --------------------------------------------------------
        # 在庫(18〜30)・ORDER(31〜max_order_col) の「0」を非表示(None)にする
        # --------------------------------------------------------
        for c in range(18, max_order_col + 1):
            val = ws_order.cell(row=r, column=c).value
            if val == 0 or val == "0" or val == 0.0:
                ws_order.cell(row=r, column=c).value = None

    print("  マージおよび書式適用完了")

# ============================================================
# メイン処理
# ============================================================
def main():
    print("=" * 50)
    print("出荷明細作成自動化スクリプト 開始 (Python直接集計・ハイブリッド版)")
    print("=" * 50)

    wb_retail = None
    success = False

    try:
        # 各種CSVの直接読み込みと集計の実行
        pivot_df = load_store_sales()
        zozo_dict = load_zozo_sales()
        oioi_dict = load_oioi_sales()
        ec_dict = load_ec_sales()

        # 新規ファイルからのデータ抽出
        hutte_dict = load_hutte()
        zozo_m_dict, oioi_m_dict = load_zozo_master()
        tokka_dict = load_tokka()

        # 出荷予定振分からベースシートを作成
        wb_retail = step1_create_retail_wb()
        
        # 直接集計結果をマージして書式を設定
        step2_python_direct_merge(wb_retail, pivot_df, zozo_dict, oioi_dict, ec_dict, hutte_dict, zozo_m_dict, oioi_m_dict, tokka_dict)
        
        # 最終保存
        retail_xlsx_path = os.path.join(BASE_DIR, "RETAIL_完成版.xlsx")
        if os.path.exists(retail_xlsx_path):
            try:
                os.remove(retail_xlsx_path)
            except PermissionError as pe:
                raise AutomationError(
                    "完成版Excelファイルが開いたままです",
                    "【発生した状況】\n完成版Excel（RETAIL_完成版.xlsx）がExcelなどのアプリケーションで開かれた状態になっています。\n\n"
                    "【対処方法】\n開いている「RETAIL_完成版.xlsx」のExcel画面を完全に閉じてから、もう一度実行してください。"
                ) from pe
            except Exception:
                pass
            
        print(f"\nRETAIL_完成版.xlsx を書き出し中...")
        try:
            wb_retail.save(retail_xlsx_path)
        except PermissionError as pe:
            raise AutomationError(
                "完成版Excelファイルが開いたままです",
                "【発生した状況】\n完成版Excel（RETAIL_完成版.xlsx）がExcelなどのアプリケーションで開かれた状態になっています。\n\n"
                "【対処方法】\n開いている「RETAIL_完成版.xlsx」のExcel画面を完全に閉じてから、もう一度実行してください。"
            ) from pe
            
        wb_retail.close()
        wb_retail = None
        
        print(f"\n処理が正常に完了しました！\n-> {os.path.basename(retail_xlsx_path)}")
        success = True
        
    except AutomationError as ae:
        # 独自例外のハンドリング
        with open(ERROR_LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"エラー分類: {ae.title}\n")
            f.write(f"エラー詳細:\n{ae.message}\n")
            f.write("=" * 60 + "\n")
            f.write(traceback.format_exc())
        print(f"\nエラー: {ae.title}\n{ae.message}\n詳細は error.log を確認してください。")
        show_error(ae.title, ae.message)
        
    except Exception as e:
        # 想定外のエラーのハンドリング
        title = "プログラム実行中にエラーが発生しました"
        msg = f"【発生した状況】\nプログラムの実行中に予期しないエラーが発生しました。\n\nエラー内容: {e}\n\n【対処方法】\n同じフォルダにある「error.log」ファイルの内容をシステム管理者または開発者に提供し、調査を依頼してください。"
        with open(ERROR_LOG_PATH, "w", encoding="utf-8") as f:
            f.write(f"エラー分類: {title}\n")
            f.write(f"エラー詳細:\n{msg}\n")
            f.write("=" * 60 + "\n")
            f.write(traceback.format_exc())
        print(f"\nエラー: {title}\n{msg}\n詳細は error.log を確認してください。")
        show_error(title, msg)
        
    finally:
        if wb_retail is not None:
            try: wb_retail.close()
            except Exception: pass

    if success:
        show_info("処理が正常に完了しました！\n\nRETAIL_完成版.xlsx を作成しました。")

if __name__ == "__main__":
    main()
