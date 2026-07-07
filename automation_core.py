# -*- coding: utf-8 -*-
"""
出荷明細作成自動化 共有モジュール (automation_core.py)
===================================================
共通ロジック、バグ修正、およびリファクタリングを統合したコアモジュールです。
"""

import os
import sys
import glob
import traceback
import tkinter as tk
from tkinter import messagebox
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule

# ============================================================
# 例外の定義
# ============================================================
class AutomationError(Exception):
    """業務自動化処理中に発生する想定内のエラー"""
    def __init__(self, title, message):
        self.title = title
        self.message = message
        super().__init__(message)

# ============================================================
# ブランド設定クラスおよびインスタンス定義
# ============================================================
class BrandConfig:
    def __init__(self, name, allowed_names, prefixes, excluded_prefixes=None):
        self.name = name
        self.allowed_names = [n.upper() for n in allowed_names]
        self.prefixes = tuple(p.upper() for p in prefixes) if prefixes else ()
        self.excluded_prefixes = tuple(p.upper() for p in excluded_prefixes) if excluded_prefixes else ()

BRAND_FRV = BrandConfig(
    name="FRV",
    allowed_names=["FRV", "FJALLRAVEN", "FJÄLLRÄVEN"],
    prefixes=[],  # すべて（除外プレフィックス以外）
    excluded_prefixes=["TNT", "TNP", "TNF", "H", "W"]
)

BRAND_TEN = BrandConfig(
    name="TEN",
    allowed_names=["TEN", "TENNEN"],
    prefixes=["TNT", "TNP", "TNF"],
    excluded_prefixes=[]
)

BRAND_HWG = BrandConfig(
    name="HWG",
    allowed_names=["HWG", "HANWAG"],
    prefixes=["H"],
    excluded_prefixes=[]
)

# ============================================================
# 判定ユーティリティ
# ============================================================
def is_target_brand(config: BrandConfig, code, brand_name=None) -> bool:
    """ブランド設定に基づき、対象ブランドか厳格に判定する。
    ブランド列が存在する場合（brand_nameが渡された場合）は必ずブランド名のみで判定し、
    コードによるフォールバックは行わない。
    """
    # 1. ブランド名が存在する場合 → ブランド名のみで即判定（コード判定へ絶対に落とさない）
    if brand_name is not None and str(brand_name).strip():
        b = str(brand_name).upper().strip()
        return b in config.allowed_names

    # 2. ブランド名がない場合のみコードで判定
    if not code:
        return False

    c = str(code).upper().strip()
    # 除外プレフィックスにマッチすれば False
    if config.excluded_prefixes and c.startswith(config.excluded_prefixes):
        return False

    # プレフィックスリストが空なら、除外プレフィックスに引っかからなかったものはすべて True (FRV用)
    if not config.prefixes:
        return True

    # プレフィックスにマッチすれば True
    if c.startswith(config.prefixes):
        return True

    return False

# ============================================================
# GUI / ダイアログ ユーティリティ
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

# ============================================================
# ファイル検索
# ============================================================
def find_single_csv(input_dir, pattern_keywords, exclude_keywords=None):
    all_csvs = glob.glob(os.path.join(input_dir, "*.csv"))
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
# データ集計処理
# ============================================================
def load_store_sales(input_dir, encoding, brand_config: BrandConfig):
    print("  店舗売上 (営業日付別売上分析) を集計中...")
    src = find_single_csv(input_dir, ["営業日付別売上分析"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)

    # ターゲットブランドのみ抽出
    mask = df.apply(lambda row: is_target_brand(brand_config, row["3rd Item No."], row.get("表記部門名1")), axis=1)
    df_filtered = df[mask].copy()

    # 数量の数値型キャストおよびマイナス値（返品・キャンセル）の0丸め
    df_filtered["数量"] = pd.to_numeric(df_filtered["数量"], errors="coerce").fillna(0)
    df_filtered["数量"] = df_filtered["数量"].clip(lower=0)
    
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
    # ここでも念のためマイナス値を0に丸める
    return max(0, int(val)) if not pd.isna(val) else 0

def load_zozo_sales(input_dir, encoding, brand_config: BrandConfig):
    print("  ZOZO売上 (卸売上明細) を集計中...")
    src = find_single_csv(input_dir, ["卸売上明細"], exclude_keywords=["卸売上明細 (1)", "卸売上明細(1)"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)

    # ターゲットブランドのみ抽出（Brand列 → ブランド名列 の順で参照）
    def get_brand(row):
        return row.get("Brand") or row.get("ブランド名") or None
    mask = df.apply(lambda row: is_target_brand(brand_config, row["商品コード"], get_brand(row)), axis=1)
    df_filtered = df[mask].copy()

    # マイナス値を0に丸める
    df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
    df_filtered["販売数量"] = df_filtered["販売数量"].clip(lower=0)
    
    agg_df = df_filtered.groupby("商品コード")["販売数量"].sum().reset_index()
    return dict(zip(agg_df["商品コード"], agg_df["販売数量"]))

def load_oioi_sales(input_dir, encoding, brand_config: BrandConfig):
    print("  OIOI売上 (卸売上明細 (1)) を集計中...")
    src = find_single_csv(input_dir, ["卸売上明細 (1)"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)

    # ターゲットブランドのみ抽出（Brand列 → ブランド名列 の順で参照）
    if "商品コード" in df.columns:
        def get_brand_oioi(row):
            return row.get("Brand") or row.get("ブランド名") or None
        mask = df.apply(lambda row: is_target_brand(brand_config, row["商品コード"], get_brand_oioi(row)), axis=1)
    else:
        mask = df.apply(lambda row: is_target_brand(brand_config, None, row.get("Brand") or row.get("ブランド名")), axis=1)
    df_filtered = df[mask].copy()

    # マイナス値を0に丸める
    if "販売数量" in df_filtered.columns and "商品コード" in df_filtered.columns:
        df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
        df_filtered["販売数量"] = df_filtered["販売数量"].clip(lower=0)
        agg_df = df_filtered.groupby("商品コード")["販売数量"].sum().reset_index()
        return dict(zip(agg_df["商品コード"], agg_df["販売数量"]))
    return {}

def load_ec_sales(input_dir, encoding, brand_config: BrandConfig):
    print("  EC売上 (order_*) を集計中...")
    src = find_single_csv(input_dir, ["order_"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)

    df = df[~df["決済方法(ステータス)"].str.contains("キャンセル", na=False)]
    df = df[~df["注文者"].str.contains("店舗客注", na=False)]
    df = df.reset_index(drop=True)

    # マイナス値を0に丸める
    df["個数"] = pd.to_numeric(df["個数"], errors="coerce").fillna(0)
    df["個数"] = df["個数"].clip(lower=0)
    
    key_col = "オプション独自コード" if "オプション独自コード" in df.columns else "商品コード"
    
    # ターゲットブランドのみ抽出
    brand_col = "ブランド名" if "ブランド名" in df.columns else None
    if brand_col:
        mask = df.apply(lambda row: is_target_brand(brand_config, row[key_col], row[brand_col]), axis=1)
    else:
        mask = df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))
    df = df[mask]

    agg_df = df.groupby(key_col)["個数"].sum().reset_index()
    return dict(zip(agg_df[key_col], agg_df["個数"]))

# ============================================================
# 新規データ抽出用ロード関数 (堅牢化＆列インデックスハードコーディング廃止)
# ============================================================
def load_hutte(input_dir, brand_config: BrandConfig):
    print("  HUTTE出荷指示 (★出荷指示フォーマット【*】hutte.xlsx) を読み込み中...")
    pattern = os.path.join(input_dir, "★出荷指示フォーマット【*】hutte.xlsx")
    files = glob.glob(pattern)
    if not files:
        print("  [情報] HUTTE出荷指示ファイルが見つかりません。")
        return {}
    latest_file = max(files, key=os.path.getmtime)
    try:
        df = pd.read_excel(latest_file, sheet_name=0, dtype=str)
        
        # 列の動的特定
        key_col = None
        val_col = None
        for col in df.columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            if any(kw in col_lower for kw in ["商品コード", "品番", "アイテムコード", "sku", "code"]):
                key_col = col
                break
        if key_col is None and len(df.columns) > 3:
            key_col = df.columns[3]
            
        for col in df.columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            if any(kw in col_lower for kw in ["数量", "指示数", "出荷数", "個数", "枚数", "qty", "quantity"]):
                val_col = col
                break
        if val_col is None and len(df.columns) > 12:
            val_col = df.columns[12]

        if key_col is not None and val_col is not None:
            df = df.dropna(subset=[key_col])
            # キー値トリム
            df[key_col] = df[key_col].astype(str).str.strip()
            # ターゲットブランドのみ抽出
            df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
            # 数量を数値型キャストしマイナス値は0丸め
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0).astype(int)
            df[val_col] = df[val_col].clip(lower=0)
            
            # 同一コードの groupby 合算
            agg_df = df.groupby(key_col)[val_col].sum().reset_index()
            return dict(zip(agg_df[key_col], agg_df[val_col]))
    except Exception as e:
        print(f"  [警告] HUTTE出荷指示の読み込みに失敗しました: {e}")
    return {}

def load_zozo_master(input_dir, brand_config: BrandConfig):
    print("  ZOZO/OIOIマスタ (商品マスタ-ZOZO-OlOl.xlsx) を読み込み中...")
    path = os.path.join(input_dir, "商品マスタ-ZOZO-OlOl.xlsx")
    if not os.path.exists(path):
        print("  [情報] ZOZO/OIOIマスタファイルが見つかりません。")
        return {}, {}
    try:
        # header=1 (2行目)
        df = pd.read_excel(path, sheet_name=0, header=1, dtype=str)
        
        # 列の動的特定
        key_col = None
        zozo_col = None
        oioi_col = None
        for col in df.columns:
            col_str = str(col).strip()
            col_lower = col_str.lower()
            if "商品コード" in col_str or "品番" in col_str:
                key_col = col
            elif "zozo" in col_lower:
                zozo_col = col
            elif "oioi" in col_lower or "丸井" in col_str:
                oioi_col = col

        if key_col is None:
            key_col = df.columns[0]
        if zozo_col is None:
            zozo_col = df.columns[26] if len(df.columns) > 26 else None
        if oioi_col is None:
            oioi_col = df.columns[27] if len(df.columns) > 27 else None

        if key_col is not None:
            df = df.dropna(subset=[key_col])
            df[key_col] = df[key_col].astype(str).str.strip()
            # ターゲットブランドのみ抽出
            df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
            
            z_dict = {}
            o_dict = {}
            if zozo_col is not None:
                # 数量数値化とマイナス丸め
                df_zozo_qty = pd.to_numeric(df[zozo_col], errors="coerce").fillna(0).astype(int).clip(lower=0)
                z_dict = dict(zip(df[key_col], df_zozo_qty))
            if oioi_col is not None:
                df_oioi_qty = pd.to_numeric(df[oioi_col], errors="coerce").fillna(0).astype(int).clip(lower=0)
                o_dict = dict(zip(df[key_col], df_oioi_qty))
                
            return z_dict, o_dict
    except Exception as e:
        print(f"  [警告] ZOZO/OIOIマスタの読み込みに失敗しました: {e}")
    return {}, {}

def load_tokka(input_dir, brand_config: BrandConfig):
    print("  特価在庫 (2026特価在庫.xlsx) を読み込み中...")
    path = os.path.join(input_dir, "2026特価在庫.xlsx")
    if not os.path.exists(path):
        print("  [情報] 特価在庫ファイルが見つかりません。")
        return {}
    try:
        df = pd.read_excel(path, sheet_name=0, dtype=str)
        
        # 列の動的特定
        key_col = None
        val_col = None
        for col in df.columns:
            col_str = str(col).strip()
            if "商品コード" in col_str or "品番" in col_str:
                key_col = col
            elif "特価" in col_str:
                val_col = col

        if key_col is None and len(df.columns) > 2:
            key_col = df.columns[2]
        if val_col is None and len(df.columns) > 22:
            val_col = df.columns[22]

        if key_col is not None and val_col is not None:
            df = df.dropna(subset=[key_col])
            df[key_col] = df[key_col].astype(str).str.strip()
            # ターゲットブランドのみ抽出
            df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
            
            # 特価数を数値キャストしマイナス値は0丸め
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0).astype(int).clip(lower=0)
            
            return dict(zip(df[key_col], df[val_col]))
    except Exception as e:
        print(f"  [警告] 特価在庫の読み込みに失敗しました: {e}")
    return {}

# ============================================================
# 手順1: 出荷予定振分.csv から RETAIL ワークブックの作成
# ============================================================
def step1_create_retail_wb(input_dir, encoding, brand_config: BrandConfig):
    print("[手順1] 出荷予定振分.csv から RETAILシートの構築")
    src = find_single_csv(input_dir, ["出荷予定振分"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)
    
    # ======== 自ブランドのみ抽出フィルタリング ========
    if len(df) > 3:
        df_headers = df.iloc[:3].copy()
        df_data = df.iloc[3:].copy()
        
        # 4列目(iloc[3])が商品コード、6列目(iloc[5])がBrand
        mask = df_data.apply(lambda row: is_target_brand(brand_config, row.iloc[3], row.iloc[5]), axis=1)
        df_data = df_data[mask].copy()
        
        df = pd.concat([df_headers, df_data], ignore_index=True)
    # ================================================
    
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
# 手順2: Python直接集計・マージおよび書式適用
# ============================================================
def step2_python_direct_merge(wb_retail, base_dir, input_dir, template_path, brand_config: BrandConfig,
                              pivot_df, zozo_dict, oioi_dict, ec_dict, hutte_dict,
                              zozo_m_dict, oioi_m_dict, tokka_dict):
    print("[手順2] Pythonによる直接マージ＆書式適用 (VBA代替処理)")
    
    if "order" in wb_retail.sheetnames:
        del wb_retail["order"]

    ws_retail = wb_retail["RETAIL"]
    
    # AW1:BY[max_row] をクリア (10000行ループを実際の最終行に削減)
    max_r_retail = max(ws_retail.max_row, 1)
    for r in range(1, max_r_retail + 1):
        for c in range(49, 78):
            ws_retail.cell(row=r, column=c).value = None
            
    # シートの複製
    ws_order = wb_retail.copy_worksheet(ws_retail)
    ws_order.title = "order"

    # "商品コード" "商品名" "ブランド名" の列を動的に特定
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
    ws_order.delete_rows(3, 1) # 3行目を削除 (店舗コード行を削除)
    
    # 1. BULK不要列の削除 (R:X 削除。元の18列目から7列分)
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

    # AR:BD列(44〜56列目)をクリア
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
    wb_temp = openpyxl.load_workbook(template_path, data_only=True)
    ws_temp = wb_temp["計算式2026"]
    
    # 3行目の店舗名をテンプレートからコピーして order シートの R3:BG3 に適用
    for c in range(18, 60):
        ws_order.cell(row=3, column=c).value = ws_temp.cell(row=3, column=c).value

    wb_temp.close()

    # --------------------------------------------------------
    # 列インデックスを動的に検索・解決する仕組みの構築
    # --------------------------------------------------------
    col_map = {}
    current_area = None
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
    master_dict = {}
    master_path = os.path.join(base_dir, "商品マスタ全項目.csv")
    if os.path.exists(master_path):
        try:
            df_m = pd.read_csv(master_path, encoding="cp932", dtype=str)
            # 列の動的特定 (バグ6対応)
            code_col = None
            name_col = None
            for col in df_m.columns:
                col_str = str(col).strip()
                if "商品コード" in col_str or "品番" in col_str:
                    code_col = col
                elif "商品名" in col_str or "品名" in col_str:
                    name_col = col
            if code_col is None: code_col = df_m.columns[0]
            if name_col is None: name_col = df_m.columns[2] if len(df_m.columns) > 2 else df_m.columns[1]

            for _, row in df_m.iterrows():
                m_code = str(row[code_col]).strip() if not pd.isna(row[code_col]) else ""
                m_name = str(row[name_col]).strip() if not pd.isna(row[name_col]) else ""
                if m_code:
                    master_dict[m_code] = {
                        'name': m_name,
                        'code': m_code
                    }
        except Exception as e:
            print(f"  警告: 商品マスタのロードに失敗しました: {e}")

    # --------------------------------------------------------
    # 売上CSVからバックアップ用の「コード ↔ 商品名」を取得
    # --------------------------------------------------------
    backup_names = {}
    try:
        src_store = find_single_csv(input_dir, ["営業日付別売上分析"])
        df_store_raw = pd.read_csv(src_store, encoding="cp932", dtype=str)
        for _, row in df_store_raw.iterrows():
            code = str(row["3rd Item No."]).strip() if not pd.isna(row["3rd Item No."]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    try:
        src_zozo = find_single_csv(input_dir, ["卸売上明細"], exclude_keywords=["卸売上明細 (1)", "卸売上明細(1)"])
        df_zozo_raw = pd.read_csv(src_zozo, encoding="cp932", dtype=str)
        for _, row in df_zozo_raw.iterrows():
            code = str(row["商品コード"]).strip() if not pd.isna(row["商品コード"]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    try:
        src_oioi = find_single_csv(input_dir, ["卸売上明細 (1)"])
        df_oioi_raw = pd.read_csv(src_oioi, encoding="cp932", dtype=str)
        for _, row in df_oioi_raw.iterrows():
            code = str(row["商品コード"]).strip() if not pd.isna(row["商品コード"]) else ""
            name = str(row["商品名"]).strip() if "商品名" in row and not pd.isna(row["商品名"]) else ""
            if code and name:
                backup_names[code] = name
    except Exception: pass

    try:
        src_ec = find_single_csv(input_dir, ["order_"])
        df_ec_raw = pd.read_csv(src_ec, encoding="cp932", dtype=str)
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
        if qty != 0 and "GIFT" not in code:
            all_sales_codes.add(code)
    if pivot_df is not None and not pivot_df.empty:
        num_cols = [c for c in pivot_df.columns if c != "3rd Item No."]
        for _, row in pivot_df.iterrows():
            code = str(row["3rd Item No."]).strip() if not pd.isna(row["3rd Item No."]) else ""
            if code:
                total_qty = pd.to_numeric(row[num_cols], errors="coerce").fillna(0).sum()
                if total_qty != 0:
                    all_sales_codes.add(code)

    # 出荷予定振分の現在のA列コード
    order_codes = set()
    for r in range(4, lastRow + 1):
        val = ws_order.cell(row=r, column=1).value
        if val:
            order_codes.add(str(val).strip())

    # 漏れている（出荷予定にないが売上があった）商品コードの抽出 (バグ2修正)
    leaked_codes = sorted(list(all_sales_codes - order_codes))
    print(f"  出荷予定にない売上発生商品を追加中... (追加数: {len(leaked_codes)})")

    # --------------------------------------------------------
    # orderシートおよびRETAILシートの末尾への行追加
    # --------------------------------------------------------
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
                val = int(zozo_dict.get(code, 0))
            elif store_name == "OIOI":
                val = int(oioi_dict.get(code, 0))
            elif store_name in ("YSEC", "EC"):
                val = int(ec_dict.get(code, 0))
            else:
                val = get_store_sales_value(pivot_df, code, store_name)
            # マイナスは0に書き換え
            ws_order.cell(row=current_order_row, column=col_idx).value = max(0, val)

        # RETAILシートへ行追加
        ws_retail.cell(row=current_retail_row, column=4).value = p_code
        ws_retail.cell(row=current_retail_row, column=5).value = p_name
        ws_retail.cell(row=current_retail_row, column=6).value = brand_config.name
        ws_retail.cell(row=current_retail_row, column=7).value = 0
        for col in range(8, 29):
            ws_retail.cell(row=current_retail_row, column=col).value = 0
            
        # RETAILシートからorderシートへの参照式を設定 (ORDERエリアの13列分)
        for c_idx in range(13):
            target_col_letter = openpyxl.utils.get_column_letter(49 + c_idx)
            source_col_letter = openpyxl.utils.get_column_letter(31 + c_idx)
            ws_retail.cell(row=current_retail_row, column=49 + c_idx).value = f"=order!{source_col_letter}{current_order_row}"

        current_order_row += 1
        current_retail_row += 1

    # 最終行変数を更新
    lastRow = current_order_row - 1
    lastRowRetail = current_retail_row - 1

    # 在庫エリアの左端と右端の動的な罫線適用
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
    for r in range(4, lastRow + 1):  # 全オリジナル行を対象にする（leaked_codesは lastRow+1 以降に追加済み）
        code_short = ws_order.cell(row=r, column=1).value
        if not code_short:
            for store_name, col_idx in order_store_map.items():
                ws_order.cell(row=r, column=col_idx).value = 0
            continue
            
        code_short = str(code_short).strip()

        # 売上値の事前計算
        raw_sales = {}
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
            # 各チャネルのロード時点で丸めているが、ここでもマイナスを0に丸める
            raw_sales[store_name] = max(0, qty)

        # 誤発注防止ロジック (バグ4対応・仕様修正)
        # 全チャネルの売上合計で在庫不足を判定し、不足時は全チャネル（ZOZO/OIOI/EC/YSEC含む）をゼロにする
        total_sales_all = sum(raw_sales.values())

        c_val = ws_order.cell(row=r, column=3).value or 0
        d_val = ws_order.cell(row=r, column=4).value or 0
        kabu_sum = sum(ws_order.cell(row=r, column=col).value or 0 for col in range(5, 18))

        if c_val == 0 and d_val == 0 and kabu_sum < total_sales_all:
            # 確保枠が全売上合計を下回る場合、全チャネルの売上を 0 にする
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
        for s in stores_list:
            b_col = bulk_store_map.get(s)
            o_col = order_store_map.get(s)
            k_col = kabusoku_store_map.get(s)
            if b_col and o_col and k_col:
                b_let = openpyxl.utils.get_column_letter(b_col)
                o_let = openpyxl.utils.get_column_letter(o_col)
                ws_order.cell(row=r, column=k_col).value = f"={b_let}{r}-{o_let}{r}"
            
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
        nwa_formula = f"IF({let_retail_kakubo}{r}<0,{let_c}{r}+{let_retail_kakubo}{r},{let_c}{r})"
        ws_order.cell(row=r, column=col_kabusoku_zan).value = f'=IF({nwa_formula}=0,"",{nwa_formula})'

    # 1行目の SUBTOTAL 計算式の書き込み (3500固定の廃止、lastRowへの動的置換 - バグ7対応)
    for c in range(1, ws_order.max_column + 1):
        ws_order.cell(row=1, column=c).value = None
        
    max_order_col = max(order_store_map.values()) if order_store_map else 43
    for c in range(3, max_order_col + 1):
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.cell(row=1, column=c).value = f"=SUBTOTAL(9,{col_letter}4:{col_letter}{lastRow})"

    # 幅と非表示設定
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
    ws_order[f"{zan_letter}3"].fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

    # 条件付き書式
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    rule = FormulaRule(formula=[f"${zan_letter}4<0"], stopIfTrue=False, fill=red_fill)
    ws_order.conditional_formatting.add(f"A4:A{lastRow}", rule)

    # RETAIL シートへの数式展開
    lastRowRetail = 5
    for r in range(ws_retail.max_row, 4, -1):
        if ws_retail.cell(row=r, column=4).value is not None:
            lastRowRetail = r
            break
            
    for r in range(5, lastRowRetail + 1):
        for c_idx in range(13):
            target_col_letter = openpyxl.utils.get_column_letter(49 + c_idx)
            source_col_letter = openpyxl.utils.get_column_letter(31 + c_idx)
            source_row = r - 1
            ws_retail[f"{target_col_letter}{r}"] = f"=order!{source_col_letter}{source_row}"

    # フィルタと枠固定
    filter_end_letter = openpyxl.utils.get_column_letter(65)
    ws_order.auto_filter.ref = f"A3:{filter_end_letter}{lastRow}"
    ws_order.freeze_panes = "A4"

    # BH〜BM列へのデータ書き込み
    print("  新規列（希望品、HUTTE、ZOZO、OIOI、特価等）のデータを書き込み中...")
    ws_order["BH3"] = "希望品"
    ws_order["BI3"] = "HUTTE"
    ws_order["BJ3"] = "ZOZO"
    ws_order["BK3"] = "OIOI"
    ws_order["BL3"] = "R"
    ws_order["BM3"] = "特価"
    
    for c in range(60, 66):
        cell = ws_order.cell(row=3, column=c)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        col_letter = openpyxl.utils.get_column_letter(c)
        ws_order.column_dimensions[col_letter].width = 8

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
        # BL: R
        ws_order.cell(row=r, column=64).value = r - 3
        
        # BM: 特価
        tokka_val = tokka_dict.get(code_str, "")
        if tokka_val == 0 or tokka_val == "0" or tokka_val == 0.0:
            tokka_val = ""
        ws_order.cell(row=r, column=65).value = tokka_val

        # C, D, 在庫(18〜30), ORDER(31〜max_order_col) の「0」を非表示(None)にする
        for c in [3, 4] + list(range(18, max_order_col + 1)):
            val = ws_order.cell(row=r, column=c).value
            if val == 0 or val == "0" or val == 0.0:
                ws_order.cell(row=r, column=c).value = None

    print("  マージおよび書式適用完了")

# ============================================================
# メインパイプラインの実行
# ============================================================
def run_pipeline(brand_config: BrandConfig, output_file_name, input_dir, base_dir, template_path, error_log_path):
    print("=" * 50)
    print(f"出荷明細作成自動化処理 開始 [{brand_config.name}用] (リファクタリング版)")
    print("=" * 50)

    wb_retail = None
    success = False
    encoding = "cp932"

    try:
        # 各種CSVの直接読み込みと集計の実行
        pivot_df = load_store_sales(input_dir, encoding, brand_config)
        zozo_dict = load_zozo_sales(input_dir, encoding, brand_config)
        oioi_dict = load_oioi_sales(input_dir, encoding, brand_config)
        ec_dict = load_ec_sales(input_dir, encoding, brand_config)

        # 新規ファイルからのデータ抽出
        hutte_dict = load_hutte(input_dir, brand_config)
        zozo_m_dict, oioi_m_dict = load_zozo_master(input_dir, brand_config)
        tokka_dict = load_tokka(input_dir, brand_config)

        # 出荷予定振分からベースシートを作成
        wb_retail = step1_create_retail_wb(input_dir, encoding, brand_config)
        
        # 直接集計結果をマージして書式を設定
        step2_python_direct_merge(
            wb_retail, base_dir, input_dir, template_path, brand_config,
            pivot_df, zozo_dict, oioi_dict, ec_dict, hutte_dict,
            zozo_m_dict, oioi_m_dict, tokka_dict
        )
        
        # 最終保存
        output_dir = os.path.join(base_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        retail_xlsx_path = os.path.join(output_dir, output_file_name)
        
        if os.path.exists(retail_xlsx_path):
            try:
                os.remove(retail_xlsx_path)
            except PermissionError as pe:
                raise AutomationError(
                    "完成版Excelファイルが開いたままです",
                    f"【発生した状況】\n完成版Excel（{output_file_name}）がExcelなどのアプリケーションで開かれた状態になっています。\n\n"
                    f"【対処方法】\n開いている「{output_file_name}」のExcel画面を完全に閉じてから、もう一度実行してください。"
                ) from pe
            except Exception:
                pass
            
        print(f"\n{output_file_name} を書き出し中...")
        try:
            wb_retail.save(retail_xlsx_path)
        except PermissionError as pe:
            raise AutomationError(
                "完成版Excelファイルが開いたままです",
                f"【発生した状況】\n完成版Excel（{output_file_name}）がExcelなどのアプリケーションで開かれた状態になっています。\n\n"
                f"【対処方法】\n開いている「{output_file_name}」のExcel画面を完全に閉じてから、もう一度実行してください。"
            ) from pe
            
        wb_retail.close()
        wb_retail = None
        
        print(f"\n処理が正常に完了しました！\n-> {os.path.basename(retail_xlsx_path)}")
        success = True
        
    except AutomationError as ae:
        with open(error_log_path, "w", encoding="utf-8") as f:
            f.write(f"エラー分類: {ae.title}\n")
            f.write(f"エラー詳細:\n{ae.message}\n")
            f.write("=" * 60 + "\n")
            f.write(traceback.format_exc())
        print(f"\nエラー: {ae.title}\n{ae.message}\n詳細は error.log を確認してください。")
        show_error(ae.title, ae.message)
        
    except Exception as e:
        title = "プログラム実行中にエラーが発生しました"
        msg = f"【発生した状況】\nプログラムの実行中に予期しないエラーが発生しました。\n\nエラー内容: {e}\n\n【対処方法】\n同じフォルダにある「error.log」ファイルの内容をシステム管理者または開発者に提供し、調査を依頼してください。"
        with open(error_log_path, "w", encoding="utf-8") as f:
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
        show_info(f"処理が正常に完了しました！\n\n{output_file_name} を作成しました。")

# ============================================================
# データ検証用 共通パイプライン
# ============================================================
def verify_pipeline(brand_config: BrandConfig, input_dir: str, result_path: str) -> bool:
    """元CSVと生成されたExcel(orderシート)の売上データを商品・チャネルレベルで網羅的に検証する"""
    import pandas as pd
    import openpyxl
    import openpyxl.utils
    import glob
    import os

    print("=" * 55)
    print(f"  検証対象ファイル: {os.path.basename(result_path)}")
    print(f"  対象ブランド: {brand_config.name}")
    print("=" * 55)

    if not os.path.exists(result_path):
        print(f"[エラー] 検証対象のExcelファイルが見つかりません: {result_path}")
        return False

    if not os.path.exists(input_dir):
        print(f"[エラー] 入力データディレクトリが見つかりません: {input_dir}")
        return False

    def find_csv_local(keywords, exclude=None):
        clean_keywords = [k.replace(" ", "").replace("　", "") for k in keywords]
        clean_exclude = [e.replace(" ", "").replace("　", "") for e in exclude] if exclude else []
        for p in glob.glob(os.path.join(input_dir, "*.csv")):
            name = os.path.basename(p)
            clean_name = name.replace(" ", "").replace("　", "")
            if all(k in clean_name for k in clean_keywords):
                if clean_exclude and any(e in clean_name for e in clean_exclude):
                    continue
                return p
        return None

    # CSVファイルの検索
    zozo_csv = find_csv_local(["卸売上明細"], exclude=["卸売上明細 (1)", "卸売上明細(1)"])
    oioi_csv = find_csv_local(["卸売上明細 (1)"])
    store_csv = find_csv_local(["営業日付別売上分析"])
    ec_csv = find_csv_local(["order_"])

    if not zozo_csv:
        print("[エラー] 卸売上明細 CSV が見つかりません。")
        return False
    if not store_csv:
        print("[エラー] 営業日付別売上分析 CSV が見つかりません。")
        return False
    if not ec_csv:
        print("[エラー] order_*.csv が見つかりません。")
        return False

    # 1. 元CSVデータのロードと集計
    # ZOZO
    df_zozo = pd.read_csv(zozo_csv, encoding="cp932", dtype=str)
    def _get_brand_z(row): return row.get("Brand") or row.get("ブランド名") or None
    mask = df_zozo.apply(lambda row: is_target_brand(brand_config, row["商品コード"], _get_brand_z(row)), axis=1)
    df_zozo = df_zozo[mask].copy()
    df_zozo["販売数量"] = pd.to_numeric(df_zozo["販売数量"], errors="coerce").fillna(0).clip(lower=0)
    zozo_sales = df_zozo.groupby("商品コード")["販売数量"].sum().to_dict()

    # OIOI
    oioi_sales = {}
    if oioi_csv:
        df_oioi = pd.read_csv(oioi_csv, encoding="cp932", dtype=str)
        def _get_brand_o(row): return row.get("Brand") or row.get("ブランド名") or None
        if "商品コード" in df_oioi.columns:
            mask = df_oioi.apply(lambda row: is_target_brand(brand_config, row["商品コード"], _get_brand_o(row)), axis=1)
        else:
            mask = df_oioi.apply(lambda row: is_target_brand(brand_config, None, _get_brand_o(row)), axis=1)
        df_oioi = df_oioi[mask].copy()
        df_oioi["販売数量"] = pd.to_numeric(df_oioi["販売数量"], errors="coerce").fillna(0).clip(lower=0)
        if "商品コード" in df_oioi.columns:
            oioi_sales = df_oioi.groupby("商品コード")["販売数量"].sum().to_dict()

    # 店舗
    df_store = pd.read_csv(store_csv, encoding="cp932", dtype=str)
    mask = df_store.apply(lambda row: is_target_brand(brand_config, row["3rd Item No."], row.get("表記部門名1")), axis=1)
    df_store = df_store[mask].copy()
    df_store["数量"] = pd.to_numeric(df_store["数量"], errors="coerce").fillna(0).clip(lower=0)
    store_pivot = df_store.pivot_table(index="3rd Item No.", columns="店舗名称", values="数量", aggfunc="sum").fillna(0)

    # EC
    df_ec = pd.read_csv(ec_csv, encoding="cp932", dtype=str)
    df_ec = df_ec[~df_ec["決済方法(ステータス)"].str.contains("キャンセル", na=False)]
    df_ec = df_ec[~df_ec["注文者"].str.contains("店舗客注", na=False)]
    df_ec["個数"] = pd.to_numeric(df_ec["個数"], errors="coerce").fillna(0).clip(lower=0)
    key_col = "オプション独自コード" if "オプション独自コード" in df_ec.columns else "商品コード"
    df_ec = df_ec[~df_ec[key_col].fillna("").str.contains("GIFT", na=False)]
    brand_col = "ブランド名" if "ブランド名" in df_ec.columns else None
    if brand_col:
        mask = df_ec.apply(lambda row: is_target_brand(brand_config, row[key_col], row[brand_col]), axis=1)
    else:
        mask = df_ec[key_col].apply(lambda c: is_target_brand(brand_config, c, None))
    df_ec = df_ec[mask]
    ec_sales = df_ec.groupby(key_col)["個数"].sum().to_dict()

    # Excelワークブックの読み込み
    wb = openpyxl.load_workbook(result_path, data_only=True)
    if "order" not in wb.sheetnames:
        print(f"[エラー] Excelファイルに 'order' シートが存在しません: {result_path}")
        wb.close()
        return False
    ws = wb["order"]

    # 2. 前提データのエラーチェック (事前に実行)
    formula_errors = []
    for r in range(4, ws.max_row + 1):
        code_val = ws.cell(row=r, column=1).value
        if not code_val:
            continue
        for col in range(3, 18):
            val = ws.cell(row=r, column=col).value
            if val is not None:
                val_str = str(val).strip()
                if val_str.startswith("#"):
                    col_letter = openpyxl.utils.get_column_letter(col)
                    col_header = ws.cell(row=3, column=col).value or ""
                    formula_errors.append(
                        f"行 {r}, 列 {col_letter} ({col_header}): エラー値 '{val_str}'"
                    )

    if formula_errors:
        print("\n[WARNING/ERROR] 前提データに数式エラーが検出されました:")
        for err in formula_errors:
            print(f"  - {err}")
        wb.close()
        return False

    # 3. Excel列インデックスの動的解決 (列ズレ脆弱性の廃止)
    stores_list = ["名古屋", "TOKYO", "ルクア大阪", "ヒュッテ", "京王新宿", "大丸心斎橋", "6142", "玉川高島屋", "NODE", "NARITA"]
    channels = ["ZOZO", "OIOI", "EC"] + stores_list
    col_map = {}

    current_area = None
    for col in range(1, ws.max_column + 1):
        area_val = ws.cell(row=2, column=col).value
        if area_val is not None:
            area_val_str = str(area_val).strip()
            if area_val_str in ("確保", "在庫", "ORDER", "過不足"):
                current_area = area_val_str

        if current_area != "ORDER":
            continue

        val = ws.cell(row=3, column=col).value
        if not val:
            continue
        val_str = str(val).strip()

        # チャネルごとに判定
        if "ZOZO" in val_str:
            col_map["ZOZO"] = col
        elif "OIOI" in val_str or "丸井" in val_str:
            col_map["OIOI"] = col
        elif "EC" in val_str:
            col_map["EC"] = col
        else:
            for s in stores_list:
                if s == "TOKYO":
                    if "TOKYO" in val_str and "NODE" not in val_str:
                        col_map["TOKYO"] = col
                elif s == "ヒュッテ":
                    if "ヒュッテ" in val_str or "HUTTE" in val_str.upper():
                        col_map["ヒュッテ"] = col
                elif s == "大丸心斎橋":
                    if "大丸" in val_str or "心斎橋" in val_str:
                        col_map["大丸心斎橋"] = col
                elif s == "京王新宿":
                    if "京王" in val_str or "新宿" in val_str:
                        col_map["京王新宿"] = col
                elif s == "玉川高島屋":
                    if "玉川" in val_str or "高島屋" in val_str:
                        col_map["玉川高島屋"] = col
                elif s == "ルクア大阪":
                    if "大阪" in val_str or "ルクア" in val_str:
                        col_map["ルクア大阪"] = col
                elif s == "NARITA":
                    if "NARITA" in val_str or "成田" in val_str:
                        col_map["NARITA"] = col
                else:
                    if s in val_str:
                        col_map[s] = col

    missing_channels = [c for c in channels if c not in col_map]
    if missing_channels:
        print(f"[エラー] 以下のチャネルがExcelのヘッダー行から特定できませんでした: {missing_channels}")
        wb.close()
        return False

    # 4. 売上ゼロ化ロジックの適用と正解値（期待値）の算出
    resolved_sales = {c: {} for c in channels}

    for r in range(4, ws.max_row + 1):
        code = ws.cell(row=r, column=1).value
        if not code:
            continue
        code = str(code).strip()

        raw_vals = {}
        raw_vals["ZOZO"] = int(zozo_sales.get(code, 0))
        raw_vals["OIOI"] = int(oioi_sales.get(code, 0))
        raw_vals["EC"] = int(ec_sales.get(code, 0))

        # 店舗
        for s in stores_list:
            if s == "TOKYO":
                cols_match = [c for c in store_pivot.columns if "TOKYO" in c and "NODE" not in c]
            elif s == "ヒュッテ":
                cols_match = [c for c in store_pivot.columns if "ヒュッテ" in c or "HUTTE" in c.upper()]
            else:
                cols_match = [c for c in store_pivot.columns if s in c]

            qty = 0
            if cols_match and code in store_pivot.index:
                qty = int(store_pivot.loc[code, cols_match].sum())
            raw_vals[s] = max(0, qty)

        # 全チャネルの売上合計で在庫不足を判定（仕様修正：全チャネルをゼロ化）
        total_sales_all = sum(max(0, raw_vals[k]) for k in channels)

        # 在庫情報
        c_val = ws.cell(row=r, column=3).value or 0
        d_val = ws.cell(row=r, column=4).value or 0
        kabu_sum = sum(ws.cell(row=r, column=col).value or 0 for col in range(5, 18))

        # 不足による売上ゼロ化（ZOZO/OIOI/EC/YSEC含む全チャネルをゼロにする）
        if c_val == 0 and d_val == 0 and kabu_sum < total_sales_all:
            for k in channels:
                resolved_sales[k][code] = 0
        else:
            for k in channels:
                resolved_sales[k][code] = max(0, raw_vals[k])

    # 正解合計値
    zozo_true = sum(resolved_sales["ZOZO"].values())
    oioi_true = sum(resolved_sales["OIOI"].values())
    ec_true = sum(resolved_sales["EC"].values())
    store_totals = {s: sum(resolved_sales[s].values()) for s in stores_list}

    # 5. Excel上の実際値の集計 (動的インデックス使用)
    def col_total(col_num):
        return sum(
            (ws.cell(row=r, column=col_num).value or 0)
            for r in range(4, ws.max_row + 1)
            if ws.cell(row=r, column=2).value is not None
        )

    actual = {c: col_total(col_map[c]) for c in channels}

    # 6. 商品レベル（行レベル）の検証
    diff_details = []
    for r in range(4, ws.max_row + 1):
        code = ws.cell(row=r, column=1).value
        if not code:
            continue
        code = str(code).strip()

        for c in channels:
            truth_val = resolved_sales[c].get(code, 0)
            col_idx = col_map[c]
            actual_val = ws.cell(row=r, column=col_idx).value or 0

            # 不一致があった場合
            if truth_val != actual_val:
                diff_details.append({
                    "row": r,
                    "code": code,
                    "channel": c,
                    "expected": truth_val,
                    "actual": actual_val
                })

    wb.close()

    # 7. 検証結果の出力
    print("\n" + "=" * 55)
    print("[1] 元CSV の合計（正解値）")
    print("=" * 55)
    print(f"  ZOZO    (卸売上明細.csv):          {zozo_true:>6}")
    print(f"  OIOI    (卸売上明細 (1).csv):      {oioi_true:>6}")
    for name, total in store_totals.items():
        print(f"  {name:<12}(営業日付別売上分析.csv): {total:>6}")
    print(f"  EC      (order_*.csv ※GIFT除く):  {ec_true:>6}")

    print("\n" + "=" * 55)
    print(f"[2] Excel の ORDER欄合計（実際値）")
    print("=" * 55)
    for c in channels:
        print(f"  {c:<12}: {actual[c]:>6}")

    print("\n" + "=" * 55)
    print("[3] 突き合わせ結果（正解値 vs 実際値）")
    print("=" * 55)

    all_ok = True

    checks = [
        ("ZOZO", zozo_true, actual["ZOZO"]),
        ("OIOI", oioi_true, actual["OIOI"]),
        ("EC", ec_true, actual["EC"])
    ] + [(s, store_totals[s], actual[s]) for s in stores_list]

    for name, truth, got in checks:
        ok = (truth == got)
        mark = "[OK]" if ok else "[NG] 合計不一致！"
        print(f"  {name:<12}: 正解={truth:>5}  実際={got:>5}  {mark}")
        if not ok:
            all_ok = False

    # 商品レベルのチェック結果の出力 (差異分析レポートの表示)
    if diff_details:
        all_ok = False
        print("\n" + "!" * 55)
        print(f"[NG] 商品レベル（セル単位）の不一致が {len(diff_details)} 件検出されました！ (差異分析レポート)")
        print("!" * 55)
        print(f"{'商品コード':<15} | {'チャネル/店舗':<10} | {'期待値(CSV)':<8} | {'実際値(Excel)':<9} | {'行番号':<5}")
        print("-" * 55)
        for diff in diff_details[:50]:
            print(f"{diff['code']:<15} | {diff['channel']:<10} | {diff['expected']:>8} | {diff['actual']:>9} | {diff['row']:>5}")
        if len(diff_details) > 50:
            print(f"... (他 {len(diff_details) - 50} 件の差異があります)")

    return all_ok

