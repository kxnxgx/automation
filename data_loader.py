# -*- coding: utf-8 -*-
"""
データロードおよび集計モジュール (data_loader.py)
==============================================
"""

import os
import glob
import pandas as pd
from config import AutomationError, BrandConfig
from utils import find_single_csv, is_target_brand, get_brand_name_safe

# ============================================================
# データ集計処理
# ============================================================
def load_store_sales(input_dir, encoding, brand_config: BrandConfig):
    print("  店舗売上 (営業日付別売上分析) を集計中...")
    src = find_single_csv(input_dir, ["営業日付別売上分析"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)
    if df.empty:
        return None

    # ターゲットブランドのみ抽出
    mask = df.apply(lambda row: is_target_brand(brand_config, row["3rd Item No."], row.get("表記部門名1")), axis=1)
    df_filtered = df[mask].copy()

    # 数量の数値型キャストおよびマイナス値（返品・キャンセル）の0丸め
    df_filtered["数量"] = pd.to_numeric(df_filtered["数量"], errors="coerce").fillna(0)
    df_filtered["数量"] = df_filtered["数量"].clip(lower=0)
    
    pivot_df = df_filtered.pivot_table(index="3rd Item No.", columns="店舗名称", values="数量", aggfunc="sum").reset_index()
    return pivot_df

def get_store_sales_value(pivot_df, code_short, store_name):
    if pivot_df is None or pivot_df.empty or "3rd Item No." not in pivot_df.columns:
        return 0
    if code_short not in pivot_df["3rd Item No."].values:
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
    return max(0, int(val)) if not pd.isna(val) else 0

def load_zozo_sales(input_dir, encoding, brand_config: BrandConfig):
    print("  ZOZO売上 (卸売上明細) を集計中...")
    src = find_single_csv(input_dir, ["卸売上明細"], exclude_keywords=["卸売上明細 (1)", "卸売上明細(1)"])
    df = pd.read_csv(src, encoding=encoding, header=0, dtype=str)
    if df.empty:
        return {}

    # ターゲットブランドのみ抽出（Brand列 → ブランド名列 の順で参照）
    mask = df.apply(lambda row: is_target_brand(brand_config, row["商品コード"], get_brand_name_safe(row)), axis=1)
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
    if df.empty:
        return {}

    # ターゲットブランドのみ抽出（Brand列 → ブランド名列 の順で参照）
    if "商品コード" in df.columns:
        mask = df.apply(lambda row: is_target_brand(brand_config, row["商品コード"], get_brand_name_safe(row)), axis=1)
    else:
        mask = df.apply(lambda row: is_target_brand(brand_config, None, get_brand_name_safe(row)), axis=1)
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
    if df.empty:
        return {}

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
        if df.empty:
            return {}
        
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
            df[key_col] = df[key_col].astype(str).str.strip()
            df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
            df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0).astype(int)
            df[val_col] = df[val_col].clip(lower=0)
            
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
        df = pd.read_excel(path, sheet_name=0, header=1, dtype=str)
        if df.empty:
            return {}, {}
        
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
            df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
            
            z_dict = {}
            o_dict = {}
            if zozo_col is not None:
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
        if df.empty:
            return {}
        
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
            
            brand_col = None
            for col in df.columns:
                if str(col).strip() in ["Brand", "BrandName"]:
                    brand_col = col
                    break
            
            if brand_col is not None:
                df = df[df.apply(lambda row: is_target_brand(brand_config, row[key_col], row[brand_col]), axis=1)]
            else:
                df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
            
            def clean_tokka_val(v):
                if pd.isna(v):
                    return ""
                v_str = str(v).strip()
                if v_str == "0" or v_str == "0.0" or v_str == "":
                    return ""
                return v_str

            df[val_col] = df[val_col].apply(clean_tokka_val)
            df = df[df[val_col] != ""]
            
            return dict(zip(df[key_col], df[val_col]))
    except Exception as e:
        print(f"  [警告] 特価在庫の読み込みに失敗しました: {e}")
    return {}
