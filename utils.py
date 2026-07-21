# -*- coding: utf-8 -*-
"""
ユーティリティモジュール (utils.py)
=================================
"""

import os
import sys
import glob
import tkinter as tk
from tkinter import messagebox
import pandas as pd
from config import AutomationError, BrandConfig

# ============================================================
# 判定ユーティリティ
# ============================================================
def is_target_brand(config: BrandConfig, code, brand_name=None) -> bool:
    """ブランド設定に基づき、対象ブランドか厳格に判定する。
    ブランド列が存在する場合（brand_nameが渡された場合）は必ずブランド名のみで判定し、
    コードによるフォールバックは行わない。
    """
    if pd.isna(brand_name) or str(brand_name).strip() == "" or str(brand_name).strip().upper() == "NAN":
        brand_name = None

    # 1. ブランド名が存在する場合 → ブランド名のみで即判定
    if brand_name is not None and str(brand_name).strip():
        b = str(brand_name).upper().strip()
        return b in config.allowed_names

    # 2. ブランド名がない場合のみコードで判定
    if not code or pd.isna(code):
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

def get_brand_name_safe(row):
    """pandasの行からBrandまたはブランド名を安全に取得する（NaN回避）"""
    b = row.get("Brand")
    if pd.isna(b) or str(b).strip() == "" or str(b).strip().upper() == "NAN":
        b = row.get("ブランド名")
    if pd.isna(b) or str(b).strip() == "" or str(b).strip().upper() == "NAN":
        return None
    return b

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
# CSV店舗構造の動的解析
# ============================================================
def parse_csv_store_structure(input_dir, encoding):
    """
    出荷予定振分.csv の店舗構造を動的に解析する。
    返り値: dict で以下のキーを含む:
      - N          : ORDER セクションの店舗数 (= BULK = 在庫 と同数のはず)
      - gap1       : BULK店舗と在庫店舗の間の空白列数
      - gap2       : 在庫店舗とORDER店舗の間の空白列数
      - store_names: ORDERセクションの店舗名リスト (len = N)
      - store_codes: ORDERセクションの店舗コードリスト (len = N)
    """
    src = find_single_csv(input_dir, ["出荷予定振分"])
    df_raw = pd.read_csv(src, encoding=encoding, header=0, dtype=str)

    row_names = list(df_raw.iloc[2].tolist())
    row_codes = list(df_raw.iloc[1].tolist())

    BULK_STORE_START = 8

    N = 0
    for i in range(BULK_STORE_START, len(row_names)):
        s = str(row_names[i]).strip()
        if s == 'nan' or s == '':
            N = i - BULK_STORE_START
            break
    else:
        N = len(row_names) - BULK_STORE_START

    if N <= 0:
        raise AutomationError(
            "店舗構造の解析失敗",
            "出荷予定振分.csv の店舗名行から店舗数を検出できませんでした。"
        )

    gap1 = 0
    for i in range(BULK_STORE_START + N, len(row_names)):
        s = str(row_names[i]).strip()
        if s == 'nan' or s == '':
            gap1 += 1
        else:
            break

    zaiko_start = BULK_STORE_START + N + gap1
    zaiko_N = 0
    for i in range(zaiko_start, len(row_names)):
        s = str(row_names[i]).strip()
        if s == 'nan' or s == '':
            zaiko_N = i - zaiko_start
            break
    else:
        zaiko_N = len(row_names) - zaiko_start

    gap2 = 0
    for i in range(zaiko_start + zaiko_N, len(row_names)):
        s = str(row_names[i]).strip()
        if s == 'nan' or s == '':
            gap2 += 1
        else:
            break

    order_start = zaiko_start + zaiko_N + gap2
    order_N = 0
    for i in range(order_start, len(row_names)):
        s = str(row_names[i]).strip()
        if s == 'nan' or s == '':
            order_N = i - order_start
            break
    else:
        order_N = len(row_names) - order_start

    store_names = [str(row_names[order_start + i]).strip() for i in range(order_N)]
    store_codes = [str(row_codes[order_start + i]).strip() for i in range(order_N)]
    final_N = order_N if order_N > 0 else N

    return {
        'N'          : final_N,
        'gap1'       : gap1,
        'gap2'       : gap2,
        'store_names': store_names,
        'store_codes': store_codes,
    }

# ============================================================
# 店舗名曖昧マッチング関数
# ============================================================
def is_store_match(col_name: str, target_store: str) -> bool:
    """列名 col_name がターゲット店舗 target_store に合致するか判定する。"""
    if pd.isna(col_name) or not col_name:
        return False
        
    from config import STORE_MATCHING_RULES
    col_str = str(col_name).strip()
    col_upper = col_str.upper()
    
    rule = STORE_MATCHING_RULES.get(target_store)
    if rule:
        include_kws = rule.get("include", [])
        exclude_kws = rule.get("exclude", [])
        
        # include 判定 (いずれかのキーワードが含まれているか)
        has_include = False
        for kw in include_kws:
            if kw.upper() in col_upper:
                has_include = True
                break
        if not has_include:
            return False
            
        # exclude 判定 (いずれかのキーワードが含まれていたら除外)
        for kw in exclude_kws:
            if kw.upper() in col_upper:
                return False
                
        return True
    
    # ルールがない場合は単純部分一致 (大文字小文字無視、半角スペース等除去)
    target_clean = target_store.replace(" ", "").replace("　", "").upper()
    col_clean = col_str.replace(" ", "").replace("　", "").upper()
    return target_clean in col_clean

