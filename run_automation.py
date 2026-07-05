# -*- coding: utf-8 -*-
"""
出荷明細作成自動化スクリプト (ピボット廃止・高速化版)
===============================
実行方法: 実行.bat をダブルクリック、または python run_automation.py

input/ フォルダに以下のCSVを置いてから実行してください:
  - 出荷予定振分*.csv       → RETAIL.csv にリネーム
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
import win32com.client

# ============================================================
# 設定
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
XLSM_PATH = os.path.join(BASE_DIR, "★RETAIL_CSV2計算式.xlsm")
RETAIL_CSV_PATH = os.path.join(BASE_DIR, "RETAIL.csv")
FORMULA_XLSX_PATH = os.path.join(BASE_DIR, "★RETAIL_CSV2計算式.xlsx")
ERROR_LOG_PATH = os.path.join(BASE_DIR, "error.log")
ENCODING = "cp932"

# シート名
SHEET_STORE_PASTE = "営業日付別売上分析 貼り付け"
SHEET_PIVOT       = "pivot"
SHEET_ZOZOOIOI    = "ZOZOOIOI"
SHEET_EC_ORDER    = "EC_ORDER"
SHEET_EC_PIVOT    = "EC_PIVOT"
SHEET_KEISAN      = "計算式2026"

STORE_COLUMNS = ['名古屋', 'TOKYO', 'ルクア大阪', 'ヒュッテ', '京王新宿', '大丸心斎橋', '6142', '玉川高島屋', 'NARITA']

# ============================================================
# ユーティリティ
# ============================================================
def show_error(msg):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("エラー", msg)
    root.destroy()
    sys.exit(1)

def show_info(msg):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("完了", msg)
    root.destroy()

def find_single_csv(pattern_keywords, exclude_keywords=None):
    all_csvs = glob.glob(os.path.join(INPUT_DIR, "*.csv"))
    matched = []
    for path in all_csvs:
        name = os.path.basename(path)
        if all(kw in name for kw in pattern_keywords):
            if exclude_keywords and any(ek in name for ek in exclude_keywords):
                continue
            matched.append(path)

    if len(matched) == 0:
        show_error(f"input/ フォルダに該当するCSVが見つかりません: {pattern_keywords}")
    if len(matched) > 1:
        names = "\n".join(os.path.basename(p) for p in matched)
        show_error(f"input/ フォルダに同種のCSVが複数あります: {names}")
    return matched[0]

def delete_all_pivots(ws):
    try:
        pts = ws.PivotTables()
        for pt in pts:
            pt.TableRange2.Clear()
    except Exception:
        pass

def write_df_to_sheet(ws, df, start_row=1, start_col=1, clear_rows=3000):
    n_rows = len(df)
    n_cols = len(df.columns)
    
    delete_all_pivots(ws)

    if clear_rows > 0:
        clear_end_row = start_row + max(n_rows, clear_rows) - 1
        clear_end_col = start_col + max(n_cols, 30) - 1
        try:
            ws.Range(ws.Cells(start_row, start_col), ws.Cells(clear_end_row, clear_end_col)).ClearContents()
        except Exception:
            pass

    if n_rows == 0 or n_cols == 0:
        return

    data_matrix = []
    for row in df.values.tolist():
        new_row = []
        for val in row:
            if pd.isna(val) if isinstance(val, float) else val is None:
                new_row.append("")
            else:
                new_row.append(val)
        data_matrix.append(new_row)

    target_range = ws.Range(ws.Cells(start_row, start_col), ws.Cells(start_row + n_rows - 1, start_col + n_cols - 1))
    target_range.Value = data_matrix

# ============================================================
# 手順1: 出荷予定振分.csv → RETAIL.csv
# ============================================================
def step1_open_retail_via_excel(xl_app):
    print("[手順1] 出荷予定振分.csv → RETAIL.csv")
    src = find_single_csv(["出荷予定振分"])
    dst_abs = os.path.abspath(RETAIL_CSV_PATH)
    
    if os.path.exists(dst_abs):
        try:
            os.remove(dst_abs)
        except Exception:
            pass

    wb_src = xl_app.Workbooks.Open(os.path.abspath(src))
    
    try:
        wb_src.SaveAs(dst_abs, FileFormat=6)
    except Exception as e:
        show_error(f"RETAIL.csv への保存失敗: {e}")

    try:
        if wb_src.Sheets(1).Name != "RETAIL":
            wb_src.Sheets(1).Name = "RETAIL"
    except Exception:
        pass
    return wb_src

# ============================================================
# 手順2: 営業日付別売上分析(旧) → 店舗売上 (raw & pivot)
# ============================================================
def step2_store_sales(wb):
    print("[手順2] 営業日付別売上分析(旧) → 店舗売上 (集計処理)")
    src = find_single_csv(["営業日付別売上分析"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    brand_mask = df["表記部門名1"].isin(["FRV", "FJALLRAVEN"])
    df_filtered = df[brand_mask].copy()
    
    raw_cols = ["店舗名称", "3rd Item No.", "表記部門名1", "StyleNo", "商品名", "ColorNo", "ColorName", "Size", "数量"]
    existing_raw_cols = [c for c in raw_cols if c in df_filtered.columns]
    df_raw = df_filtered[existing_raw_cols].reset_index(drop=True)

    ws_paste = wb.Sheets(SHEET_STORE_PASTE)
    write_df_to_sheet(ws_paste, df_raw, start_row=1, start_col=1)
    
    df_filtered["数量"] = pd.to_numeric(df_filtered["数量"], errors="coerce").fillna(0)
    pivot_df = df_filtered.pivot_table(index="3rd Item No.", columns="店舗名称", values="数量", aggfunc="sum").reset_index()

    result_df = pd.DataFrame()
    result_df["商品コード"] = pivot_df["3rd Item No."] if "3rd Item No." in pivot_df.columns else []
    for store in STORE_COLUMNS:
        if store in pivot_df.columns:
            result_df[store] = pivot_df[store]
        else:
            result_df[store] = 0
            
    result_df = result_df.fillna(0)
    ws_pivot = wb.Sheets(SHEET_PIVOT)
    write_df_to_sheet(ws_pivot, result_df, start_row=5, start_col=1, clear_rows=3000)

# ============================================================
# 手順3: 卸売上明細.csv (ZOZO) → ZOZOOIOI
# ============================================================
def step3_zozo(wb):
    print("[手順3] 卸売上明細.csv (ZOZO) → 集計処理")
    src = find_single_csv(["卸売上明細"], exclude_keywords=["卸売上明細 (1)", "卸売上明細(1)"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    brand_mask = df["ブランド名"].isin(["FRV", "FJALLRAVEN"])
    df_filtered = df[brand_mask].copy()

    raw_cols = ["商品コード", "ブランド名", "StyleNo", "商品名", "ColorNo", "ColorName", "Size", "標準単価", "販売数量"]
    existing_raw_cols = [c for c in raw_cols if c in df_filtered.columns]
    df_raw = df_filtered[existing_raw_cols].reset_index(drop=True)
    
    ws = wb.Sheets(SHEET_ZOZOOIOI)
    write_df_to_sheet(ws, df_raw, start_row=1, start_col=2)  # B1

    df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
    agg_df = df_filtered.groupby("商品コード")["販売数量"].sum().reset_index()

    write_df_to_sheet(ws, agg_df, start_row=2, start_col=12, clear_rows=3000)

# ============================================================
# 手順4: 卸売上明細 (1).csv (OIOI) → ZOZOOIOI
# ============================================================
def step4_oioi(wb):
    print("[手順4] 卸売上明細 (1).csv (OIOI) → 集計処理")
    src = find_single_csv(["卸売上明細 (1)"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    brand_mask = df.get("ブランド名", pd.Series(dtype=str)).isin(["FRV", "FJALLRAVEN"])
    df_filtered = df[brand_mask].copy()

    raw_cols = ["商品コード", "ブランド名", "StyleNo", "商品名", "ColorNo", "ColorName", "Size", "標準単価", "販売数量"]
    existing_raw_cols = [c for c in raw_cols if c in df_filtered.columns]
    df_raw = df_filtered[existing_raw_cols].reset_index(drop=True)
    
    ws = wb.Sheets(SHEET_ZOZOOIOI)
    write_df_to_sheet(ws, df_raw, start_row=1, start_col=16)  # P1

    if "販売数量" in df_filtered.columns and "商品コード" in df_filtered.columns:
        df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
        agg_df = df_filtered.groupby("商品コード")["販売数量"].sum().reset_index()
    else:
        agg_df = pd.DataFrame(columns=["商品コード", "販売数量"])

    write_df_to_sheet(ws, agg_df, start_row=2, start_col=26, clear_rows=3000)

# ============================================================
# 手順5: order_YYYYMMDD.csv (EC) → EC_ORDER / EC_PIVOT
# ============================================================
def step5_ec_order(wb):
    print("[手順5] order_*.csv (EC) → 集計処理")
    src = find_single_csv(["order_"])
    df = pd.read_csv(src, encoding=ENCODING, header=0, dtype=str)

    df = df[~df["決済方法(ステータス)"].str.contains("キャンセル", na=False)]
    df = df[~df["注文者"].str.contains("店舗客注", na=False)]
    df = df.reset_index(drop=True)

    ws_raw = wb.Sheets(SHEET_EC_ORDER)
    write_df_to_sheet(ws_raw, df, start_row=1, start_col=1)

    df["個数"] = pd.to_numeric(df["個数"], errors="coerce").fillna(0)
    agg_df = df.groupby("商品コード")["個数"].sum().reset_index()

    ws_agg = wb.Sheets(SHEET_EC_PIVOT)
    write_df_to_sheet(ws_agg, agg_df, start_row=3, start_col=1, clear_rows=3000)

# ============================================================
# 手順6: PythonによるVBA代替処理 (UI・フォーマット作成と数式コピー)
# ============================================================
def step6_python_vba_replacement(xl_app, wb, wb_retail):
    print("[手順6] PythonによるVBA代替処理 (orderシートの作成と数式コピー)")
    
    try:
        wb_retail.Sheets("order").Delete()
    except Exception:
        pass

    ws_retail = wb_retail.Sheets("RETAIL")
    ws_retail.Range("AW1:BY10000").Clear()
    
    ws_retail.Copy(After=wb_retail.Sheets(1)) # Sheets(1)の後ろにコピー (VBAと同じ挙動)
    ws_order = wb_retail.Sheets(2)
    ws_order.Name = "order"

    ws_order.Columns("A:C").Delete()
    ws_order.Columns("A:A").NumberFormatLocal = "0_);[赤](0)"
    ws_order.Columns("A:A").HorizontalAlignment = -4131 
    ws_order.Columns("A:A").VerticalAlignment = -4108 
    ws_order.Columns("B:B").EntireColumn.AutoFit()
    ws_order.Columns("C:C").Delete()
    ws_order.Range("D4").Value = "RETAIL確保"
    ws_order.Rows(3).Delete()
    
    rng = ws_order.Range("C2:Q2")
    rng.ClearContents()
    rng.Merge()
    rng.Value = "確保"
    rng.HorizontalAlignment = -4108
    rng.Interior.Color = 14277081 
    ws_order.Range("C3:Q3").Interior.Color = 14277081
    
    ws_order.Columns("R:X").Delete()
    rng = ws_order.Range("R2:AD2")
    rng.ClearContents()
    rng.Merge()
    rng.Value = "在庫"
    rng.HorizontalAlignment = -4108
    rng.Interior.Color = 11854022 
    ws_order.Range("R3:AD3").Interior.Color = 11854022
    
    ws_order.Columns("AE:AK").Delete()
    rng = ws_order.Range("AE2:AQ2")
    rng.ClearContents()
    rng.Merge()
    rng.Value = "ORDER"
    rng.HorizontalAlignment = -4108
    rng.Interior.Color = 65535 
    ws_order.Range("AE3:AQ3").Interior.Color = 65535
    
    ws_order.Range("AR:BA").ClearContents()
    
    ws_order.Range("C3:AQ3").Orientation = -4166 
    ws_order.Columns("C:AQ").ColumnWidth = 5
    
    ws_order.Range("A2:B2").Cut(Destination=ws_order.Range("A3:B3"))
    
    lastRow = ws_order.Cells(ws_order.Rows.Count, "B").End(-4162).Row 
    if lastRow < 4: lastRow = 4
    
    borders_range = ws_order.Range(f"R2:AD{lastRow}")
    borders_range.Borders(7).LineStyle = 1 
    borders_range.Borders(10).LineStyle = 1 
    
    ws_order.Rows(1).ClearContents()
    ws_order.Range("C1").Formula = f"=SUBTOTAL(9,C4:C{lastRow})"
    ws_order.Range("C1").Copy(Destination=ws_order.Range("D1:AQ1"))
    ws_order.Columns("E:Q").Hidden = True
    
    # 計算式2026 シートの式を文字列として取得し、明示的に .xlsm へのリンクに書き換える
    # (Excelのコピー＆ペーストを使うと、内部キャッシュで古い .xlsx にリンクしてしまうバグを防ぐため)
    ws_formula = wb.Sheets("計算式2026")
    formulas_3 = [ws_formula.Cells(3, c).Formula if ws_formula.Cells(3, c).HasFormula else ws_formula.Cells(3, c).Value for c in range(31, 60)]
    formulas_4 = [ws_formula.Cells(4, c).Formula if ws_formula.Cells(4, c).HasFormula else ws_formula.Cells(4, c).Value for c in range(31, 60)]
    
    xlsm_name = "[★RETAIL_CSV2計算式.xlsm]"
    
    def force_link(val):
        if not val or not str(val).startswith("="):
            return val
        f = str(val)
        for sheet in ["ZOZOOIOI", "pivot", "計算式2026", "EC_PIVOT"]:
            f = f.replace(f"{sheet}!", f"'{xlsm_name}{sheet}'!")
        return f

    for idx in range(29): # AE~BG = 29 columns
        col = 31 + idx
        ws_order.Cells(3, col).Value = force_link(formulas_3[idx])
        ws_order.Cells(4, col).Formula = force_link(formulas_4[idx])
        
    # 4行目の式を最終行まで一気にオートフィル
    ws_order.Range("AE4:BG4").AutoFill(Destination=ws_order.Range(f"AE4:BG{lastRow}"), Type=0)
    
    xl_app.CutCopyMode = False
        
    ws_order.Columns("AR:BG").ColumnWidth = 4.43
    ws_order.Columns("AR:BF").Hidden = True
    
    # 欠落していた「過不足がマイナスの行はA列を赤く」する条件付き書式を追加
    ws_order.Columns("A:A").FormatConditions.Delete()
    fc = ws_order.Columns("A:A").FormatConditions.Add(Type=2, Formula1="=$BG1<0") # 2 = xlExpression
    fc.SetFirstPriority()
    fc.Interior.Color = 255 # 赤
    fc.StopIfTrue = False
    
    ws_retail.Columns("D:D").NumberFormatLocal = "0_);[赤](0)"
    ws_retail.Range("AW5").FormulaR1C1 = "=order!R[-1]C[-18]"
    ws_retail.Range("AW5").AutoFill(Destination=ws_retail.Range("AW5:BY5"), Type=0) 
    
    lastRowRetail = ws_retail.Cells(ws_retail.Rows.Count, "B").End(-4162).Row
    if lastRowRetail < 6: lastRowRetail = 6
    
    ws_retail.Range("AW5:BY5").Copy()
    ws_retail.Range(f"AW6:BY{lastRowRetail}").PasteSpecial(Paste=-4104)
    xl_app.CutCopyMode = False
    
    ws_retail.Columns("BJ:BP").Delete()
    print("  PythonによるVBA代替処理完了")

# ============================================================
# メイン処理
# ============================================================
def main():
    print("=" * 50)
    print("出荷明細作成自動化スクリプト 開始 (堅牢・完全Python版)")
    print("=" * 50)

    xl_app = None
    wb = None
    wb_retail = None
    success = False

    try:
        print("\nExcel をバックグラウンドで起動中...")
        xl_app = win32com.client.DispatchEx("Excel.Application")
        xl_app.Visible = True  
        xl_app.DisplayAlerts = False

        wb_retail = step1_open_retail_via_excel(xl_app)
        wb = xl_app.Workbooks.Open(XLSM_PATH)

        step2_store_sales(wb)
        step3_zozo(wb)
        step4_oioi(wb)
        step5_ec_order(wb)
        
        step6_python_vba_replacement(xl_app, wb, wb_retail)

        # テンプレートは上書き保存して閉じる
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        
        # RETAIL.csv は書式（色や数式）を保持するために .xlsx として別名保存
        retail_xlsx_path = os.path.join(BASE_DIR, "RETAIL_完成版.xlsx")
        if os.path.exists(retail_xlsx_path):
            try: os.remove(retail_xlsx_path)
            except Exception: pass
            
        wb_retail.SaveAs(retail_xlsx_path, FileFormat=51) # 51 = xlOpenXMLWorkbook (.xlsx)
        print(f"\n処理が正常に完了しました！ 色付けや数式を保持するため、{os.path.basename(retail_xlsx_path)} として保存しました。")
        success = True
        
    except SystemExit:
        raise
    except Exception as e:
        with open(ERROR_LOG_PATH, "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        print(f"予期しないエラーが発生しました。\n詳細は error.log を確認してください。\n\n{e}")
    finally:
        print("\nクリーンアップ処理を実行中...")
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
        if wb_retail is not None:
            try: wb_retail.Close(SaveChanges=False)
            except Exception: pass
        if xl_app is not None:
            try:
                xl_app.DisplayAlerts = True
                xl_app.Quit()
            except Exception: pass

if __name__ == "__main__":
    main()
