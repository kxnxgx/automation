"""
データ検証スクリプト: RETAIL_完成版.xlsx vs 元CSV の合計突き合わせ
"""
import pandas as pd
import openpyxl
import glob, os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
result_path = os.path.join(BASE_DIR, "RETAIL_完成版.xlsx")
input_dir = os.path.join(BASE_DIR, "input")

def find_csv(keywords, exclude=None):
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

# ────────────────────────────────────────
# tennenブランドの除外判定
# ────────────────────────────────────────
def is_tennen(code, brand_name=None):
    if brand_name and str(brand_name).upper().strip() == "TENNEN":
        return True
    if not code:
        return False
    c = str(code).upper().strip()
    return c.startswith("TNT") or c.startswith("TNP") or c.startswith("TNF")

# ────────────────────────────────────────
# 1. 元CSVから商品ごとの売上データを集計
# ────────────────────────────────────────
# ZOZO
df_zozo = pd.read_csv(find_csv(["卸売上明細"], exclude=["卸売上明細 (1)", "卸売上明細(1)"]), encoding="cp932", dtype=str)
df_zozo = df_zozo[df_zozo["ブランド名"].isin(["FRV", "FJALLRAVEN"])].copy()
df_zozo = df_zozo[~df_zozo["商品コード"].apply(is_tennen)]
df_zozo["販売数量"] = pd.to_numeric(df_zozo["販売数量"], errors="coerce").fillna(0)
zozo_sales = df_zozo.groupby("商品コード")["販売数量"].sum().to_dict()

# OIOI
df_oioi = pd.read_csv(find_csv(["卸売上明細 (1)"]), encoding="cp932", dtype=str)
df_oioi = df_oioi[df_oioi["ブランド名"].isin(["FRV", "FJALLRAVEN"])].copy()
if "商品コード" in df_oioi.columns:
    df_oioi = df_oioi[~df_oioi["商品コード"].apply(is_tennen)]
df_oioi["販売数量"] = pd.to_numeric(df_oioi["販売数量"], errors="coerce").fillna(0)
oioi_sales = df_oioi.groupby("商品コード")["販売数量"].sum().to_dict() if "商品コード" in df_oioi.columns else {}

# 店舗
df_store = pd.read_csv(find_csv(["営業日付別売上分析"]), encoding="cp932", dtype=str)
df_store = df_store[df_store["表記部門名1"].isin(["FRV", "FJALLRAVEN"])].copy()
df_store = df_store[~df_store["3rd Item No."].apply(is_tennen)]
df_store["数量"] = pd.to_numeric(df_store["数量"], errors="coerce").fillna(0)
store_pivot = df_store.pivot_table(index="3rd Item No.", columns="店舗名称", values="数量", aggfunc="sum").fillna(0)

# EC
df_ec = pd.read_csv(find_csv(["order_"]), encoding="cp932", dtype=str)
df_ec = df_ec[~df_ec["決済方法(ステータス)"].str.contains("キャンセル", na=False)]
df_ec = df_ec[~df_ec["注文者"].str.contains("店舗客注", na=False)]
df_ec["個数"] = pd.to_numeric(df_ec["個数"], errors="coerce").fillna(0)
key_col = "オプション独自コード" if "オプション独自コード" in df_ec.columns else "商品コード"
df_ec = df_ec[~df_ec[key_col].fillna("").str.contains("GIFT", na=False)]
brand_col = "ブランド名" if "ブランド名" in df_ec.columns else None
if brand_col:
    df_ec = df_ec[~df_ec.apply(lambda row: is_tennen(row[key_col], row[brand_col]), axis=1)]
else:
    df_ec = df_ec[~df_ec[key_col].apply(is_tennen)]
ec_sales = df_ec.groupby(key_col)["個数"].sum().to_dict()

# ────────────────────────────────────────
# 2. Excelの在庫データ（C, D, E〜Q）をロードし、売上ゼロ化を適用
# ────────────────────────────────────────
wb = openpyxl.load_workbook(result_path, data_only=True)
ws = wb["order"]

# 共通コード別の売上データを保持
resolved_sales = {
    "ZOZO": {}, "OIOI": {}, "EC": {},
    "名古屋": {}, "TOKYO": {}, "ルクア大阪": {}, "ヒュッテ": {},
    "京王新宿": {}, "大丸心斎橋": {}, "6142": {}, "玉川高島屋": {},
    "NODE": {}, "NARITA": {}
}

stores_list = ["名古屋", "TOKYO", "ルクア大阪", "ヒュッテ", "京王新宿", "大丸心斎橋", "6142", "玉川高島屋", "NODE", "NARITA"]

for r in range(4, ws.max_row + 1):
    code = ws.cell(row=r, column=1).value
    if not code:
        continue
    code = str(code).strip()
    
    # 元々の売上を集計
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
        raw_vals[s] = qty
        
    total_sales = sum(raw_vals.values())
    
    # 在庫情報
    c_val = ws.cell(row=r, column=3).value or 0
    d_val = ws.cell(row=r, column=4).value or 0
    kabu_sum = sum(ws.cell(row=r, column=col).value or 0 for col in range(5, 18))
    
    # 不足による売上ゼロ化
    if c_val == 0 and d_val == 0 and kabu_sum < total_sales:
        for k in resolved_sales:
            resolved_sales[k][code] = 0
    else:
        for k in resolved_sales:
            resolved_sales[k][code] = raw_vals[k]

zozo_true = sum(resolved_sales["ZOZO"].values())
oioi_true = sum(resolved_sales["OIOI"].values())
ec_true = sum(resolved_sales["EC"].values())

store_totals = {s: sum(resolved_sales[s].values()) for s in stores_list}

print("=" * 55)
print("[1] 元CSV の合計（正解値）")
print("=" * 55)
print(f"  ZOZO    (卸売上明細.csv):          {zozo_true:>6}")
print(f"  OIOI    (卸売上明細 (1).csv):      {oioi_true:>6}")
for name, total in store_totals.items():
    print(f"  {name:<12}(営業日付別売上分析.csv): {total:>6}")
print(f"  EC      (order_*.csv ※GIFT除く):  {ec_true:>6}")

# ────────────────────────────────────────
# 2. RETAIL_完成版.xlsx から実際の集計値
# ────────────────────────────────────────
print("\n" + "=" * 55)
print("[2] RETAIL_完成版.xlsx の ORDER欄合計（実際値）")
print("=" * 55)

wb = openpyxl.load_workbook(result_path, data_only=True)
ws = wb["order"]

def col_total(col_num):
    return sum(
        (ws.cell(row=r, column=col_num).value or 0)
        for r in range(4, ws.max_row + 1)
        if ws.cell(row=r, column=2).value is not None
    )

# 列番号: AE(31)=ZOZO, AF(32)=名古屋, AG(33)=OIOI, AH(34)=EC
# AI(35)=TOKYO, AJ(36)=ルクア大阪, AK(37)=ヒュッテ, AL(38)=京王新宿
# AM(39)=6142, AN(40)=玉川高島屋, AO(41)=NODE, AP(42)=大丸心斎橋, AQ(43)=NARITA
actual = {
    "ZOZO":      col_total(31),
    "名古屋":    col_total(32),
    "OIOI":      col_total(33),
    "EC":        col_total(34),
    "TOKYO":     col_total(35),
    "ルクア大阪": col_total(36),
    "ヒュッテ":  col_total(37),
    "京王新宿":  col_total(38),
    "6142":      col_total(39),
    "玉川高島屋": col_total(40),
    "NODE":      col_total(41),
    "大丸心斎橋": col_total(42),
    "NARITA":    col_total(43),
}
for name, val in actual.items():
    print(f"  {name:<12}: {val:>6}")

wb.close()

# ────────────────────────────────────────
# 3. 突き合わせ（差異チェック）
# ────────────────────────────────────────
print("\n" + "=" * 55)
print("[3] 突き合わせ結果（正解値 vs 実際値）")
print("=" * 55)

checks = [
    ("ZOZO",      zozo_true,             actual["ZOZO"]),
    ("OIOI",      oioi_true,             actual["OIOI"]),
    ("EC",        ec_true,               actual["EC"]),
    ("名古屋",    store_totals["名古屋"], actual["名古屋"]),
    ("TOKYO",     store_totals["TOKYO"],  actual["TOKYO"]),
    ("ルクア大阪",store_totals["ルクア大阪"], actual["ルクア大阪"]),
    ("ヒュッテ",  store_totals["ヒュッテ"],  actual["ヒュッテ"]),
    ("京王新宿",  store_totals["京王新宿"],  actual["京王新宿"]),
    ("大丸心斎橋",store_totals["大丸心斎橋"],actual["大丸心斎橋"]),
    ("6142",      store_totals["6142"],   actual["6142"]),
    ("玉川高島屋",store_totals["玉川高島屋"],actual["玉川高島屋"]),
    ("NODE",      store_totals["NODE"],   actual["NODE"]),
    ("NARITA",    store_totals["NARITA"], actual["NARITA"]),
]

all_ok = True
for name, truth, got in checks:
    ok = (truth == got)
    mark = "[OK]" if ok else "[NG] 不一致！"
    print(f"  {name:<12}: 正解={truth:>5}  実際={got:>5}  {mark}")
    if not ok:
        all_ok = False

print("\n" + ("[OK] すべて一致しています！" if all_ok else "[NG] 不一致があります。上記を確認してください。"))
