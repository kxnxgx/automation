"""
compare_frv.py
完全版（ルートフォルダ）と自動化出力（outputフォルダ）の
20260721 RETAIL FRV.xlsx を全項目比較してレポートを出力するスクリプト。

比較項目:
  1. シート構成（シート名・枚数）
  2. 列構成（列数・ヘッダー名・順番）
  3. 行構成（行数・天井行・データ開始行）
  4. 数値（セルの値）
  5. 数式（セルの数式文字列）
  6. セル書式（フォント色・背景色・太字・イタリック・エンファシス）
  7. 条件付き書式（設定有無・適用範囲・ルールタイプ）
"""

import sys
import os
import json
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────────
# パス設定
# ──────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
MANUAL_FILE = BASE_DIR / "20260721 RETAIL FRV.xlsx"       # 完全版（正解）

# output フォルダ内の最新の RETAIL FRV ファイルを自動検出
output_dir = BASE_DIR / "output"
auto_files = list(output_dir.glob("*RETAIL FRV.xlsm")) + list(output_dir.glob("*RETAIL FRV.xlsx"))
if auto_files:
    AUTO_FILE = sorted(auto_files, key=lambda x: x.stat().st_mtime, reverse=True)[0]
else:
    AUTO_FILE = output_dir / "20260721 RETAIL FRV.xlsm"


# ──────────────────────────────────────────────
# ヘルパー
# ──────────────────────────────────────────────

def rgb_to_hex(color_obj):
    """openpyxl の Color オブジェクトを #RRGGBB 文字列に変換"""
    try:
        if color_obj is None:
            return None
        t = color_obj.type
        if t == "rgb":
            return f"#{color_obj.rgb}"
        if t == "theme":
            return f"theme:{color_obj.theme},tint:{color_obj.tint:.4f}"
        if t == "indexed":
            return f"indexed:{color_obj.indexed}"
    except Exception:
        pass
    return None


def cell_format_summary(cell):
    """セル書式をdict形式で返す"""
    font = cell.font
    fill = cell.fill
    return {
        "bold":       font.bold,
        "italic":     font.italic,
        "font_size":  font.size,
        "font_color": rgb_to_hex(font.color),
        "fill_type":  fill.fill_type,
        "fgColor":    rgb_to_hex(fill.fgColor),
        "bgColor":    rgb_to_hex(fill.bgColor),
    }


def cf_summary(ws):
    """条件付き書式の概要リストを返す（openpyxl新旧両対応）"""
    result = []
    cf_obj = ws.conditional_formatting
    try:
        # openpyxl >= 3.1 では ConditionalFormattingList を直接イテレートする
        # cf_obj は {sqref: [rule, ...]} のdict-likeか、またはイテレート可能
        if hasattr(cf_obj, 'cf_rules'):
            # 旧API
            items = cf_obj.cf_rules.items()
        else:
            # 新API: ws.conditional_formatting[sqref] でルールを取得
            items = [(str(sqref), cf_obj[sqref]) for sqref in cf_obj]
    except Exception:
        return result

    for sqref, rules in items:
        for rule in rules:
            formula_list = []
            try:
                formula_list = [str(f) for f in (rule.formula or [])]
            except Exception:
                pass
            result.append({
                "range": str(sqref),
                "type":  getattr(rule, 'type', 'unknown'),
                "priority": getattr(rule, 'priority', None),
                "formula": formula_list,
            })
    return result


# ──────────────────────────────────────────────
# メイン比較処理
# ──────────────────────────────────────────────

def compare_files(manual_path, auto_path):
    print(f"\n{'='*70}")
    print(f"【完全版】{manual_path}")
    print(f"【自動化】{auto_path}")
    print(f"{'='*70}\n")

    # data_only=False で数式も読み込む
    wb_m = load_workbook(manual_path, data_only=False)
    wb_a = load_workbook(auto_path,   data_only=False)
    # 数値取得用（data_only=True）
    wb_m_val = load_workbook(manual_path, data_only=True)
    wb_a_val = load_workbook(auto_path,   data_only=True)

    report = []  # [{"sheet":..., "category":..., "detail":...}, ...]
    ok_count   = 0
    warn_count = 0

    def log_ok(sheet, category, detail):
        nonlocal ok_count
        ok_count += 1
        report.append({"status": "OK",   "sheet": sheet, "category": category, "detail": detail})

    def log_diff(sheet, category, detail):
        nonlocal warn_count
        warn_count += 1
        report.append({"status": "DIFF", "sheet": sheet, "category": category, "detail": detail})

    # ──────────────────────────────────────────
    # 1. シート構成
    # ──────────────────────────────────────────
    sheets_m = wb_m.sheetnames
    sheets_a = wb_a.sheetnames

    print("■ 1. シート構成チェック")
    if sheets_m == sheets_a:
        print(f"  [OK] シート名・順番が一致: {sheets_m}")
        log_ok("all", "シート構成", f"一致: {sheets_m}")
    else:
        only_m = [s for s in sheets_m if s not in sheets_a]
        only_a = [s for s in sheets_a if s not in sheets_m]
        order_diff = sheets_m != sheets_a
        msg = f"完全版にのみ存在: {only_m} / 自動化にのみ存在: {only_a} / 順番差異: {order_diff}"
        print(f"  [DIFF] {msg}")
        log_diff("all", "シート構成", msg)

    # 共通シートのみ詳細比較
    common_sheets = [s for s in sheets_m if s in sheets_a]

    for sheet_name in common_sheets:
        ws_m = wb_m[sheet_name]
        ws_a = wb_a[sheet_name]
        ws_m_val = wb_m_val[sheet_name]
        ws_a_val = wb_a_val[sheet_name]

        print(f"\n{'─'*60}")
        print(f"■ シート: 【{sheet_name}】")
        print(f"{'─'*60}")

        # ──────────────────────────────────────
        # 2. 行列構成
        # ──────────────────────────────────────
        max_row_m = ws_m.max_row
        max_row_a = ws_a.max_row
        max_col_m = ws_m.max_column
        max_col_a = ws_a.max_column

        print(f"\n  ◆ 2. 行列構成")
        if max_row_m == max_row_a:
            print(f"    [OK] 最大行数: {max_row_m}")
            log_ok(sheet_name, "最大行数", f"{max_row_m}")
        else:
            print(f"    [DIFF] 最大行数: 完全版={max_row_m} / 自動化={max_row_a}")
            log_diff(sheet_name, "最大行数", f"完全版={max_row_m} / 自動化={max_row_a}")

        if max_col_m == max_col_a:
            print(f"    [OK] 最大列数: {max_col_m} ({get_column_letter(max_col_m)}列)")
            log_ok(sheet_name, "最大列数", f"{max_col_m}")
        else:
            print(f"    [DIFF] 最大列数: 完全版={max_col_m}({get_column_letter(max_col_m)}) / 自動化={max_col_a}({get_column_letter(max_col_a)})")
            log_diff(sheet_name, "最大列数", f"完全版={max_col_m}({get_column_letter(max_col_m)}) / 自動化={max_col_a}({get_column_letter(max_col_a)})")

        # ──────────────────────────────────────
        # 3. ヘッダー行の列名比較（1行目）
        # ──────────────────────────────────────
        print(f"\n  ◆ 3. ヘッダー行（1行目）列構成")
        headers_m = [ws_m.cell(1, c).value for c in range(1, max_col_m + 1)]
        headers_a = [ws_a.cell(1, c).value for c in range(1, max_col_a + 1)]

        if headers_m == headers_a:
            print(f"    [OK] ヘッダー完全一致 ({len(headers_m)}列)")
            log_ok(sheet_name, "ヘッダー列構成", f"一致 ({len(headers_m)}列)")
        else:
            max_h = max(len(headers_m), len(headers_a))
            diffs = []
            for i in range(max_h):
                hm = headers_m[i] if i < len(headers_m) else "（存在しない）"
                ha = headers_a[i] if i < len(headers_a) else "（存在しない）"
                if hm != ha:
                    diffs.append(f"  列{i+1}({get_column_letter(i+1)}): 完全版='{hm}' / 自動化='{ha}'")
            print(f"    [DIFF] ヘッダーに差異あり ({len(diffs)}箇所)")
            for d in diffs:
                print(f"    {d}")
            log_diff(sheet_name, "ヘッダー列構成", "\n".join(diffs))

        # ──────────────────────────────────────
        # 4. 数値比較（全セル）
        # ──────────────────────────────────────
        print(f"\n  ◆ 4. 数値比較（全セル）")
        val_diffs = []
        max_r = max(max_row_m, max_row_a)
        max_c = max(max_col_m, max_col_a)

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                vm = ws_m_val.cell(r, c).value
                va = ws_a_val.cell(r, c).value

                # 数値の微差を許容（float の丸め誤差対策）
                if isinstance(vm, float) and isinstance(va, float):
                    if abs(vm - va) > 0.001:
                        val_diffs.append(f"  R{r}C{c}({get_column_letter(c)}{r}): 完全版={vm} / 自動化={va}")
                elif vm != va:
                    val_diffs.append(f"  R{r}C{c}({get_column_letter(c)}{r}): 完全版={repr(vm)} / 自動化={repr(va)}")

        if not val_diffs:
            print(f"    [OK] 全セルの数値が一致")
            log_ok(sheet_name, "数値", "全セル一致")
        else:
            # 上限表示
            show = val_diffs[:50]
            print(f"    [DIFF] 数値差異: {len(val_diffs)}件（上位50件を表示）")
            for d in show:
                print(f"    {d}")
            if len(val_diffs) > 50:
                print(f"    ... 他 {len(val_diffs)-50} 件")
            log_diff(sheet_name, "数値", f"{len(val_diffs)}件の差異\n" + "\n".join(show[:20]))

        # ──────────────────────────────────────
        # 5. 数式比較（全セル）
        # ──────────────────────────────────────
        print(f"\n  ◆ 5. 数式比較（全セル）")
        formula_diffs = []

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                fm = ws_m.cell(r, c).value
                fa = ws_a.cell(r, c).value

                fm_is_formula = isinstance(fm, str) and fm.startswith("=")
                fa_is_formula = isinstance(fa, str) and fa.startswith("=")

                if fm_is_formula != fa_is_formula:
                    formula_diffs.append(
                        f"  R{r}C{c}({get_column_letter(c)}{r}): "
                        f"完全版={'数式あり' if fm_is_formula else '数式なし'} / "
                        f"自動化={'数式あり' if fa_is_formula else '数式なし'}"
                        f" | 完全版={repr(fm)} / 自動化={repr(fa)}"
                    )
                elif fm_is_formula and fa_is_formula and fm != fa:
                    formula_diffs.append(
                        f"  R{r}C{c}({get_column_letter(c)}{r}): 完全版={fm} / 自動化={fa}"
                    )

        if not formula_diffs:
            print(f"    [OK] 全セルの数式が一致（または両方数式なし）")
            log_ok(sheet_name, "数式", "全セル一致")
        else:
            show = formula_diffs[:30]
            print(f"    [DIFF] 数式差異: {len(formula_diffs)}件（上位30件を表示）")
            for d in show:
                print(f"    {d}")
            if len(formula_diffs) > 30:
                print(f"    ... 他 {len(formula_diffs)-30} 件")
            log_diff(sheet_name, "数式", f"{len(formula_diffs)}件の差異\n" + "\n".join(show[:15]))

        # ──────────────────────────────────────
        # 6. セル書式比較（全セル）
        # ──────────────────────────────────────
        print(f"\n  ◆ 6. セル書式比較")
        fmt_diffs = []

        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                cm = ws_m.cell(r, c)
                ca = ws_a.cell(r, c)
                fm_fmt = cell_format_summary(cm)
                fa_fmt = cell_format_summary(ca)
                if fm_fmt != fa_fmt:
                    diffs_detail = {k: (fm_fmt[k], fa_fmt[k]) for k in fm_fmt if fm_fmt[k] != fa_fmt[k]}
                    fmt_diffs.append(
                        f"  R{r}C{c}({get_column_letter(c)}{r}): {diffs_detail}"
                    )

        if not fmt_diffs:
            print(f"    [OK] 全セルの書式が一致")
            log_ok(sheet_name, "セル書式", "全セル一致")
        else:
            show = fmt_diffs[:30]
            print(f"    [DIFF] 書式差異: {len(fmt_diffs)}件（上位30件を表示）")
            for d in show:
                print(f"    {d}")
            if len(fmt_diffs) > 30:
                print(f"    ... 他 {len(fmt_diffs)-30} 件")
            log_diff(sheet_name, "セル書式", f"{len(fmt_diffs)}件の差異\n" + "\n".join(show[:15]))

        # ──────────────────────────────────────
        # 7. 条件付き書式
        # ──────────────────────────────────────
        print(f"\n  ◆ 7. 条件付き書式")
        cf_m = cf_summary(ws_m)
        cf_a = cf_summary(ws_a)

        cf_m_set = set(json.dumps(r, sort_keys=True) for r in cf_m)
        cf_a_set = set(json.dumps(r, sort_keys=True) for r in cf_a)

        only_m_cf = [json.loads(r) for r in cf_m_set - cf_a_set]
        only_a_cf = [json.loads(r) for r in cf_a_set - cf_m_set]

        if not only_m_cf and not only_a_cf:
            print(f"    [OK] 条件付き書式が完全一致 ({len(cf_m)}件)")
            log_ok(sheet_name, "条件付き書式", f"一致 ({len(cf_m)}件)")
        else:
            msg_parts = []
            if only_m_cf:
                msg_parts.append(f"完全版のみに存在({len(only_m_cf)}件):")
                for r in only_m_cf[:10]:
                    msg_parts.append(f"    range={r['range']}, type={r['type']}, formula={r.get('formula')}")
            if only_a_cf:
                msg_parts.append(f"自動化のみに存在({len(only_a_cf)}件):")
                for r in only_a_cf[:10]:
                    msg_parts.append(f"    range={r['range']}, type={r['type']}, formula={r.get('formula')}")
            detail = "\n".join(msg_parts)
            print(f"    [DIFF] 条件付き書式に差異あり")
            print(f"    {detail}")
            log_diff(sheet_name, "条件付き書式", detail)

    # ──────────────────────────────────────────
    # 最終サマリー
    # ──────────────────────────────────────────
    print(f"\n{'='*70}")
    print(f"■ 比較完了サマリー")
    print(f"  OK件数  : {ok_count}")
    print(f"  差異件数: {warn_count}")
    print(f"{'='*70}\n")

    # 差異一覧のみ再出力
    if warn_count > 0:
        print("■ 差異一覧（DIFF のみ）")
        for item in report:
            if item["status"] == "DIFF":
                print(f"\n  [{item['sheet']}] {item['category']}")
                for line in item["detail"].split("\n"):
                    print(f"    {line}")
    else:
        print("✅ すべての項目が完全一致しています！")

    return report


if __name__ == "__main__":
    if not MANUAL_FILE.exists():
        print(f"[ERROR] 完全版ファイルが見つかりません: {MANUAL_FILE}")
        sys.exit(1)
    if not AUTO_FILE.exists():
        print(f"[ERROR] 自動化出力ファイルが見つかりません: {AUTO_FILE}")
        sys.exit(1)

    compare_files(MANUAL_FILE, AUTO_FILE)
