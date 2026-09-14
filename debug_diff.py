# -*- coding: utf-8 -*-
"""
差異デバッグスクリプト: debug_diff.py (リファクタリング版)
"""
import os
import sys
import datetime
from automation_core import BRAND_FRV, BRAND_TEN, BRAND_HWG, verify_pipeline

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
today_str = datetime.datetime.now().strftime("%Y%m%d")

def main():
    # 引数からブランド名を取得（指定がない場合はFRVがデフォルト）
    brand_name = "FRV"
    if len(sys.argv) > 1:
        brand_name = sys.argv[1].upper()

    if brand_name == "TEN":
        brand_config = BRAND_TEN
        output_file = f"{today_str} RETAIL TEN.xlsm"
        input_folder = "input"
    elif brand_name == "HWG" or brand_name == "HANWAG":
        brand_config = BRAND_HWG
        output_file = f"{today_str} RETAIL HWG.xlsm"
        input_folder = "input"
    else:
        brand_config = BRAND_FRV
        output_file = f"{today_str} RETAIL FRV.xlsm"
        input_folder = "input"

    result_path = os.path.join(BASE_DIR, "output", output_file)
    input_dir = os.path.join(BASE_DIR, input_folder)

    print(f"=== {brand_config.name} 差異デバッグ実行中 ===")
    success = verify_pipeline(brand_config, input_dir, result_path)
    if success:
        print(f"\n[{brand_config.name}] 差異はありません。すべて一致しています！")
    else:
        print(f"\n[{brand_config.name}] 不一致またはエラーが検出されました。上記の詳細を確認してください。")

if __name__ == "__main__":
    main()
