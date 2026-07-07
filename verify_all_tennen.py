# -*- coding: utf-8 -*-
"""
データ検証スクリプト: {OUTPUT_FILE_NAME} vs 元CSV の合計突き合わせ [TEN用]
"""
import os
import datetime
import sys
from automation_core import BRAND_TEN, verify_pipeline

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
today_str = datetime.datetime.now().strftime("%Y%m%d")
OUTPUT_FILE_NAME = f"{today_str} RETAIL TEN.xlsx"
result_path = os.path.join(BASE_DIR, "output", OUTPUT_FILE_NAME)
input_dir = os.path.join(BASE_DIR, "input")

if __name__ == "__main__":
    print(f"=== TEN データ検証開始 ===")
    success = verify_pipeline(BRAND_TEN, input_dir, result_path)
    if success:
        print("\n[OK] すべて一致しています！")
        sys.exit(0)
    else:
        print("\n[NG] 不一致があります。上記を確認してください。")
        sys.exit(1)
