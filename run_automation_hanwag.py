# -*- coding: utf-8 -*-
"""
出荷明細作成自動化スクリプト [HWG用] (リファクタリング版)
===================================================
実行方法: 実行_hanwag.bat をダブルクリック、または python run_automation_hanwag.py
"""

import os
import datetime
from automation_core import BRAND_HWG, run_pipeline

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
today_str = datetime.datetime.now().strftime("%Y%m%d")
OUTPUT_FILE_NAME = f"{today_str} RETAIL HWG.xlsx"
INPUT_DIR = os.path.join(BASE_DIR, "input")
TEMPLATE_PATH = os.path.join(BASE_DIR, "★RETAIL_テンプレート.xlsx")
ERROR_LOG_PATH = os.path.join(BASE_DIR, "error.log")

def main():
    run_pipeline(
        brand_config=BRAND_HWG,
        output_file_name=OUTPUT_FILE_NAME,
        input_dir=INPUT_DIR,
        base_dir=BASE_DIR,
        template_path=TEMPLATE_PATH,
        error_log_path=ERROR_LOG_PATH
    )

if __name__ == "__main__":
    main()
