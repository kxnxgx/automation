import os
from automation_core import run_pipeline, verify_pipeline, BRAND_FRV

if __name__ == "__main__":
    base_dir = r"c:\Users\kesuzuki\Desktop\出荷表自動化"
    input_dir = os.path.join(base_dir, "archive", "20260713_095050")
    template = os.path.join(base_dir, "★RETAIL_テンプレート.xlsx")
    err_log = os.path.join(base_dir, "error.log")
    
    print("Running pipeline...")
    # set NO_GUI=1 to bypass tkinter messagebox dialogs
    os.environ["NO_GUI"] = "1"
    
    output_file = "test_output.xlsx"
    success_run = run_pipeline(BRAND_FRV, output_file, input_dir, base_dir, template, err_log)
    print("Run Success:", success_run)
    
    if success_run:
        result_path = os.path.join(base_dir, "output", output_file)
        success_verify = verify_pipeline(BRAND_FRV, input_dir, result_path)
        print("Verification Success:", success_verify)
        if success_verify:
            print("[OK] Test passed completely!")
        else:
            print("[NG] Test failed verification!")
