# 実行コンテキスト (Context)

## ディレクトリとファイル構成
- **自動化フォルダ**: `C:\Users\kxnxg\OneDrive\デスクトップ\automation`
  - 主要スクリプト: `run_automation.py` (自動集計メインスクリプト)
  - 実行ファイル: `実行.bat` (非エンジニアが起動するためのバッチファイル)
  - 検証スクリプト: `verify_all.py` (集計値とCSVの合計値を比較検証するスクリプト)
  - テンプレート: `★RETAIL_テンプレート.xlsx` / `★RETAIL_CSV2計算式.xlsm`
  - 入力用フォルダ: `input/`
- **ワークスペースフォルダ**: `C:\Users\kxnxg\antigravity\valiant-oppenheimer`
  - ここに各種管理ドキュメントを配置する。

## 目標・制約
- **マニュアルの対象読者**: 非エンジニア（PC操作に不慣れなスタッフ）。
- **用語制限**: 「スクリプト」「ディレクトリ」「パス」「引数」などの技術用語はNG。「フォルダ」「ダブルクリック」「ファイルのコピー」などの日常PC用語に言い換える。
- **ドキュメント出力先**:
  - 対象マニュアル: `C:\Users\kxnxg\OneDrive\デスクトップ\automation\README.md`
  - 管理用ドキュメント:
    - `C:\Users\kxnxg\antigravity\valiant-oppenheimer\docs\shipment_aggregation_manual\task.md`
    - `C:\Users\kxnxg\antigravity\valiant-oppenheimer\docs\shipment_aggregation_manual\implementation_plan.md`
    - `C:\Users\kxnxg\antigravity\valiant-oppenheimer\docs\shipment_aggregation_manual\walkthrough.md`
