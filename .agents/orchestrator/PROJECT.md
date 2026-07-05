# Project: 出荷明細自動集計マニュアル作成

## Architecture
- `run_automation.py` は、入力データ（CSV）を読み込んで出荷明細の自動集計を行うスクリプト。
- `実行.bat` は、非エンジニアがダブルクリックで実行できるようにするためのバッチファイル。
- `verify_all.py` は、自動集計結果（RETAIL_完成版.xlsx）と元のCSVの合計値を比較して検証するスクリプト。
- 本プロジェクトの目的は、非エンジニア向けに専門用語を排除した分かりやすい取扱説明書（README.md）と、管理用ドキュメントを作成すること。

## Milestones
| # | Name | Scope | Dependencies | Status | Conversation ID |
|---|------|-------|-------------|--------|-----------------|
| 4 | M4: 調査・検証 | スクリプトおよびバッチの動作・入出力仕様の確認と実証 | なし | DONE | e7915e82-4380-4b37-80d7-0ee2006a8822 |
| 5 | M5: 管理ドキュメント作成 | `task.md`, `implementation_plan.md`, `walkthrough.md` の作成 | M4 | DONE | 658cb20a-79ef-4bdb-a112-ca45268d7607 |
| 6 | M6: 取扱説明書作成 | `README.md` の作成 | M5 | DONE | 658cb20a-79ef-4bdb-a112-ca45268d7607 |
| 7 | M7: 最終検証 | 動作・マニュアル記載内容の突き合わせ検証 | M6 | DONE | 5e00728f-8170-4ca5-bb41-b1b05764c277 |

## Code Layout
- `C:\Users\kxnxg\OneDrive\デスクトップ\automation\run_automation.py` - メインスクリプト
- `C:\Users\kxnxg\OneDrive\デスクトップ\automation\実行.bat` - バッチファイル
- `C:\Users\kxnxg\OneDrive\デスクトップ\automation\verify_all.py` - 検証スクリプト
- `C:\Users\kxnxg\OneDrive\デスクトップ\automation\README.md` - 取扱説明書（成果物）
- `C:\Users\kxnxg\antigravity\valiant-oppenheimer\docs\shipment_aggregation_manual\` - 管理用ドキュメント保存先
