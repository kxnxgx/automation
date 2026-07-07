# Project: 出荷表自動化 誤発注防止ロジック修正

## Architecture
- `run_automation.py` 等の自動化スクリプトが動作する。
- 共通処理は `automation_core.py` に切り出されている。
- 在庫不足時に、誤発注防止ロジックがトリガーされる。

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| 1 | M1: 調査 | 誤発注防止ロジックの現状実装およびテストの動作状況調査 | none | IN_PROGRESS |
| 2 | M2: 修正とドキュメント更新 | automation_core.py の修正および docs/REVIEW_REPORT.md の更新 | M1 | PLANNED |
| 3 | M3: 検証 | テストスクリプトの実行およびインテグリティ監査 | M2 | PLANNED |

## Code Layout
- `automation_core.py`: 共有モジュール（集計コアロジックおよび誤発注防止ロジックを含む）
- `run_automation.py`等: 各ブランド向けの個別実行スクリプト
- `verify_all.py`等: 検証・テスト用スクリプト
- `docs/REVIEW_REPORT.md`: レビュー結果および仕様のドキュメント
