# Project: 出荷表自動化のレビュー・バグ修正・モジュール化

## Architecture
- `run_automation*.py`: 入力データ（CSV）からブランドごとの出荷明細・売上集計を自動作成するスクリプト群（FRV, tennen, hanwag）。
- `verify_all*.py`: 自動集計結果と元のCSVデータを照合して検証するスクリプト群。
- `debug_diff.py`: 検証やデバッグ時にデータ差異を特定するための補助スクリプト。
- `実行*.bat`: 非エンジニアがダブルクリックで各ブランドの集計を実行・検証するためのバッチファイル群。
- 本プロジェクトの目的:
  1. `run_automation*.py` に潜む集計バグやエッジケース（マイナス値、ブランドフィルタリング、合算仕様、Excel読み込み等）の特定と修正。
  2. 重複コード（50行以上）を共有モジュール（`automation_core.py`）に切り出すリファクタリング。
  3. `verify_all*.py` の統合・共通化、検証項目の強化、`debug_diff.py` の改善。
  4. バッチファイル（`実行.bat`, `実行_FRV.bat`, `実行_hanwag.bat`, `実行_tennen.bat`）の整合性確認と動作確認。
  5. 成果物として `docs/REVIEW_REPORT.md` (日本語) の作成と、必要に応じた `README.md` の更新。

## Milestones
| # | Name | Scope | Dependencies | Status | Conversation ID |
|---|------|-------|-------------|--------|-----------------|
| 8 | M8: 解析とバグ特定 | 各Pythonスクリプトの静的解析とバグ・仕様不整合の特定、リファクタリング方針策定 | なし | DONE | 97c8ab35, 6e2315d7, 790ec4f8 |
| 9 | M9: バグ修正と共通コードモジュール化 | 共有モジュール `automation_core.py` の作成、各集計スクリプトのリファクタリングおよびバグ修正 | M8 | DONE | be71f419 |
| 10| M10: 検証・デバッグスクリプトの改善 | `verify_all*.py` の共通モジュール化と検証項目の強化、`debug_diff.py` のリファクタリング・改善 | M9 | DONE | 796e2ca5 |
| 11| M11: バッチファイルの整合性確認と動作検証 | `実行*.bat` を用いた動作テストと不具合修正、実機検証 | M10 | DONE | a7b5c47e |
| 12| M12: ドキュメンテーションと最終報告 | `docs/REVIEW_REPORT.md` の作成、`README.md` の更新、監査パス | M11 | DONE | be47c704, b94048ce |

## Code Layout
- `c:\Users\kesuzuki\Desktop\出荷表自動化\run_automation.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\run_automation_tennen.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\run_automation_hanwag.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\verify_all.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\verify_all_tennen.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\verify_all_hanwag.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\debug_diff.py`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\実行.bat`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\実行_FRV.bat`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\実行_tennen.bat`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\実行_hanwag.bat`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\docs\REVIEW_REPORT.md`
- `c:\Users\kesuzuki\Desktop\出荷表自動化\README.md`
