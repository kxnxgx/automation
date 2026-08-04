# Project: 出荷表自動化 誤発注防止ロジック修正

## Architecture
- `run_automation.py` 等の自動化スクリプトが動作する。
- 共通処理は分割モジュール化され、`automation_core.py` が統括する。
- 在庫不足時に、誤発注防止ロジック（全チャネル売上ゼロ化）がトリガーされる。
- 店舗数は動的に解析され、列ズレを起こさない。

## Milestones
| # | Name | Scope | Dependencies | Status |
|---|------|-------|-------------|--------|
| 1 | M1: 調査 | 誤発注防止ロジックの現状実装およびテストの動作状況調査 | none | DONE |
| 2 | M2: 修正とドキュメント更新 | automation_core.py の修正および docs/REVIEW_REPORT.md の更新 | M1 | DONE |
| 3 | M3: 検証 | テストスクリプトの実行およびインテグリティ監査 | M2 | DONE |
| 4 | M4: FRVファイル精査とレイアウト修正 | 出荷予定日列(gap3/BQ列)対応、D列/A列数値書式設定、BJ~BP列非表示対応 | M3 | DONE |

## Code Layout
- `config.py` : ブランド構成や共通エラークラスの定義
- `utils.py` : ブランド判定、店舗構造解析（gap1/gap2/gap3動的検出）、ダイアログ表示等の共通処理
- `data_loader.py` : 各種CSV・Excelファイルのデータロード・集計処理
- `generate_sheets.py` : Excelシートの構築、マージ、数式書き込み、書式適用（数値書式・非表示列設定含む）
- `automation_core.py`: 各モジュールを呼び出し、全体処理を制御するメイン窓口（および再エクスポート窓口）
- `run_automation.py`等: 各ブランド向けの個別実行スクリプト
- `verify_all.py`等: 検証・テスト用スクリプト
- `docs/REVIEW_REPORT.md`: レビュー結果および仕様のドキュメント
- `docs/自動化引継ぎ資料.md`: システム設計および仕様の引継ぎ資料
- `docs/引き継ぎ書.md`: システム運用・保守・引継ぎマニュアル

