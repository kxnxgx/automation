# タスクリスト

## run_automation.py の改修
- [ ] 独自例外 `AutomationError` の定義と、エラー時の `sys.exit(1)` を例外スローに修正
- [ ] `main()` で `AutomationError` と `Exception` の例外処理を行い、必ず `error.log` にトレースバックを出力
- [ ] Tkinterのダイアログとログの冒頭メッセージを、非エンジニア向けに分かりやすい日本語メッセージ「[発生した状況] \n\n[対処方法]」の形式で統一し、専門用語を排除
- [ ] 完成版Excel（`RETAIL_完成版.xlsx`）が開いたまま実行された際の `PermissionError` をキャッチし、分かりやすい通知を表示
- [ ] `step2_python_direct_merge` 内での列操作・セル書き込み・数式ループにおいて、ヘッダー名を走査して対象列のインデックス（列番号）を動的に検索・解決する仕組みに修正
- [ ] 不要セルクリア処理（合計39万回のループ）を実際の最終行または妥当なクリア範囲に対して一括で処理するようリファクタリング

## README.md の改修
- [ ] 表示または `error.log` に出力する日本語エラーメッセージと、トラブルシューティング項目を100%合致させるように修正・加筆

## 動作確認と検証
- [ ] 正常系テスト（`verify_all.py` を実行して「すべて一致しています！」と表示されるか確認）
- [ ] 異常系テスト（CSVファイル欠損、Excelファイル開きっぱなし）での挙動・エラー表示・ログ出力の確認
- [ ] 検証結果を `C:\Users\kxnxg\OneDrive\デスクトップ\automation\.agents\orchestrator_adversarial\worker_report.md` に保存

## ドキュメントの作成
- [ ] docs/automation_improvement/walkthrough.md の作成
