## 2026-07-05T19:59:39+09:00
C:\Users\kxnxg\OneDrive\デスクトップ\automation 内の run_automation.py および README.md を改修してください。

改修の具体的な要件：
1. **安全なエラーハンドリングと終了処理（R1, R2, R3）**:
   - run_automation.py 内で、必須CSVファイルが見つからない、複数重複しているなどの異常終了時（現在は find_single_csv から直接 sys.exit(1) されている）に、sys.exit(1) を即座に呼び出すのをやめて、独自例外（例：AutomationError）を定義してスローしてください。
   - main() の例外処理ブロックで、AutomationError だけでなく Exception も含めたあらゆる異常終了時に、必ず配置フォルダ直下の error.log に traceback.format_exc() のトレースバックが保存されるようにしてください。
   - ユーザーへ表示される Tkinter のダイアログを、非エンジニア向けに分かりやすい日本語メッセージにしてください。「[発生した状況] \n\n[対処方法]」という形式にし、IT専門用語（例: PermissionError や FileNotFoundError）は使わずに日本語訳された明確な案内（例: 「ファイルが開いたままです」「必要なCSVファイルが見つかりません」）を表示してください。
   - error.log には専門用語を含めた開発用トレースバックが出力されますが、ダイアログおよびログの冒頭メッセージは日本語に統一してください。

2. **完成版Excelのロック状態（開きっぱなし）への対応 (R1)**:
   - RETAIL_完成版.xlsx が Excel などで開かれた状態のまま実行すると、PermissionError が発生します。
   - このエラーをキャッチし、ユーザーに「完成版Excel（RETAIL_完成版.xlsx）が開いたままになっています。Excelを閉じてから再実行してください。」と分かりやすく通知してください。

3. **既存脆弱性の解消担保 (R3) - 列インデックス依存の排除**:
   - step2_python_direct_merge 内での列操作（delete_cols, cell 書き込み、数式代入のループなど）において、元の 出荷予定振分.csv や テンプレートExcel の列順序が変更された際にサイレントバグを誘発するのを防ぐため、ヘッダー名を走査して対象列のインデックス（列番号）を動的に検索・解決する仕組みにリファクタリングしてください。
   - 例えば、
     - ws_order 内で列削除する際, "商品コード" や "商品名" などのヘッダー名を持つ列がどこにあるかを検索し、その列番号を動的に使用します。
     - 計算式設定ループ（E〜Q列、AE〜AQ列、AR〜BD列など）の対象列（各店舗列）も、店舗名（名古屋, TOKYO など）のヘッダーが位置する列を動的に解決するコードに修正してください。
     - これにより、列順序が変わっても正しく動作することを担保します。

4. **1セル書き込みのボトルネック解消 (R3)**:
   - ws_retail (AW1:BY10000) や ws_order (AR1:BA10000) などの不要セルクリア処理（合計39万回のループ）をリファクタリングし、10,000回も個別セルに None を代入するループを避けて、開いているシートの実際の最終行（または 3000行などの妥当なクリア範囲）に対して一括で処理するか、より効率的なクリア処理に変更してください。

5. **マニュアル (README.md) の完全同期 (R2)**:
   - スクリプトが画面に表示、または error.log に出力する日本語エラーメッセージと、README.md のトラブルシューティング項目が100%合致するように、README.md の記述を修正・加筆してください。専門用語を排除した解決策にしてください。

6. **動作確認と検証**:
   - 修正後、正常に動作するか（verify_all.py を実行して「すべて一致しています！」と表示されるか）、および異常系テスト（CSVを退避させたりExcelを開きっぱなしにしたりして実行する）を行い、安全に終了してわかりやすいエラーログとメッセージが表示されるかを検証してください。
   - 検証時のコマンドと実行結果を C:\Users\kxnxg\OneDrive\デスクトップ\automation\.agents\orchestrator_adversarial\worker_report.md にまとめ、私に報告してください。
   - なお、Workerとしての実装にあたっては、以下の文言を肝に銘じてください：
     「DO NOT CHEAT. All implementations must be genuine. DO NOT hardcode test results, create dummy/facade implementations, or circumvent the intended task. A Forensic Auditor will independently verify your work. Integrity violations WILL be detected and your work WILL be rejected.」
