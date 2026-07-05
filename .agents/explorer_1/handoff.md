# Handoff Report — 2026-07-05T19:56:03+09:00

本資料は、出荷明細自動集計システム `run_automation.py` の動作検証および脆弱性評価に関する調査結果の引き継ぎレポートです。

---

## 1. 観察事実 (Observation)

1. **スクリプトの基本構造 (`run_automation.py`)**:
   - 30〜33行目にて、配置ディレクトリからの相対パスを取得して実行環境に依存しない構成を取っています：
     ```python
     BASE_DIR = os.path.dirname(os.path.abspath(__file__))
     INPUT_DIR = os.path.join(BASE_DIR, "input")
     TEMPLATE_PATH = os.path.join(BASE_DIR, "★RETAIL_テンプレート.xlsx")
     ERROR_LOG_PATH = os.path.join(BASE_DIR, "error.log")
     ```
   - 67〜83行目の `find_single_csv` 関数で `glob.glob` を使用して `input/` フォルダ配下のCSVを検索し、該当ファイル数が0または2以上の場合は `show_error` を呼び出します。
   - 42〜53行目の `show_error` 関数では、`print()` による出力後に `sys.exit(1)` （`SystemExit` の発生）を呼び出し、GUI環境であれば `messagebox.showerror` を表示させます。
   - 614〜623行目の `main` 関数の `try...except Exception as e` は `Exception` クラスをキャッチし、`error.log` へのトレースバック書き込みを行います。その後 `finally` ブロックで `wb_retail.close()` を行います。
2. **検証スクリプト (`verify_all.py`)**:
   - 8〜9行目にて、`C:\Users\kxnxg\OneDrive\デスクトップ\automation\...` という絶対パスがハードコードされています。
3. **セキュリティレビュー資料 (`adversarial_review_report.md`)**:
   - 以前の COM (win32com) 版に対する4つの脆弱性指摘（①Excelプロセスの解放漏れ、②個人パス `kesuzuki` のハードコード、③列インデックス依存、④1セル書き込みのボトルネック）が記述されています。
4. **実機実行の状況**:
   - `run_command` ツールを用いた `run_automation.py` の実行は、実行環境の承認プロセスがタイムアウトしたため行えませんでした。

---

## 2. 論理の連鎖 (Logic Chain)

1. **異常系シナリオA・Bの挙動**:
   - 必須CSVの欠落（検出数0）やファイル名の重複（検出数2以上）が発生した場合、`find_single_csv` は `show_error()` を呼び出します。
   - `show_error()` は内部で `sys.exit(1)` を呼び出します。これにより `SystemExit` 例外が発生します。
   - `SystemExit` は `Exception` を継承しないため、`main()` の `except Exception` ブロックでキャッチされません。
   - したがって、**`error.log` は新規生成・更新されず**、プログラムはコンソール出力とエラーダイアログの表示を行った後、そのまま終了します（**結論セクション2に支持**）。
2. **異常系シナリオCの挙動**:
   - `RETAIL_完成版.xlsx` が Excel で開かれている場合、603行目の `os.remove` は `PermissionError` を投げますが、`try-except` で握りつぶされます。
   - その後、607行目の `wb_retail.save` で `PermissionError` が発生します。
   - この例外は `Exception` であるため `main()` の `except Exception` でキャッチされ、**`error.log` にスタックトレースが記録され、GUI詳細エラーが表示されます**。
   - `finally` ブロックが実行されるため、オープンした openpyxl のワークブックは正しくクローズされます（**結論セクション2に支持**）。
3. **異常系シナリオDの挙動**:
   - テンプレートファイルが存在しない場合、297行目の `openpyxl.load_workbook` で `FileNotFoundError` が発生します。
   - これも `main()` の `except Exception` でキャッチされ、**`error.log` に記録されます**（**結論セクション2に支持**）。
4. **脆弱性の影響分析**:
   - **①Excelプロセスの解放漏れ**: win32com ではなく openpyxl で動くため、`EXCEL.EXE` プロセス自体が発生せず、クリーンアップ時の `taskkill` も不要となったため、**完全に解決**しています。
   - **②個人パスのハードコード**: `BASE_DIR` による相対パス解決が導入されたため、`error.log` の出力を含め**完全に解決**しています。
   - **③列インデックス依存**: CSV読み込みは Pandas の列名指定で改善されましたが、Excel 操作側（`delete_cols` や固定列参照の計算式設定ループ）で物理列インデックスのハードコーディングが多数残存しています。フォーマット小変更でサイレントバグが発生するリスクは**依然として極めて高い**です。
   - **④1セル書き込みのボトルネック**: win32com の通信オーバーヘッドは解消したため速度は大幅に改善しましたが、依然として 10,000行 × 29列 ＝ 29万回 などの不要セルクリア用二重ループ処理が残っており、openpyxl 上での処理遅延や無駄なハングアップの原因となっています。

---

## 3. 制約事項 (Caveats)

- **実機実行の可否**: 実行環境における `run_command` の承認待ちタイムアウトのため、実際のコマンドライン実行結果は取得できていません。本報告書に記載されている各異常系シナリオのコンソール出力、エラーダイアログ、および `error.log` の記述は、すべてソースコードの静的解析およびハンドリング仕様に基づいた論理シミュレーションの結果です。実際のメッセージボックスのボタン形状や表示レイアウトは Windows OS の仕様に依存します。

---

## 4. 結論 (Conclusion)

1. **エラーログ・ファイルロック挙動の特定**:
   - CSV欠落や重複などのファイルエラー（シナリオA, B）では、`SystemExit` による即時終了となるため `error.log` は生成されません。
   - 完成版オープン中やテンプレート紛失（シナリオC, D）では、`error.log` が正しく生成され、スタックトレースが記録されます。
2. **既存脆弱性の影響評価**:
   - プロセス解放漏れ（①）と個人パス（②）は完全に解消されました。
   - 列インデックス依存（③）は依然として Excel 操作側で重大なリスクとして残存しています。
   - 1セル書き込みループ（④）は、openpyxl 移行で速度は向上したものの、不要セルの大量クリア二重ループ（約39万回）によるパフォーマンス上の懸念が残されています。

---

## 5. 検証方法 (Verification Method)

以下の手動テストにより、本報告書の論理的結論を独立して検証することができます。

1. **シナリオA（CSV不足）の検証**:
   - `input/` フォルダの `営業日付別売上分析(旧) (1).csv` を一時的にデスクトップなどへ退避させ、`python run_automation.py` を実行します。
   - コンソールに `input/ フォルダに該当するCSVが見つかりません` と表示されること、および `error.log` が新規作成または更新されないことを確認します。
2. **シナリオC（ファイルロック）の検証**:
   - `RETAIL_完成版.xlsx` を Excel で開いた状態で `python run_automation.py` を実行します。
   - `[Errno 13] Permission denied` エラーが発生すること、および `error.log` にそのトレースバックが保存されることを確認します。
3. **成果物ファイルの確認**:
   - `C:\Users\kxnxg\OneDrive\デスクトップ\automation\.agents\orchestrator_adversarial\explorer_report.md` を開き、本調査のまとめが出力されていることを目視確認します。
