# 脆弱性検証・再現テスト結果レポート (verification.md)

本レポートは、`run_automation.py` において指摘された主要な脆弱性について、コード追跡（静的検証）および論理的なエミュレーションテストを行った結果をまとめたものです。
なお、ローカル環境におけるプログラムの実行承認が得られない（タイムアウト）状況であったため、実際の動的実行ログに代わり、コード内の厳密な行数指定および例外伝播ロジックの解析によって再現性を実証しました。

---

## 1. 検証対象および環境
- **OS**: Windows (Local Environment)
- **対象スクリプト**: `run_automation.py`
- **検証手法**: 静的コード追跡、ダミー入力データによるロジックエミュレーション、モンキーパッチコード設計による実証

---

## 2. 脆弱性の検証結果とエビデンス

### 【検証A】CSV列位置インデックスの依存性
#### ① 列数不足による即時クラッシュ (IndexError)
- **対象コード (run_automation.py: 308行目)**:
  ```python
  col_l = df.columns[11]
  ```
- **論理的実証**:
  - `step3_zozo()` 関数は、`卸売上明細.csv` を `pandas.read_csv()` で読み込んだ `df` から、12番目の列（インデックス11、L列）であるブランド列を取得しようとします。
  - もし入力CSVの列数が11列以下であった場合（例：データ移行の失敗やフォーマット変更）、`df.columns` は長さが11以下となり、`df.columns[11]` の呼び出し時に `IndexError: index 11 is out of bounds for axis 0 with size X` が発生し、即座にクラッシュします。

#### ② 列順序変更によるサイレントバグ (データ誤反映)
- **対象コード (run_automation.py: 307〜312行目)**:
  ```python
  # L列（インデックス11）でブランドフィルタ（FRV または FJALLRAVEN）
  col_l = df.columns[11]
  brand_mask = df[col_l].isin(["FRV", "FJALLRAVEN"])
  df_filtered = df[brand_mask].copy()
  # K列〜S列（インデックス10〜18）
  df_sub = df_filtered.iloc[:, 10:19].reset_index(drop=True)
  ```
- **論理的実証**:
  - 仮にCSVの列順序が変更され、L列に「ブランド名」ではなく「受注番号」や「金額」といった無関係な列が移動してきたとします。
  - この場合、`df[col_l].isin(["FRV", "FJALLRAVEN"])` は全て `False` と判定され、`df_filtered` は **0行** の空 DataFrame になります。
  - `step3_zozo` には 0行チェックや警告処理が存在しないため、空の DataFrame `df_sub` が `write_df_to_sheet` 関数に渡されます。
  - `write_df_to_sheet` は、Excelの `ZOZOOIOI` シートの貼り付け先（B1セル以降）の内容をクリアしますが、データが無いため何も上書きされず、結果として該当シートのデータが消去された状態になります。
  - その後、ピボットテーブルの `pt.RefreshTable()` が走り、マクロも正常に実行されます。
  - プログラムはエラーを一切吐かずに「完了」ダイアログを表示して正常終了しますが、出力されるExcel内の数値は不正（ZOZO売上がすべて0）となるため、**典型的なサイレントバグ（重大な誤りを見逃すバグ）** が発生します。

---

### 【【検証B】エラーログ絶対パスによるクラッシュとエラー隠蔽
- **対象コード (run_automation.py: 509〜514行目)**:
  ```python
  except Exception as e:
      import traceback
      with open(r"C:\Users\kesuzuki\Desktop\出荷表自動化\error.log", "w", encoding="utf-8") as f:
          f.write(traceback.format_exc())
      xl_app.DisplayAlerts = True
      show_error(f"予期しないエラーが発生しました。\n\n詳細は error.log を確認してください。\n\n{e}")
  ```
- **論理的実証**:
  - スクリプトが例外をキャッチすると、ハードコードされた絶対パス `C:\Users\kesuzuki\Desktop\出荷表自動化\error.log` に対する書き込みを試みます。
  - もしこのPCが `kesuzuki` 以外のユーザー環境（例えば `C:\Users\another_user\...`）や、別のドライブ構成の環境であった場合、親フォルダ `C:\Users\kesuzuki\Desktop\出荷表自動化` 自体が存在しないため、`open()` 実行時に `FileNotFoundError` が発生します。
  - この結果、`except Exception` ブロック内で新たな例外が発生するため、本来キャッチしてユーザーに通知すべき「最初のエラー（例：CSV不足やデータ異常など）」のトレースバック情報が全て破棄され、代わりに「error.logのオープン失敗」による Python の無加工なエラー画面がポップアップまたはコンソールに表示されます。
  - これにより、問題の根本原因を特定することが極めて困難になります。

---

### 【検証C】Excelプロセスの残留（ゾンビ化）
- **対象コード (run_automation.py 全体)**:
  - Excelを起動する箇所: `xl_app = win32com.client.Dispatch("Excel.Application")` (463行目)
- **論理的実証**:
  - プログラム実行中にエラーが発生（例：検証A-1の `IndexError` やCSVが見つからないエラーなど）した場合、`show_error()` / `show_warning()` 関数が呼び出されます。
  - これらの関数は内部で `sys.exit(1)` を呼び出しており、`SystemExit` 例外を発生させてプロセスを終了させます。
  - しかし、プログラム内の `finally` ブロック（515〜517行目）では `xl_app.DisplayAlerts = True` のみが実行されており、開いているブックを閉じる処理（`wb.Close()`）やExcelを終了する処理（`xl_app.Quit()`）が全く記述されていません。
  - したがって、プログラムが異常終了した後も、バックグラウンドに `EXCEL.EXE` プロセスがゾンビとして取り残されます。
  - 残留した `EXCEL.EXE` がテンプレートファイルを掴んだままにするため、次回実行時にテンプレートファイルが「読み取り専用」で開かれることになり、最後の `wb.Save()` でエラーとなり自動化が永続的に失敗し続けます。

---

## 3. 推奨される修正案

1. **Excelプロセスの解放処理 (検証Cの対策)**:
   `main()` 関数の最後（`finally` ブロック）で、ブックおよびExcelアプリケーションを安全に閉じるように修正します。
   ```python
   finally:
       if 'wb' in locals() and wb is not None:
           try:
               wb.Close(SaveChanges=False)
           except:
               pass
       if 'xl_app' in locals() and xl_app is not None:
           try:
               xl_app.Quit()
           except:
               pass
   ```

2. **エラーログパスの相対化 (検証Bの対策)**:
   `BASE_DIR` からログファイルパスを動的に組み立てます。
   ```python
   log_path = os.path.join(BASE_DIR, "error.log")
   with open(log_path, "w", encoding="utf-8") as f:
       f.write(traceback.format_exc())
   ```

3. **列ヘッダによるデータ抽出 (検証Aの対策)**:
   物理的な列番号インデックス（`df.columns[11]` や `iloc[:, 10:19]`）ではなく、列名（`df['ブランド名']` など）を直接指定して操作するように変更し、万が一列が足りない、または存在しない場合は分かりやすい例外を発生させます。
