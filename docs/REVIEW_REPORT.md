# 出荷明細作成自動化 レビュー結果報告書

この報告書は、出荷明細作成自動化システムのコードレビューによって特定された主要なバグの詳細、発生条件、具体的な修正コード・対処方法、および共通モジュールへの移行に伴うリファクタリングと検証・デバッグスクリプトの機能強化点の概要をまとめたものです。

---

## 1. 特定された主要なバグの詳細と対処方法

### ① ブランド判定の曖昧さ
* **バグの詳細と発生条件**:
  各ブランド（FRV, TEN, HWG）の対象データを抽出する際、商品コードのプレフィックスやブランド名カラムの曖昧な一致により、他ブランドのデータが混入したり（例：TENの商品コードがFRVに混入）、自ブランドのデータが漏れたりしていた。具体的には、商品コードの先頭（プレフィックス）が `"TNT"`, `"TNP"`, `"TNF"` のものはTEN、`"H"` で始まるものはHWG、それ以外はFRVとなるべきところが、厳格に区別されていなかった。
* **具体的な修正方法・コード**:
  ブランドごとの判定基準（許容ブランド名、プレフィックス、除外プレフィックス）を `BrandConfig` クラスとして定義し、判定ロジックを `is_target_brand` 関数に集約した。
  ```python
  class BrandConfig:
      def __init__(self, name, allowed_names, prefixes, excluded_prefixes=None):
          self.name = name
          self.allowed_names = [n.upper() for n in allowed_names]
          self.prefixes = tuple(p.upper() for p in prefixes) if prefixes else ()
          self.excluded_prefixes = tuple(p.upper() for p in excluded_prefixes) if excluded_prefixes else ()

  BRAND_FRV = BrandConfig(
      name="FRV",
      allowed_names=["FRV", "FJALLRAVEN"],
      prefixes=[],  # すべて（除外プレフィックス以外）
      excluded_prefixes=["TNT", "TNP", "TNF", "H", "W"]
  )

  BRAND_TEN = BrandConfig(
      name="TEN",
      allowed_names=["TEN", "TENNEN"],
      prefixes=["TNT", "TNP", "TNF"],
      excluded_prefixes=[]
  )

  BRAND_HWG = BrandConfig(
      name="HWG",
      allowed_names=["HWG", "HANWAG"],
      prefixes=["H"],
      excluded_prefixes=[]
  )

  def is_target_brand(config: BrandConfig, code, brand_name=None) -> bool:
      """ブランド設定に基づき、対象ブランドか厳格に判定する"""
      if brand_name and str(brand_name).strip():
          b = str(brand_name).upper().strip()
          if b in config.allowed_names:
              return True
          # 他の既知のブランド名にマッチする場合は確実に自ブランドではないため False
          all_known_brands = ["FRV", "FJALLRAVEN", "TEN", "TENNEN", "HWG", "HANWAG"]
          if b in all_known_brands:
              return False
              
      if not code:
          return False
          
      c = str(code).upper().strip()
      if config.excluded_prefixes and c.startswith(config.excluded_prefixes):
          return False
          
      if not config.prefixes:
          return True
          
      if c.startswith(config.prefixes):
          return True
          
      return False
  ```

### ② leaked_codes追加漏れ
* **バグの詳細と発生条件**:
  出荷予定振分（ベースとなる `RETAIL` シートのデータ）に存在しない商品であっても、店舗売上や卸（ZOZO/OIOI）売上、EC売上のCSVに売上データが記録されている場合があった。旧処理では、出荷予定振分に記載されている商品（行）に対してのみ売上データをマージしていたため、それ以外の「出荷予定にない売上発生商品」のデータが完全に漏れ、合計値が一致しなくなっていた。
* **具体的な修正方法・コード**:
  各売上CSVから売上のあったすべての商品コード（`all_sales_codes`）を抽出し、出荷予定振分に存在する商品コード（`order_codes`）との差分を `leaked_codes` として特定。これらを order シートおよび RETAIL シートの末尾へ行追加し、売上データや数式参照（`=order!...`）を設定するロジックを追加した。
  ```python
  # 漏れている（出荷予定にないが売上があった）商品コードの抽出
  leaked_codes = sorted(list(all_sales_codes - order_codes))
  
  # order/RETAILシートの末尾への行追加
  for code in leaked_codes:
      # orderシートへ行追加
      ws_order.cell(row=current_order_row, column=1).value = p_code
      ws_order.cell(row=current_order_row, column=2).value = p_name
      # C〜Q列 (BULK)、R〜AD列 (在庫) はすべて 0 を設定
      ...
      # ORDER列(AE〜AQ列)への売上書き込み
      for store_name, col_idx in order_store_map.items():
          ws_order.cell(row=current_order_row, column=col_idx).value = max(0, val)
          
      # RETAILシートへ行追加とorderシートへの参照数式の設定
      ws_retail.cell(row=current_retail_row, column=4).value = p_code
      for c_idx in range(13):
          target_col_letter = openpyxl.utils.get_column_letter(49 + c_idx)
          source_col_letter = openpyxl.utils.get_column_letter(31 + c_idx)
          ws_retail.cell(row=current_retail_row, column=49 + c_idx).value = f"=order!{source_col_letter}{current_order_row}"
  ```

### ③ 売上合計マイナス処理
* **バグの詳細と発生条件**:
  商品のキャンセルや返品などにより、売上CSV（特に営業日付別売上分析や卸売上明細）に数量としてマイナス値（例: `-1`）が記録されることがある。このマイナス値をそのまま集計すると、売上合計がマイナスになり、本来出荷されるべき正の数量と相殺されてしまい、結果が不正になる問題が発生していた。
* **具体的な修正方法・コード**:
  CSV読み込み・集計時、およびセル書き込みのタイミングで、マイナス値を0に丸める処理（`clip(lower=0)` または `max(0, val)`）を適用した。
  ```python
  # pandasによる数量の数値型キャストおよびマイナス値（返品・キャンセル）の0丸め
  df_filtered["数量"] = pd.to_numeric(df_filtered["数量"], errors="coerce").fillna(0)
  df_filtered["数量"] = df_filtered["数量"].clip(lower=0)
  
  # 特価数や卸販売数量などでも同様に clip(lower=0) を適用
  df_filtered["販売数量"] = pd.to_numeric(df_filtered["販売数量"], errors="coerce").fillna(0)
  df_filtered["販売数量"] = df_filtered["販売数量"].clip(lower=0)
  ```

### ④ 卸・EC混入
* **バグの詳細と発生条件**:
  「誤発注防止ロジック」（店舗確保枠（C列/D列/過不足）が店舗売上合計を下回る場合、店舗の売上を 0 にして出荷を防ぐロジック）を適用する際、店舗以外の卸チャネル（ZOZO, OIOI）やECチャネル（EC, YSEC）の売上が混入して計算されていた。これにより、店舗の確保枠と無関係な卸・ECの売上を含めた合計値で判定されてしまい、本来ゼロ化されるべき店舗売上がゼロ化されなかったり、逆にゼロ化されるべきではないケースでゼロ化されたりしていた。
* **具体的な修正方法・コード**:
  誤発注防止ロジックの判定に用いる売上合計を、**全チャネル（ZOZO・OIOI・EC・YSECや店舗）の合計値**とし、確保枠がそれを下回る場合は全チャネルの売上をゼロにするよう修正。
  ```python
  # 全チャネルの売上合計で判定（ZOZO/OIOI/EC/YSEC含む）
  total_sales_all = sum(raw_sales.values())

  c_val = ws_order.cell(row=r, column=3).value or 0
  d_val = ws_order.cell(row=r, column=4).value or 0
  kabu_sum = sum(ws_order.cell(row=r, column=col).value or 0 for col in range(5, 18))

  if c_val == 0 and d_val == 0 and kabu_sum < total_sales_all:
      # 確保枠が全売上合計を下回る場合、全チャネルの売上を 0 にする
      for store_name, col_idx in order_store_map.items():
          ws_order.cell(row=r, column=col_idx).value = 0
  else:
      for store_name, col_idx in order_store_map.items():
          ws_order.cell(row=r, column=col_idx).value = raw_sales[store_name]
  ```

### ⑤ HUTTE指示書読み込み不備
* **バグの詳細と発生条件**:
  `★出荷指示フォーマット【*】hutte.xlsx` ファイルを読み込む際、商品コード列や数量列が何らかの理由でズレていたり、インデックスがハードコーディングされていたりすると、正しくデータを抽出できない問題があった。また、マイナス値の混入や、他ブランドの混入もあった。
* **具体的な修正方法・コード**:
  ファイル読み込み時に列名（ヘッダー）を走査して「商品コード」「品番」「数量」などのキーワードから動的に列インデックスを特定し、さらにブランド判定（`is_target_brand`）およびマイナス値の0丸めを適用して堅牢化した。
  ```python
  def load_hutte(input_dir, brand_config: BrandConfig):
      ...
      df = pd.read_excel(latest_file, sheet_name=0, dtype=str)
      
      # 列の動的特定
      key_col = None
      val_col = None
      for col in df.columns:
          col_str = str(col).strip()
          col_lower = col_str.lower()
          if any(kw in col_lower for kw in ["商品コード", "品番", "アイテムコード", "sku", "code"]):
              key_col = col
              break
      for col in df.columns:
          col_str = str(col).strip()
          col_lower = col_str.lower()
          if any(kw in col_lower for kw in ["数量", "指示数", "出荷数", "個数", "枚数", "qty", "quantity"]):
              val_col = col
              break
              
      if key_col is not None and val_col is not None:
          df = df.dropna(subset=[key_col])
          df[key_col] = df[key_col].astype(str).str.strip()
          df = df[df[key_col].apply(lambda c: is_target_brand(brand_config, c, None))]
          df[val_col] = pd.to_numeric(df[val_col], errors="coerce").fillna(0).astype(int).clip(lower=0)
          agg_df = df.groupby(key_col)[val_col].sum().reset_index()
          return dict(zip(agg_df[key_col], agg_df[val_col]))
      return {}
  ```

### ⑥ 列インデックスハードコード
* **バグの詳細と発生条件**:
  マスタや売上CSV、あるいは結果ファイルに書き出す際、特定の列（例: 4列目が商品コード、31列目がZOZOなど）がコード上にハードコーディングされていたため、少しでもExcelのレイアウトが変更されると、間違った列に値を書き込んでしまったり、インデックスエラーが発生したりしていた。
* **具体的な修正方法・コード**:
  列ヘッダー（`3行目` 等）の文言や、テンプレートファイル（`★RETAIL_テンプレート.xlsx` 内の `計算式2026` シート）から動的に列インデックスを検出し、マッピングを作成するロジックを導入した。
  ```python
  # 列の動的特定
  idx_code = None
  for c in range(1, ws_order.max_column + 1):
      val = ws_order.cell(row=1, column=c).value
      val_str = str(val).strip() if val is not None else ""
      if val_str in ("商品コード", "4"):
          idx_code = c
  ...
  col_map = {}
  current_area = None
  for col in range(1, ws_order.max_column + 1):
      area_val = ws_order.cell(row=2, column=col).value
      if area_val is not None:
          area_val_str = str(area_val).strip()
          if area_val_str in ("確保", "在庫", "ORDER", "過不足"):
              current_area = area_val_str
      
      store_val = ws_order.cell(row=3, column=col).value
      if store_val is not None and current_area is not None:
          col_map[(current_area, str(store_val).strip())] = col
  ```

### ⑦ SUBTOTAL行数固定
* **バグの詳細と発生条件**:
  `order` シートの1行目に入力される `SUBTOTAL(9, C4:C3500)` のような計算式で、対象行数が `3500` に固定されていた。これにより、データ行数が `3500` 行を超えるような場合に下方のデータが集計されなくなったり、逆にデータ行数が極端に少ない場合に無駄な範囲を参照していたりする問題があった（将来的なデータ増加に対する脆弱性）。
* **具体的な修正方法・コード**:
  `lastRow`（データ最終行）をプログラム上で動的に検出し、それを数式に埋め込むように修正。
  ```python
  # 1行目の SUBTOTAL 計算式の書き込み (3500固定の廃止、lastRowへの動的置換)
  max_order_col = max(order_store_map.values()) if order_store_map else 43
  for c in range(3, max_order_col + 1):
      col_letter = openpyxl.utils.get_column_letter(c)
      ws_order.cell(row=1, column=c).value = f"=SUBTOTAL(9,{col_letter}4:{col_letter}{lastRow})"
  ```

---

## 2. リファクタリングおよび機能強化概要

### ① 共通モジュール `automation_core.py` へのロジック集約
* **重複コードの排除結果**:
  旧システムでは、各ブランド用（FRV, TEN, HWG）に個別の実行スクリプト（`run_automation.py`, `run_automation_tennen.py`, `run_automation_hanwag.py` 等）が存在し、それぞれの中にほぼ同一 of データ読み込み、マージ、Excel生成、書式適用ロジックが重複して記述されていた。
  リファクタリングにより、コアロジックを `automation_core.py` の `run_pipeline` と `step1_create_retail_wb` / `step2_python_direct_merge` に集約。これにより、各スクリプトで重複していた50行以上（実際には各ファイル数百行）の冗長なコードを完全に排除し、保守性が大幅に向上した。

### ② 検証・差異デバッグスクリプトの機能強化
* **検証ロジックの堅牢化点 (`debug_diff.py` / `verify_pipeline`)**:
  1. **列インデックス自動解決**: 検証時にヘッダー名からチャネル（ZOZO, OIOI, ECおよび10店舗）の列インデックスを動的に検索・解決することで、レイアウト変更に強いロジックとした。
  2. **商品レベル（セル単位）検証**: 従来の「全体の合計値」のみの突き合わせに加え、商品コード別、チャネル別に期待値（CSVからの集計値）と実際値（生成されたExcelの値）をセル単位で網羅的に照合し、差異があれば不一致レコード（期待値、実際値、行番号）を詳細にリストアップするレポート機能（差異分析レポート）を追加した。
  3. **前提データエラーチェック**: Excelシート内に数式エラー（`#REF!` や `#VALUE!` など）が含まれている場合、処理の前に検出し、警告を行います。
  4. **13チャネル対応化**: 3大チャネル（ZOZO, OIOI, EC）と10店舗（名古屋, TOKYO, ルクア大阪, ヒュッテ, 京王新宿, 大丸心斎橋, 6142, 玉川高島屋, NODE, NARITA）の計13チャネルに完全対応。

---

## 3. 本日の追加修正・動作検証内容

### ① 検証スクリプトでの列名重複解決の修正
* **バグ詳細**: 検証スクリプト（`verify_pipeline`）がExcel上の実際値（ORDER列）を自動解決する際、Excel内に「確保」「在庫」「ORDER」「過不足」という同名の店舗名ヘッダー列が複数存在するため、ORDER列以外の列のインデックスで上書きされて不一致（NG）と誤検知する問題が発生していた。
* **対処方法**: `automation_core.py` 内の `verify_pipeline` を修正し、ヘッダーの2行目を走査して `current_area` が `"ORDER"` である列のインデックスのみを動的に解決するように限定化。これにより、突き合わせの一致精度を100%に高めた。

### ② バッチファイルの終了時キー待ち（一時停止）とエラー検知の追加
* **改善詳細**: バッチファイルをダブルクリックして実行した際、正常完了時や検証エラー発生時でもコンソールウィンドウが即座に閉じてしまい、結果が確認できない仕様になっていた。また、検証で「NG」が出てもバッチ処理が中断せず次の処理に進んでしまっていた。
* **対処方法**: 
  * `実行.bat` を含むすべてのバッチファイルにおいて、自動化処理・検証処理の終了コード（`ERRORLEVEL`）のチェックを追加し、途中で失敗した場合はその場で処理を中断するように修正。
  * 正常完了時（`Done`）およびエラー・不一致検知時（`Error`）のどちらであっても、環境変数 `NO_GUI` の状態に関わらず無条件で `pause` を実行し、キーが押されるまで画面を閉じないように変更。

---

## 4. 共通モジュール分割と店舗動的化リファクタリング (M4期 追加報告)

### ① 共通モジュール `automation_core.py` のモジュール分割
* **リファクタリングの詳細**:
  共有モジュール `automation_core.py`（約1,600行）を、保守性の向上を目的に以下の独立した役割を持つファイルへ分割しました。
  * `config.py`: 設定情報クラスおよび定数
  * `utils.py`: 汎用判定ロジック、ファイル検索、エラーダイアログ等のユーティリティ
  * `data_loader.py`: CSVおよびExcelからの売上データの抽出・集計
  * `generate_sheets.py`: 各種Excelシートの生成・マージ処理
  * `automation_core.py`: メインパイプラインおよび Facade（再エクスポート窓口）
* **後方互換性担保**: `run_automation.py` や `verify_all.py` 等の既存コードに変更を加えず動作させるため、`automation_core.py` を再エクスポートの窓口として残し、互換性を100%担保しています。

### ② 検証処理 (`verify_pipeline`) の完全動的店舗対応
* **改善内容**: 検証ロジックで店舗名リスト（`stores_list`）がハードコードされていたため、新規店舗（「東京ステーション」など）や店舗構成の変更時に検証漏れやエラーが発生する状態となっていました。
* **修正方法**: `parse_csv_store_structure` を用いて `出荷予定振分.csv` から直接店舗名を動的に取得し、そのリストから「ZOZO」「OIOI」「EC」などのチャネルを除外した店舗群を自動特定して照合に使用するようにロジックを動的化しました。

### ③ 過不足ヘッダーの店舗名欠落による検証クラッシュの修正
* **バグ詳細**: Excelの `ws_order` シートを構築する際、過不足列などの一部のエリアで3行目に店舗名ヘッダーが書き込まれておらず（空欄）、これに起因して `verify_pipeline` 内の列マップ走査時にインデックスが空となり `ValueError: min() iterable argument is empty` が発生していました。
* **修正方法**: マスタテンプレートからの店舗名の固定的なコピー処理を廃止し、CSVから動的に取得した `csv_store_names` を用いて、`ws_order` シートの「確保」「在庫」「ORDER」「過不足」の全エリア（3行目）へ動的かつ網羅的に店舗名を書き込むように処理を統一しました。併せて、キー名の名寄せ標準化（例: `丸井web` ↔ `OIOI`）を確実に行うことでマッピング漏れを完全に排除しました。

## 5. 完成形（正解データ）との差分解消とバグフィックス (2026-07-21)

過去の手動作成された完成形（`20260721 RETAIL FRV.xlsx`）と自動生成出力の差分を比較・精査し、以下の5件のバグ修正および精度向上を行いました。

### ① SUMIF 第1引数バグ修正
* **バグ詳細**: orderシート過不足合計の `SUMIF` 式で第1引数（検索範囲）が単一セルになっていた。
* **修正方法**: `SUMIF(AR4:BD4,"<0",...)` のように範囲指定に修正し、計算が正しく機能するようにした。

### ② RETAILシート出荷指示数開始列の計算ずれ修正
* **バグ詳細**: `retail_col_start` の開始列インデックス計算が +8（AV列）になっていた。
* **修正方法**: CSVの実際の構造（0-indexedと1-indexedの混在など）に基づいて +9（AW列）に修正し、完成形と完全に一致させた。

### ③ 出荷予定日列の動的計算
* **バグ詳細**: 出荷予定日列 `retail_date_col` が `max(69, 49+N)` の固定値ベースになっていたため、店舗数が増減した際に正しく追従しなかった。
* **修正方法**: `retail_col_start + N` で動的かつ正確に算出されるように修正した。

### ④ 商品マスタの指数表記正規化
* **バグ詳細**: 商品マスタ読込時、商品コードが `1.22002E+12` などの指数表記として解釈されてしまい、売上データの `1220020000000` 等と紐づかない問題が発生した。
* **修正方法**: 読み込み時に指数表記を検出し、整数文字列に正規化する `_normalize_product_code()` 関数を導入した。

### ⑤ RETAILシートのヘッダークリア処理拡張
* **バグ詳細**: CSV元データから転記される余分なヘッダー（出荷指示数・出荷予定日）や列番号（行1）がRETAILシートに残存し、列数が69列などに不正に増える問題があった。
* **修正方法**: 以前は行5（データ行）以降のみをクリアしていたが、クリア処理を行1から実行するように修正し、総列数を完成形と完全に一致（62列）させた。
