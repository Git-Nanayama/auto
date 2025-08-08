  月次処理指示書

  この指示書は、売上伝票_YYYYMMDDHHMMSS.csvファイルからfreee会計インポート用のExcelファイルfreee_import_
  sales_data_YYYYMM.xlsxを生成する手順を説明します。

  前提条件:
   * Pythonがインストールされており、pandas, python-levenshtein,
     thefuzzライブラリがインストールされていること。
   * 以下のファイルがC:\Users\kyomirai01\Downloads\auto\KMJ\3. Sales\ディレクトリに存在すること。
       * Account_Master_all.csv (販売管理システムのアカウントマスタ)
       * tenkura_Master_all.csv (販売管理システムのマスタ)
       * freee_master.csv (会計システムのマスタ)
       * Freee_import_format.csv (freee会計のインポートフォーマット)
       * process_sales_data.py
       * fuzzy_match_companies.py
       * convert_to_freee_format.py
       * convert_csv_to_excel.py

  月次処理の手順:

   1. 新しい売上伝票ファイルの配置:
       * 毎月、新しい売上伝票_YYYYMMDDHHMMSS.csvファイルをC:\Users\kyomirai01\Downloads\auto\KMJ\3.
         Sales\ディレクトリに配置します。ファイル名の日付部分は、処理対象の売上伝票ファイルの日付と時刻を
         反映していることを確認してください。

   2. 売上データへの部門情報付与:
       * コマンドプロンプトまたはターミナルを開き、以下のコマンドを実行します。

   1         cd C:\Users\kyomirai01\Downloads\auto\KMJ\3. Sales\
   2         python process_sales_data.py
       * このスクリプトは、最新の売上伝票_YYYYMMDDHHMMSS.csvを読み込み、Account_Master_all.csvに基づいて部
         門情報を付与し、Sales_Data_YYYYMM.csvという名前で保存します。YYYYMMは現在の年月です。

   3. 得意先名のファジーマッチング (初回のみ、またはマスタ更新時):
       * 注意: このステップは、tenkura_Master_all.csvまたはfreee_master.csvが更新された場合、または初回実
         行時のみ必要です。毎月実行する必要はありません。
       * コマンドプロンプトまたはターミナルで、以下のコマンドを実行します。
   1         cd C:\Users\kyomirai01\Downloads\auto\KMJ\3. Sales\
   2         python fuzzy_match_companies.py
       * このスクリプトは、tenkura_Master_all.csvとfreee_master.csvを比較し、fuzzy_matched_results.csvとfu
         zzy_unmatched_results.csvを生成します。

   4. freee会計インポートフォーマットへの変換:
       * コマンドプロンプトまたはターミナルで、以下のコマンドを実行します。
   1         cd C:\Users\kyomirai01\Downloads\auto\KMJ\3. Sales\
   2         python convert_to_freee_format.py
       * このスクリプトは、Sales_Data_YYYYMM.csvとfuzzy_matched_results.csvを読み込み、freee会計インポート
         フォーマットに変換したfreee_import_sales_data_YYYYMM.csvを生成します。

   5. CSVからExcelへの変換:
       * コマンドプロンプトまたはターミナルで、以下のコマンドを実行します。
   1         cd C:\Users\kyomirai01\Downloads\auto\KMJ\3. Sales\
   2         python convert_csv_to_excel.py
       * このスクリプトは、freee_import_sales_data_YYYYMM.csvを読み込み、最終的なExcelファイルfreee_import
         _sales_data_YYYYMM.xlsxを生成します。

  出力ファイル:
   * Sales_Data_YYYYMM.csv: 部門情報が付与された売上データ。
   * fuzzy_matched_results.csv: 天倉マスタとfreeeマスタのファジーマッチング結果。
   * fuzzy_unmatched_results.csv: 天倉マスタとfreeeマスタでマッチしなかった結果。
   * freee_import_sales_data_YYYYMM.csv:
     freee会計インポートフォーマットに変換された売上データ（CSV形式）。
   * freee_import_sales_data_YYYYMM.xlsx:
     freee会計インポートフォーマットに変換された売上データ（Excel形式）。