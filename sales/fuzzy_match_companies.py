import pandas as pd
from thefuzz import process, fuzz
import re
import os
import datetime

def clean_name(name):
    """会社名をクリーニング"""
    if not isinstance(name, str):
        return ''
    name = re.sub(r'株式会社|有限会社|合同会社|\(株\)|\(有\)|\(同\)|㈲|㈱', '', name)
    name = re.sub(r'\s+|　+', '', name)
    name = name.upper()
    return name

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define file paths relative to the script directory
tenkura_master_path = os.path.join(script_dir, 'tenkura_Master_all.csv')
freee_master_path = os.path.join(script_dir, 'freee_master.csv')

# Function to read CSV with encoding fallback
def read_csv_with_fallback(file_path):
    encodings = ['cp932', 'shift_jis', 'utf-8'] # Prioritize Japanese encodings
    for encoding in encodings:
        try:
            df = pd.read_csv(file_path, encoding=encoding)
            return df
        except UnicodeDecodeError:
            continue
        except Exception as e:
            raise
    raise Exception(f"Could not read {file_path} with any of the tested encodings: {encodings}")

# データの読み込み
try:
    tenkura_df = read_csv_with_fallback(tenkura_master_path)
    freee_df = read_csv_with_fallback(freee_master_path)
except Exception as e:
    print(f"ファイル読み込み中にエラーが発生しました: {e}")
    exit()


# 名前のクリーニング
tenkura_df['cleaned_name'] = tenkura_df['取引先名'].apply(clean_name)
freee_df['cleaned_name'] = freee_df['名前（通称）'].apply(clean_name)

# マッチング処理
matches = []
unmatched = []
freee_names = freee_df['cleaned_name'].dropna().unique()
SCORE_CUTOFF = 80 # 類似度のしきい値 (0-100)。80くらいがおすすめです。

for index, row in tenkura_df.iterrows():
    tenkura_name = row['cleaned_name']
    original_tenkura_name = row['取引先名']
    
    if not tenkura_name:
        continue

    # 最も似ているものをfreeeマスタから探す
    best_match = process.extractOne(tenkura_name, freee_names, scorer=fuzz.token_sort_ratio, score_cutoff=SCORE_CUTOFF)

    if best_match:
        matched_freee_name_cleaned = best_match[0]
        score = best_match[1]
        original_freee_name = freee_df[freee_df['cleaned_name'] == matched_freee_name_cleaned]['名前（通称）'].iloc[0]
        matches.append([original_tenkura_name, original_freee_name, score])
    else:
        unmatched.append([original_tenkura_name])

# 結果をDataFrameに変換してCSVに出力
matched_df = pd.DataFrame(matches, columns=['Tenkura Name', 'Matched Freee Name', 'Match Score'])
unmatched_df = pd.DataFrame(unmatched, columns=['Unmatched Tenkura Name'])

matched_output_filename = f"fuzzy_matched_results.csv" # Removed YYYYMM suffix
unmatched_output_filename = f"fuzzy_unmatched_results.csv" # Removed YYYYMM suffix

matched_df.to_csv(os.path.join(script_dir, matched_output_filename), index=False, encoding='utf-8')
unmatched_df.to_csv(os.path.join(script_dir, unmatched_output_filename), index=False, encoding='utf-8')

print("ファジーマッチングが完了しました。")
print(f"マッチした件数: {len(matched_df)}")
print(f"マッチしなかった件数: {len(unmatched_df)}")
