import pandas as pd
import os

# --- 設定 ---
# 変換元ファイルがあるディレクトリ
SOURCE_DIRECTORY = '.'
# 変換後のファイルを保存するディレクトリ
OUTPUT_DIRECTORY = '.' # 同じ場所に保存
# ----------------

def convert_excel_to_csv(filepath):
    """
    Excel形式のファイルを読み込み、UTF-8形式のCSVとして保存する
    """
    basename = os.path.basename(filepath)
    filename_no_ext = os.path.splitext(basename)[0]
    output_csv_path = os.path.join(OUTPUT_DIRECTORY, f"{filename_no_ext}_converted.csv")

    print(f"--- 変換処理中: {basename} ---")

    try:
        # Excelファイルとして読み込み
        df = pd.read_excel(filepath, engine='openpyxl')
        
        # UTF-8 (BOM付き) のCSVとして保存
        # index=False で行番号を保存しない
        df.to_csv(output_csv_path, index=False, encoding='utf-8-sig')
        
        print(f"✓ 変換完了: {output_csv_path}")
        return True

    except Exception as e:
        # もしファイルが本物のCSVだった場合など
        print(f"[警告] Excelとしての読み込みに失敗しました。スキップします。理由: {e}")
        return False

if __name__ == '__main__':
    # ディレクトリにあるExcelファイル (.xlsx) を取得
    source_files = [f for f in os.listdir(SOURCE_DIRECTORY) if f.endswith('.xlsx')]
    
    if not source_files:
        print("[情報] 変換対象のファイルが見つかりません。")
    else:
        print(f"{len(source_files)}個のファイルを変換します...")
        converted_count = 0
        for filename in source_files:
            filepath = os.path.join(SOURCE_DIRECTORY, filename)
            if convert_excel_to_csv(filepath):
                converted_count += 1
        print(f"\n--- {converted_count}個のファイルの変換が完了しました ---")
