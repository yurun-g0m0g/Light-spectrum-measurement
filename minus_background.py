import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# 統合されたExcelファイルのパス
integrated_file_path = '/Users/sidareyanagi542/Desktop/授業資料/4年/研究室/実験/0405/統合ファイル.xlsx'

# 参照するExcelファイルのパス
reference_file_path = '/Users/sidareyanagi542/Desktop/授業資料/4年/研究室/実験/0405/background.xlsx'

# 参照するExcelファイルを読み込む
ref_df = pd.read_excel(reference_file_path, engine='openpyxl')

# 統合されたExcelファイルを開く
wb = load_workbook(integrated_file_path)

# 統合されたExcelファイルの各シートを処理
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    # C-D列の29行目から56行目の各セルについて処理
    for row in range(29, 57):  # 29行目から56行目
        for col in ['C', 'D']:  # C-D列
            cell = ws[f"{col}{row}"]
            # 参照するセルの値を取得 (参照DataFrameから値を取得)
            ref_value = ref_df.loc[row-29, col].item() if not pd.isna(ref_df.loc[row-29, col]) else 0

            try:
                # セルの値を数値に変換し、参照セルの値を引く
                if cell.value is not None and cell.value != '':
                    cell_value = float(cell.value)
                    cell.value = cell_value - ref_value
            except ValueError as e:
                print(f"変換エラー: シート '{sheet_name}', セル {col}{row}, エラー: {e}")

# 変更を保存
wb.save(integrated_file_path)

print(f"{integrated_file_path} の各シートのC-D列の29行目から56行目のセルが更新されました。")