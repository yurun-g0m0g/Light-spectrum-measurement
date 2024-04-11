import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows  

# 変換したいファイルがあるフォルダのパス
#ダウンロードしてから使用する場合、ここを編集してください！！
folder_path = ''

# 統合されたシートを保存する新しいExcelファイルのパス
#ダウンロードしてから使用する場合、ここを編集してください！！
output_file_path = ''

# シート名とデータフレームを格納する一時的な辞書
sheets_dict = {}

# 指定されたフォルダ内の全ての.xlsxファイルを探索
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx"):
        # Excelファイルのフルパス
        file_path = os.path.join(folder_path, filename)
        
        # Excelファイル内の全シートを読み込み
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name)
            # シート名とデータフレームを辞書に追加
            sheets_dict[f"{filename}_{sheet_name}"] = df

# 新しいワークブックを作成
wb = Workbook()
wb.remove(wb.active)  # デフォルトで作成される空のシートを削除

# シート名でソートしてデータを新しいワークブックに書き込む
for sheet_name in sorted(sheets_dict.keys()):
    ws = wb.create_sheet(title=sheet_name)
    df = sheets_dict[sheet_name]
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

# 新しいExcelファイルに保存
wb.save(output_file_path)

print(f"全てのシートが {output_file_path} に統合され、あいうえお順に並び替えられました。")