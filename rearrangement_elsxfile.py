import pandas as pd
import os

# 変換したいファイルがあるフォルダのパス
#ダウンロードしてから使用する場合、ここを編集してください！！
folder_path = ''

# xlsxファイルの数をカウントする変数
xlsx_file_count = 0

# 指定されたフォルダ内の全ての.xlsxファイルを探索
for filename in os.listdir(folder_path):
    if filename.endswith(".xlsx"):
        # xlsxファイルのカウントを増やす
        xlsx_file_count += 1

        # Excelファイルのフルパス
        file_path = os.path.join(folder_path, filename)

        # Excelファイルを読み込む際に、engineを明示的に指定する
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # A列のデータをスペースで区切って新しい列に分割
        split_columns = df.iloc[:, 0].str.split(' ', expand=True)
        
        # 分割されたデータを元のデータフレームに結合
        for i, column in enumerate(split_columns.columns):
            df[f'New_Column{i+1}'] = split_columns[column]
        
        # 変更を同じファイルに保存
        df.to_excel(file_path, index=False, engine='openpyxl')

# 処理が終わったら、xlsxファイルの総数を出力
print(f"処理された.xlsxファイルの総数: {xlsx_file_count}")