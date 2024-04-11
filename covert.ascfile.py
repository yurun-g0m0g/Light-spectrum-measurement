import pandas as pd
import os

# 変換したいファイルがあるフォルダのパス
#ダウンロードしてから使用する場合、ここを編集してください！！
folder_path = ''

# 変換されたファイル数をカウントする変数
converted_files_count = 0

# 指定されたフォルダ内の全ての.ascファイルを探索
for filename in os.listdir(folder_path):
    if filename.endswith(".asc"):
        # ファイルパスの作成
        file_path = os.path.join(folder_path, filename)
        
        # .ascファイルを読み込み
        data = pd.read_csv(file_path, delimiter='\t') # ここで、データの区切り文字を指定します。タブが標準ですが、必要に応じて変更してください。
        
        # Excelファイル名の作成（.ascを.xlsxに変更）
        excel_filename = filename.replace('.asc', '.xlsx')
        excel_path = os.path.join(folder_path, excel_filename)
        
        # Excelファイルとして保存
        data.to_excel(excel_path, index=False)

        # 変換されたファイル数をインクリメント
        converted_files_count += 1

# 変換されたファイル数の表示
print(f"全ての.ascファイルが.xlsxに変換されました。変換されたファイル数: {converted_files_count}")