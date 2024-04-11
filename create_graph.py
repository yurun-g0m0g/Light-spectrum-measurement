from openpyxl import load_workbook
from openpyxl.chart import ScatterChart, Reference, Series

# 統合されたExcelファイルのパス
integrated_file_path = '/Users/sidareyanagi542/Desktop/授業資料/4年/研究室/実験/0405/統合ファイル.xlsx'

# 統合されたExcelファイルを開く
wb = load_workbook(integrated_file_path)

# 統合されたExcelファイルの各シートを処理
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    # グラフのデータ範囲を指定（29行目から最終行までのD列とE列）
    min_row = 29
    max_row = ws.max_row
    x_values = Reference(ws, min_col=5, min_row=min_row, max_row=max_row)  # E列
    y_values = Reference(ws, min_col=4, min_row=min_row, max_row=max_row)  # D列

    # 散布図を作成
    chart = ScatterChart()
    chart.title = f"Scatter Chart for {sheet_name}"
    chart.x_axis.title = 'E column values'
    chart.y_axis.title = 'D column values'
    chart.legend = None  # 凡例は不要

    # データ系列を追加（点と点を線で繋がない設定）
    series = Series(y_values, x_values, title_from_data=False)
    series.marker.symbol = "circle"  # 点を丸で表示
    series.graphicalProperties.line.noFill = True  # 線を表示しない
    chart.series.append(series)

    # グラフをシートに追加
    ws.add_chart(chart, "G10")  # グラフを配置するセルを指定（例: G10）

# 変更を保存
wb.save(integrated_file_path)

print(f"{integrated_file_path} の各シートに散布図が追加されました。")
