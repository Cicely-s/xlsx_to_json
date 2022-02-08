"""xlsxファイルを入力として、jsonファイルを出力します"""
import os
import openpyxl
import tkinter
from tkinter import messagebox, filedialog
import json
from datetime import date, datetime

# date, datetime の変換関数
def json_serial(obj):
    """日付型をJSONで文字列として認識できるようにする関数

    ・日付型の場合には、文字列に変換します
    ・それ以外はエラーを返します
    """
    if isinstance(obj, datetime):
        return obj.isoformat()


# ファイル選択ダイアログ
root = tkinter.Tk()
root.withdraw()

# ファイルタイプ設定
ftype = [("", "*.xlsx")]

dir = 'D:\\2022\\'
messagebox.showinfo('xlsx to json プログラム', '処理ファイルを選択してください')
file = filedialog.askopenfilename(filetypes = ftype, initialdir = dir)

# xlsx -> json 変換
book = openpyxl.load_workbook(file)
for ws in book:
    sheet_name = ws.title
    vals = [[cell.value for cell in row] for row in ws]
    keys = vals[0]
    # ワークシートごとにjsonファイルに書き込む内容を初期化する
    dic_list = []

    for i, row in enumerate(vals[1:]):
        # fields の中身をrowごとに初期化する
        fields = {}
        for j, col in enumerate(row):
            fields[keys[j]] = col
        data_dic = {
            "models":sheet_name,
            "pk": i+1,
            "fields": fields
            }
        dic_list.append(data_dic)
    # シートごとにjsonとして出力
    with open(os.path.join(os.path.dirname(file), sheet_name+'.json'), 'w', encoding='utf-8') as f:
        json.dump(dic_list, f, indent=4, ensure_ascii=False, default=json_serial)
