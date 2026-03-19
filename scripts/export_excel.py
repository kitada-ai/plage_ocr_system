import sys
import json
import os
import openpyxl
from openpyxl.utils import get_column_letter

def export_excel(payload_path, template_path, output_path):
    # JSONペイロードを読み込み
    with open(payload_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # テンプレートを読み込み
    wb = openpyxl.load_workbook(template_path)
    if "申込書" not in wb.sheetnames:
        raise ValueError("申込書シートが見つかりません")
    
    ws = wb["申込書"]

    # 定数集 (1-based index for openpyxl)
    ROW_FACILITY = 3
    COL_FACILITY = 9 # I
    COL_DATE = 16 # P
    ROW_MENU_HEADER = 9
    ROW_PRICE = 10
    MENU_START_COL = 7 # G
    TEMPLATE_MENU_COUNT = 6
    ROW_DATA_START = 12
    COL_NO = 2 # B
    COL_ROOM = 3 # C
    COL_NAME = 4 # D
    COL_GENDER = 6 # F
    COL_TOTAL = 13 # M
    COL_TIME_START = 14 # N (N,O,P)
    COL_SERVICE = 17 # Q
    COL_REMARKS = 18 # R

    # 施設名
    ws.cell(row=ROW_FACILITY, column=COL_FACILITY).value = f"施設名：{data.get('facilityName', '')}"

    # 施術日
    year = data.get('year')
    month = data.get('month')
    day = data.get('day')
    if year and month and day:
        # 令和計算
        try:
            reiwa = int(year) - 2018
            ws.cell(row=ROW_FACILITY, column=COL_DATE).value = f"令和{reiwa}年{month}月{day}日"
        except:
            pass

    # メニューと単価
    menu_items = data.get('menuItems', [])
    for i in range(min(len(menu_items), TEMPLATE_MENU_COUNT)):
        menu = menu_items[i]
        ws.cell(row=ROW_MENU_HEADER, column=MENU_START_COL + i).value = menu.get('name', '')
        ws.cell(row=ROW_PRICE, column=MENU_START_COL + i).value = menu.get('price', 0)

    # 顧客データ
    customers = data.get('customers', [])
    for idx, customer in enumerate(customers):
        row_num = ROW_DATA_START + idx
        
        # 基本情報
        ws.cell(row=row_num, column=COL_NO).value = customer.get('no', idx + 1)
        ws.cell(row=row_num, column=COL_ROOM).value = customer.get('room', '')
        ws.cell(row=row_num, column=COL_NAME).value = customer.get('name', '')
        ws.cell(row=row_num, column=COL_GENDER).value = customer.get('gender', '')

        # メニュー選択 (〇をつける)
        selected_menus = customer.get('selectedMenus', [])
        for i in range(min(len(menu_items), TEMPLATE_MENU_COUNT)):
            menu_name = menu_items[i].get('name', '')
            if menu_name in selected_menus:
                ws.cell(row=row_num, column=MENU_START_COL + i).value = "〇"
            else:
                ws.cell(row=row_num, column=MENU_START_COL + i).value = ""

        # 合計料金の数式
        formula_parts = []
        for i in range(min(len(menu_items), TEMPLATE_MENU_COUNT)):
            col_letter = get_column_letter(MENU_START_COL + i)
            # IF(G12="〇",G10,0)
            formula_parts.append(f'IF({col_letter}{row_num}="〇",{col_letter}{ROW_PRICE},0)')
        
        if formula_parts:
            ws.cell(row=row_num, column=COL_TOTAL).value = "=" + "+".join(formula_parts)

        # 希望時間
        times = customer.get('preferredTimes', [])
        for i in range(min(len(times), 3)):
            ws.cell(row=row_num, column=COL_TIME_START + i).value = times[i]

        # 施術実施
        if customer.get('hasService'):
            ws.cell(row=row_num, column=COL_SERVICE).value = "✓"

        # 備考
        ws.cell(row=row_num, column=COL_REMARKS).value = customer.get('remarks', '')

    # 「申込書」以外のシートを削除 (ユーザーの要望: 請求書ではなく申込書として出力)
    for sheetname in wb.sheetnames:
        if sheetname != "申込書":
            del wb[sheetname]

    # 保存
    wb.save(output_path)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: python export_excel.py payload.json template.xlsx output.xlsx")
        sys.exit(1)
    
    try:
        export_excel(sys.argv[1], sys.argv[2], sys.argv[3])
        print("SUCCESS")
    except Exception as e:
        print(f"ERROR: {str(e)}", file=sys.stderr)
        sys.exit(1)
