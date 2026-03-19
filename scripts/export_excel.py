import sys
import json
import os
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy

def copy_style(src_cell, dst_cell):
    """セルスタイルをコピーするヘルパー"""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)

def export_excel(payload_path, template_path, output_path):
    # JSONペイロードを読み込み
    with open(payload_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # テンプレートを読み込み
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    # --- レイアウトの動的検出 ---
    row_facility = 3
    col_facility = 9 # I
    col_date = 16 # P
    row_header = 13
    row_total = 14
    row_example = 15
    row_data_start = 16
    
    # セルの内容から各行・列を特定
    for r in range(1, 21):
        for c in range(1, 26):
            cell_val = str(ws.cell(row=r, column=c).value or "")
            if "施設名" in cell_val:
                row_facility = r
                col_facility = c
            if "施術日" in cell_val:
                col_date = c
            if "No." in cell_val:
                row_header = r
            if "合計人数" in cell_val:
                row_total = r
            if "記入例" in cell_val:
                row_example = r
                row_data_start = r + 1

    # --- ヘッダー書き込み ---
    # 施設名 (施設名：XXX)
    ws.cell(row=row_facility, column=col_facility).value = f"施設名：{data.get('facilityName', '')}"
    
    # 施術日
    year = data.get('year')
    month = data.get('month')
    day = data.get('day')
    if year and month and day:
        try:
            reiwa = int(year) - 2018
            # 形式: 施術日：令和 8 年 3 月 19 日
            date_str = f"施術日：令和 {reiwa} 年 {month} 月 {day} 日"
            ws.cell(row=row_facility, column=col_date).value = date_str
        except:
            ws.cell(row=row_facility, column=col_date).value = f"施術日：令和    年   月   日"
    else:
        ws.cell(row=row_facility, column=col_date).value = f"施術日：令和    年   月   日"

    # --- メニュー列の特定 ---
    menu_cols = {} # {menu_name: column_index}
    # ヘッダー行 (row_header-1 〜 row_header) でメニュー名を探す
    for r in [row_header-1, row_header]:
        for c in range(7, 14): # G-M列あたり
            v = str(ws.cell(row=r, column=c).value or "")
            if v and v not in ["No.", "部屋番号", "氏名", "性別", "合計料金", "施術開始時間の希望", "合計", "料金"]:
                # 「カット」などの名前が含まれているか。改行を除く。
                clean_v = v.replace("\n", "").split(" ")[0].split("(")[0].strip()
                if clean_v:
                    menu_cols[clean_v] = c

    # --- 顧客データ書き込み ---
    customers = data.get('customers', [])
    for idx, customer in enumerate(customers):
        row_num = row_data_start + idx
        
        # B: No
        ws.cell(row=row_num, column=2).value = customer.get('no', idx + 1)
        # C: 部屋番号
        ws.cell(row=row_num, column=3).value = customer.get('room', '')
        # D: 氏名
        ws.cell(row=row_num, column=4).value = customer.get('name', '')
        # F: 性別
        ws.cell(row=row_num, column=6).value = customer.get('gender', '')

        # スタイルを「記入例」行からコピー
        for c in range(2, 21):
            copy_style(ws.cell(row=row_example, column=c), ws.cell(row=row_num, column=c))

        # メニュー選択 (〇)
        selected_menus = customer.get('selectedMenus', [])
        for m_name in selected_menus:
            # 前方一致などでマッチング
            match_col = None
            for col_name, col_idx in menu_cols.items():
                if col_name in m_name or m_name in col_name:
                    match_col = col_idx
                    break
            if match_col:
                ws.cell(row=row_num, column=match_col).value = "〇"

        # L: 合計
        # M-O: 希望時間
        times = customer.get('preferredTimes', [])
        for i in range(min(len(times), 3)):
            ws.cell(row=row_num, column=13 + i).value = times[i]

        # P: ご案内有無, Q: 施術実施有無, R: 追加メニュー可否, S: オーダーメイド, T: 備考
        ws.cell(row=row_num, column=16).value = customer.get('isGuided', '')
        if customer.get('hasService'):
            ws.cell(row=row_num, column=17).value = "サイン" # テンプレートに合わせる
        ws.cell(row=row_num, column=18).value = customer.get('isAdditionalMenuAllowed', '可・否')
        ws.cell(row=row_num, column=19).value = customer.get('isCustomOrder', '本人・お任せ')
        ws.cell(row=row_num, column=20).value = customer.get('remarks', '')

    # --- 合計人数行の書き込み (Row 14 or 15) ---
    # 通常 D14/D15 に人数の合計
    if row_total:
        ws.cell(row=row_total, column=4).value = len(customers)
        # D列以外 (F-M列など) に 0 または合計値を設定
        for c in range(6, 14):
            if ws.cell(row=row_total, column=c).value is not None:
                # 〇の数をカウントする数式を入れる (例: =COUNTIF(F16:F100, "〇"))
                col_ltr = get_column_letter(c)
                ws.cell(row=row_total, column=c).value = f"=COUNTIF({col_ltr}{row_data_start}:{col_ltr}{row_data_start + len(customers) + 10}, \"〇\")"

    # --- シートのクリーンアップ ---
    for sheetname in wb.sheetnames:
        if sheetname not in ["申込書", ws.title]:
             del wb[sheetname]
    
    wb.save(output_path)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        sys.exit(1)
    try:
        export_excel(sys.argv[1], sys.argv[2], sys.argv[3])
        print("SUCCESS")
    except Exception as e:
        print(f"ERROR: {str(e)}", file=sys.stderr)
        sys.exit(1)
