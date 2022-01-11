import os
import platform
from datetime import datetime as dt

import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


def ReadConditionCSV(file_path):
    with open(file_path, mode="r", encoding="utf8") as f:
        temp_data_list = f.readlines()

    data_list = [data.strip().split(",") for data in temp_data_list[3:]]
    date_pos = 0
    year_set = set()
    data_dict: dict[dt, list[str]] = {}
    for data in data_list:
        dt_val = dt.strptime(data[date_pos], r"%Y/%m/%d")
        year_set.add(dt_val.strftime("%Y"))
        data_dict[dt_val] = data[date_pos + 1 :]
    return data_dict, year_set


def Input2Excel(file_path, data_dict, year_set):
    xlsx_dir = os.path.dirname(file_path)
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_list = wb.sheetnames
    ws_header_cnt = 2

    for sheet_name in sheet_list:
        if sheet_name not in year_set:
            continue

        ws: Worksheet = wb[sheet_name]
        for dt_cell, condition_cell, comment_cell in zip(
            ws["A:A"][ws_header_cnt:],
            ws["C:C"][ws_header_cnt:],
            ws["D:D"][ws_header_cnt:],
        ):
            if data_dict.get(dt_cell.value) is None:
                continue

            data_list = data_dict[dt_cell.value]
            condition_str = data_list[0]
            condition_cell.value = (
                int(condition_str) if condition_str.isdigit() else condition_str
            )
            if len(data_list) > 1:
                comment_cell.value = data_list[1]

    wb.save(os.path.join(xlsx_dir, "test.xlsx"))


if __name__ == "__main__":
    # iOSで動いているかの判定
    is_iOS = False
    if "iPhone" in platform.platform() or "iPad" in platform.platform():
        is_iOS = True
        from FilePickerPyto import FilePickerPyto

        csv_path = FilePickerPyto(
            file_types=["public.data"], allows_multiple_selection=False
        )[0]
        xlsx_path = FilePickerPyto(
            file_types=["public.data"], allows_multiple_selection=False
        )[0]
    else:
        from FilePicker import GetFilePathByGUI

        csv_path = GetFilePathByGUI(
            file_type=(["csvファイル", "*.csv"],),
        )[0]
        xlsx_path = GetFilePathByGUI(
            file_type=(["xlsxファイル", "*.xlsx"],),
        )[0]

    condition_data_dict, year_set = ReadConditionCSV(csv_path)
    Input2Excel(xlsx_path, condition_data_dict, year_set)
