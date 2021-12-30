import os
import openpyxl as xl
from glob import glob
from copy import copy


out_name = 'Output.xlsx'
if os.path.exists(out_name):
    os.remove(out_name)

file_paths = sorted(glob('*.xlsx'))
out_wb = xl.Workbook()
out_ws = out_wb.active
out_r, out_c = 1, 1
header = False
for file_path in file_paths:
    print(file_path)
    wb = xl.load_workbook(file_path)
    ws = wb.active

    if not header:
        is_header = True
        for rh in range(1, ws.max_row+1):
            if not is_header:
                break
            for ch in range(1, ws.max_column+1):
                cell = ws.cell(rh, ch)
                out_ws.cell(out_r, out_c).value = cell.value
                if cell.has_style:
                    out_ws.cell(out_r, out_c).fill = copy(cell.fill)
                    out_ws.cell(out_r, out_c).border = copy(cell.border)
                    out_ws.cell(out_r, out_c).alignment = copy(cell.alignment)
                    out_ws.cell(out_r, out_c).font = copy(cell.font)
                    out_ws.cell(out_r, out_c).protection = copy(
                        cell.protection)
                out_c += 1
                if cell.value == "(27)":
                    is_header = False
                    break
            out_r += 1
            out_c = 1
        for mr in list(ws.merged_cells)[1:]:
            out_ws.merge_cells(str(mr))
        header = True

    out_r -= 1
    is_data = False
    for r in range(1, ws.max_row+1):
        for c in range(1, ws.max_column+1):
            cell = ws.cell(r, c)
            if is_data:
                out_ws.cell(out_r, out_c).value = cell.value
                if cell.has_style:
                    out_ws.cell(out_r, out_c).fill = copy(cell.fill)
                    # out_ws.cell(out_r, out_c).border = copy(cell.border)
                    # out_ws.cell(out_r, out_c).alignment = copy(cell.alignment)
                out_c += 1
            if cell.value == "(27)":
                is_data = True
                break

        if is_data:
            out_r += 1
            out_c = 1

out_wb.save(out_name)
print('Done!')
