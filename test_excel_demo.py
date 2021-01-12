#ecoding=utf-8
# author:herui
# time:2021/1/11 11:35
# function:
import xlrd


def test_sheet_by_sheets():
    excel = xlrd.open_workbook(r'test.xlsx',encoding_override="utf-8")
    all_sheet = excel.sheets()
    sheet_name = []
    for sheet in all_sheet:
        sheet_name.append(sheet.name)
    print(sheet_name)

def test_sheet_by_index():
    excel = xlrd.open_workbook(r'test.xlsx',encoding_override="utf-8")
    all_sheet = excel.sheets()
    sheet_name = []
    for i in range(len(all_sheet)):
        sheet_name.append(excel.sheet_by_index(i).name)
    print(sheet_name)

def test_get_sheet_info():
    excel = xlrd.open_workbook(r'test.xlsx', encoding_override="utf-8")
    # 获取sheet对象
    all_sheet = excel.sheets()

    sheet_name = []
    sheet_row = []
    sheet_col = []
    for sheet in all_sheet:
        sheet_name.append(sheet.name)
        sheet_row.append(sheet.nrows)
        sheet_col.append(sheet.ncols)
        print(f"该文件共有{len(all_sheet)}个sheet")
        print(f"当前sheet为:{sheet.name},该sheet有:{sheet.nrows}行，{sheet.ncols}列")

        # 按行打印数据
        # for n in range(sheet.nrows):
        #     row_info = sheet.row_values(n)
        #     print(f"当前为第{n}行：{row_info}")

        # 按列打印数据
        # for n in range(sheet.ncols):
        #     col_info = sheet.col_values(n)
        #     print(f"当前为第{n}列：{col_info}")

        # 打印指定单元格数据

def get_more_info():
    excel = xlrd.open_workbook(r'test.xlsx', encoding_override="utf-8")
    # 获取sheet对象
    all_sheet = excel.sheets()

    sheet_name = []
    sheet_row = []
    sheet_col = []
    for sheet in all_sheet:
        # 获取对象
        sheet_cell = sheet.cell(0,0)
        # 获取值
        sheet_value = sheet_cell.value
        print(sheet_cell, sheet_value)

        sheet_cell_value = sheet.cell_value(0,0)
        print(sheet_cell_value)

