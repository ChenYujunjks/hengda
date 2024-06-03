import os
import openpyxl
from data_processing import process_data

# 构建相对路径
source_file = os.path.join('..', 'test', 'source.xlsx')
target_file = os.path.join('..', 'test', 'target.xlsx')
new_sheet_name = 'try250'

def modify_excel_data(source_file, target_file, new_sheet_name):
    source_wb = openpyxl.load_workbook(source_file)
    source_sheet = source_wb.active  # 假设源文件中需要复制的数据在第一个工作表
    # 打开目标文件（如果不存在则创建一个新的文件）
    try:
        target_wb = openpyxl.load_workbook(target_file)
    except FileNotFoundError:
        target_wb = openpyxl.Workbook()

# 在目标文件中创建一个新的工作表
    if new_sheet_name in target_wb.sheetnames:
        target_sheet = target_wb[new_sheet_name]
    else:
        target_sheet = target_wb.create_sheet(title=new_sheet_name)

    process_data(source_sheet, target_sheet)
    target_wb.save(target_file)

if __name__ == "__main__":
    modify_excel_data(source_file, target_file, new_sheet_name)
