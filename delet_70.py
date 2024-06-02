from openpyxl import load_workbook, Workbook
import os

source_file = "origin/source.xlsx"
new_file = "source.xlsx" 

# 检查文件是否存在
if not os.path.exists(source_file):
    print(f"文件 {source_file} 不存在，请检查文件路径")
else:
    try:
        wb = load_workbook(source_file)
        ws = wb.active

        # 创建一个新的工作簿
        new_wb = Workbook()
        new_ws = new_wb.active

        # 复制前 250 行的数据到新的工作簿
        for row in ws.iter_rows(min_row=1, max_row=250):
            for cell in row:
                new_ws[cell.coordinate] = cell.value

        # 保存新的 Excel 文件到根目录
        new_wb.save(new_file)
        print(f"前 250 行的数据已保存到 {new_file}")
    except Exception as e:
        print(f"处理文件时出现错误: {e}")
