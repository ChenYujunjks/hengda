import openpyxl

# 打开Excel文件
wb = openpyxl.load_workbook('target.xlsx')

file_delete="try4"

# 检查工作表是否存在
if file_delete in wb.sheetnames:
    # 删除工作表
    ws = wb[file_delete]
    wb.remove(ws)
    # 保存更改
    wb.save('target.xlsx')
    print("工作表 {} 已被删除。".format(file_delete))
else:
    print("工作表 {}不存在。".format(file_delete))
