import openpyxl

# 打开Excel文件
wb = openpyxl.load_workbook('target.xlsx')

# 检查工作表是否存在
if 'try2' in wb.sheetnames:
    # 删除工作表
    ws = wb['try2']
    wb.remove(ws)
    # 保存更改
    wb.save('target.xlsx')
    print("工作表 'try2' 已被删除。")
else:
    print("工作表 'try2' 不存在。")
