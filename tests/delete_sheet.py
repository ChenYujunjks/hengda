import openpyxl

# 打开Excel文件
wb = openpyxl.load_workbook('target.xlsx')

# 检查工作表是否存在
if 'try4' in wb.sheetnames:
    # 删除工作表
    ws = wb['try4']
    wb.remove(ws)
    # 保存更改
    wb.save('target.xlsx')
    print("工作表 'try3' 已被删除。")
else:
    print("工作表 'try3' 不存在。")
