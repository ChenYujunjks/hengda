import openpyxl

def second_star_index(text):
    first_star_index = text.find('*')
    if first_star_index == -1:
        return -1  # 如果没有找到第一个星号，返回-1
    
    second_star_index = text.find('*', first_star_index + 1)
    return second_star_index

def get_scbh_from_shengchan(shengchan_file, fphm): #获取生产编号
    print(f"Reading {shengchan_file} to find SCBH for FPHM: {fphm}")
    shengchan_wb = openpyxl.load_workbook(shengchan_file)
    shengchan_sheet = shengchan_wb.active  # 假设生产文件中的数据在第一个工作表
    
    for row in shengchan_sheet.iter_rows(min_row=2, values_only=True):  # 假设第1行为标题行，从第2行开始
        if row[4] == fphm:  # E列是第五列（索引从0开始）
            print(f"Found SCBH: {row[1]} for FPHM: {fphm}")
            return row[1]  # B列是第二列（索引从0开始）
    return None