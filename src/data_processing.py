def second_star_index(text):
    first_star_index = text.find('*')
    if (first_star_index == -1):
        return -1  # 如果没有找到第一个星号，返回-1
    
    second_star_index = text.find('*', first_star_index + 1)
    return second_star_index

def process_data(source_sheet, target_sheet):
    for row_idx, row in enumerate(source_sheet.iter_rows(values_only=True), start=1):
        if row_idx == 1:
            # 处理标题行
            new_row = ["序号", "生产编号", row[5], row[2], "客户名称", "货物名称", "备注软件名称", "数量", "不含税金额", "税额", "合计"]
            target_sheet.append(new_row)
        else:
            # 数据行 条件筛选
            try:
                SUM_IJ = float(row[9]) + float(row[11])
            except ValueError:
                SUM_IJ = None
                continue
            if "含" in row[17]:  #备注软件名称
                index1 = row[17].find("含")
                bzrjmc = row[17][index1:]
                row6index = second_star_index(row[6])   # 货物名称 中第二个星号的index
                new_6row = row[6][row6index + 1:]
                # 开始写 new row
                new_row = [
                    row_idx - 1, 
                    0,
                    row[5], 
                    row[2], 
                    row[4], 
                    new_6row, 
                    bzrjmc, 
                    1,
                    row[9], 
                    row[11], 
                    SUM_IJ
                ]
                target_sheet.append(new_row)
