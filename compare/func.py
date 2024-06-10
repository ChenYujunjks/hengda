
def compute_sum_diff(sheet_output,fill1,fill2,i,output_row):
        # 计算 I 列蓝色的总和减去橘色的总和
    sum_1 = 0
    sum_2 = 0
    for row in range(2, output_row):
        cell = sheet_output.cell(row=row, column=9)  # I 列是第9列
        if cell.fill == fill1:
            sum_1 += cell.value if isinstance(cell.value, (int, float)) else 0
        elif cell.fill == fill2:
            sum_2 += cell.value if isinstance(cell.value, (int, float)) else 0

    difference = sum_1 - sum_2

    # 在 i 列的最后一行输出结果，在 i-1 列写入 "差值:"
    sheet_output.cell(row=output_row, column=i-1, value="差值:")  # H 列是第8列
    sheet_output.cell(row=output_row, column=i, value=difference)  # I 列是第9列
    
    return difference