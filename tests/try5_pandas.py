import openpyxl
import pandas as pd

source_file = 'source.xlsx'
shengchan_file = "help.xlsx"
target_file = 'target.xlsx'
new_sheet_name = 'try4'

def second_star_index(text):
    first_star_index = text.find('*')
    if first_star_index == -1:
        return -1  # 如果没有找到第一个星号，返回-1
    
    second_star_index = text.find('*', first_star_index + 1)
    return second_star_index

def get_scbh_from_shengchan(shengchan_df, fphm): #获取生产编号
    matching_row = shengchan_df.loc[shengchan_df['E列'] == fphm]
    if not matching_row.empty:
        return matching_row.iloc[0]['B列']
    return None

def modify_excel_data(source_file, target_file, new_sheet_name, shengchan_file):
    # 使用 Pandas 读取 Excel 文件
    source_df = pd.read_excel(source_file)
    shengchan_df = pd.read_excel(shengchan_file)

    # 创建目标 DataFrame
    target_df = pd.DataFrame(columns=["序号", "生产编号", '开票日期', "发票号码", "客户名称", "货物名称", "备注软件名称", "数量", "不含税金额", "税额", "合计"])

    index_xh = 1
    for row_idx, row in source_df.iterrows():
        if row_idx == 0:
            continue  # 跳过标题行
        try:
            SUM_IJ = float(row[16]) + float(row[18])
        except ValueError:
            continue
        if "含" in row[26]:   #改动 备注软件名称
            index1 = row[26].find("含")  # 找到转账备注中的 含：恒达富士乘客电梯变频控制软件V1.0的数据
            bzrjmc = row[26][index1:]  
            row6index = second_star_index(row[11])   # 货物名称 中第二个星号的index
            hwmc = row[11][row6index+1:]

            # 获取发票号码对应的生产编号
            fphm = row[3]  # 假设发票号码在第四列
            scbh = get_scbh_from_shengchan(shengchan_df, fphm)

            # 添加新行到目标 DataFrame
            new_row = {
                "序号": index_xh, 
                "生产编号": scbh if scbh else 0,  # 如果找不到生产编号，则填0
                "开票日期": row[8],  # 开票日期
                "发票号码": row[3],  # 发票号码
                "客户名称": row[7],  # 客户名称
                "货物名称": hwmc,  # 货物名称
                "备注软件名称": bzrjmc,  # 备注软件名称
                "数量": row[14],  # 数量
                "不含税金额": row[16],  # 不含税金额
                "税额": row[18],  # 税额
                "合计": SUM_IJ   # 合计
            }
            target_df = target_df.append(new_row, ignore_index=True)
            index_xh += 1

    # 将目标 DataFrame 写入 Excel 文件
    with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
        target_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

modify_excel_data(source_file, target_file, new_sheet_name, shengchan_file)
