import pandas as pd
import openpyxl

source_file = 'apr/input_source.xlsx'
shengchan_file = 'apr/help.xlsx'
target_file = 'apr/target.xlsx'
new_sheet_name = 'default'

def second_star_index(text):
    first_star_index = text.find('*')
    if first_star_index == -1:
        return -1  # 如果没有找到第一个星号，返回-1
    
    second_star_index = text.find('*', first_star_index + 1)
    return second_star_index

def get_scbh_from_shengchan(shengchan_df, fphm):  # 获取生产编号
    row = shengchan_df.loc[shengchan_df['FPHM'] == fphm]
    if not row.empty:
        scbh = row.iloc[0]['SCBH']
        return scbh
    return None

def modify_excel_data(source_file, shengchan_file, target_file, new_sheet_name):
    try:
        # 使用pandas读取源文件
        source_df = pd.read_excel(source_file)
        shengchan_df = pd.read_excel(shengchan_file)
    except FileNotFoundError as e:
        print(f"文件不存在或者路径错误===>:\n {e}")
        return
    except Exception as e:
        print(f"An unexpected error occurred while loading the source file: {e}")
        return
    
    # 创建目标DataFrame
    target_df = pd.DataFrame(columns=["序号", "生产编号", "开票日期", "发票号码", "客户名称", "货物名称", "备注软件名称", "数量", "不含税金额", "税额", "合计", "硬件成本"])
    
    index_xh = 1
    for idx, row in source_df.iterrows():
        if idx == 0:
            continue  # 跳过标题行
        
        try:
            SUM_IJ = float(row[16]) + float(row[18])
        except ValueError:
            continue
        
        if pd.notna(row[26]) and "含" in row[26]:  # Modify remark software name
            index1 = row[26].find("含")
            bzrjmc = row[26][index1:]
            row6index = second_star_index(row[11])
            hwmc = row[11][row6index+1:]
            
            # 获取发票号码对应的生产编号
            fphm = row[3]
            scbh = get_scbh_from_shengchan(shengchan_df, fphm)
            
            # 创建新行
            new_row = {
                "序号": index_xh,
                "生产编号": scbh if scbh else 0,
                "开票日期": row[8],
                "发票号码": fphm,
                "客户名称": row[7],
                "货物名称": hwmc,
                "备注软件名称": bzrjmc,
                "数量": row[14],
                "不含税金额": row[16],
                "税额": row[18],
                "合计": SUM_IJ,
                "硬件成本": row[20] if len(row) > 20 else None  # 假设硬件成本在第21列
            }
            target_df = target_df.append(new_row, ignore_index=True)
            index_xh += 1

    # 将目标DataFrame写入目标Excel文件
    with pd.ExcelWriter(target_file, engine='openpyxl') as writer:
        target_df.to_excel(writer, sheet_name=new_sheet_name, index=False)

if __name__ == "__main__":
    modify_excel_data(source_file, shengchan_file, target_file, new_sheet_name)
