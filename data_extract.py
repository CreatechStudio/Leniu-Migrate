import pandas as pd
import numpy as np
import re
import datetime
import os

# 定义输入文件名
file_name = input("请将xlsm首先导出为xlsx，并将其放入和本程序同一文件夹，并输入文件名（不含后缀名）:")
input_file_name = f'{file_name}.xlsx'
print(f'正在读取{input_file_name}...')

# 读取Excel文件
df = pd.read_excel(input_file_name)

# 设置pandas选项以避免科学计数法显示
pd.set_option('display.float_format', lambda x: f'{int(x)}' if np.isfinite(x) else '')

# 提取表头
headers = df.columns

# 创建空字典存储数据
data_dict = {}

# 遍历每一行
for index, row in df.iterrows():
    # 遍历每个表头和对应的单元格值
    for header in headers:
        value = row[header]
        # 存储到字典中
        if header not in data_dict:
            data_dict[header] = []
        data_dict[header].append(value)

# 公布变量
pub_data_dict = data_dict

# 打印全部原始数据
# for header, values in data_dict.items():
#     print(f"{header}: {values}")
# 打印pub_data_dict
# print(pub_data_dict)
# print(pub_invalid_rows)
# 测试输出特定数据
# print(pub_data_dict['CN_Name'][0])
# print(pub_data_dict['School_ID'][0])