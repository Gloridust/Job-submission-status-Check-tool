# Configuration
excel_name = "名单.xlsx"
name_column = "姓名"
status_column = "提交情况"
file_extensions = ['.doc', '.docx']

from ltp import LTP  # 引入LTP
import pandas as pd
import os

# 初始化LTP模型
ltp = LTP()

def extract_name_from_filename(filename):
    seg, hidden = ltp.seg([filename])
    ner = ltp.ner(hidden)[0]

    for tag, start, end in ner:
        if tag == 'Nh':  # 'Nh' 代表人名
            return ''.join(seg[0][start:end + 1])
    return None

# 声明
print("你正在使用由 Gloridust 制作的 Job-submission-status-Check-tool\n如果你喜欢这个程序，不妨在Github上面start该项目：https://github.com/Gloridust/Job-submission-status-Check-tool\n在使用软件之前，请确保您正确放置了所有文件并正确配置了文件\n如果有任何问题，你可以在该项目的readme文件或issue中寻求帮助\n")
print(f"当前的配置信息如下：\n名单表格：{excel_name}\n姓名所在列：{name_column}\n提交情况列：{status_column}\n提交的文件类型：{file_extensions}")
input("确认无误后，按下Enter键开始")

# 读取Excel文件
df = pd.read_excel(excel_name)

# 初始化姓名字典
name_dic = {name: 0 for name in df[name_column]}

# 初始化一个空的tuple用于存放作业文件中提取的姓名
work_list = tuple()

# 遍历当前目录下的所有文件
for filename in os.listdir('.'):
    # 使用ltp从文件名中提取姓名
    name = extract_name_from_filename(filename)
    if name:
        work_list += (name,)

# 遍历name_dic中的每个姓名
for name in name_dic.keys():
    # 检查姓名是否在work_list中
    if name in work_list:
        name_dic[name] = 1  # 存在，更新为已提交
    else:
        name_dic[name] = 0  # 不存在，保持为未提交

have_sub_num = 0
have_sub = "已提交的有："
for name, status in name_dic.items():
    if status == 1:
        have_sub_num += 1
        have_sub += (name + ",")
print("已提交人数：", have_sub_num)
print(have_sub)

not_sub_num = 0
not_sub = "还未提交的有："
for name, status in name_dic.items():
    if status == 0:
        not_sub_num += 1
        not_sub += (name + ",")
print("未提交人数：", not_sub_num)
print(not_sub)

# 将提交情况转换为“已交”或“未交”
df[status_column] = df[name_column].map(lambda name: '已交' if name_dic[name] == 1 else '未交')

# 将更新后的DataFrame写回Excel文件，这里假设您想保留原文件名，进行覆盖保存
df.to_excel(excel_name, index=False)
print("已将提交情况保存至表格")
input("按下Enter键结束...")
