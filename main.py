# 配置信息
excel_name="名单.xlsx"
name_column="姓名"
status_column="提交情况"
file_extensions = ['.doc', '.docx']
name_is_before="230"

# 声明
print("你正在使用由 Gloridust 制作的 Job-submission-status-Check-tool \n如果你喜欢这个程序，不妨在Github上面start该项目：https://github.com/Gloridust/Job-submission-status-Check-tool \n在使用软件之前，请确保您正确放置了所有文件并正确配置了文件 \n如果有任何问题，你可以在该项目的readme文件或issue中寻求帮助\n")
print("当前的配置信息如下：\n名单表格：{excel_name}\n姓名所在列：{name_column}\n提交情况列：{status_column}\n提交的文件类型：{file_extensions}\n名字部分前置于：{name_is_before}")
input("确认无误后，按下Enter键开始")

import pandas as pd

# 读取Excel文件
df = pd.read_excel(excel_name)

# 初始化姓名字典
name_dic = {name: 0 for name in df[name_column]}

# print(name_dic)

import os

# 初始化一个空的tuple用于存放作业文件中提取的姓名
work_list = tuple()

# 遍历当前目录下的所有文件
for filename in os.listdir('.'):
    # 假设文件名格式为“姓名其他信息.docx”，我们通过提取文件名来获取姓名
    # 这里使用的是简单的字符串分割方法，可能需要根据实际情况调整
    if any(filename.endswith(ext) for ext in file_extensions):  # 确保处理的是文档文件
        name_part = filename.split(name_is_before)[0]  # 使用姓名和学号之间的分隔符进行分割
        work_list += (name_part,)  # 将姓名加入到work_list中

# print(work_list)

# 遍历name_dic中的每个姓名
for name in name_dic.keys():
    # 检查姓名是否在work_list中
    if name in work_list:
        name_dic[name] = 1  # 存在，更新为已提交
    else:
        name_dic[name] = 0  # 不存在，保持为未提交

# print(name_dic)
have_sub_num = 0
have_sub = "已提交的有："
for name, status in name_dic.items():
    if status == 1:  # 检查值是否为1
        have_sub_num += 1
        have_sub += (name + ",")  # 记录
print("已提交人数：",have_sub_num)
print(have_sub)

not_sub_num = 0
not_sub = "还未提交的有："
for name, status in name_dic.items():
    if status == 0:  # 检查值是否为0
        not_sub_num += 1
        not_sub += (name + ",")  # 记录
print("未提交人数：",not_sub_num)
print(not_sub)

# 将提交情况转换为“已交”或“未交”
df[status_column] = df[name_column].map(lambda name: '已交' if name_dic[name] == 1 else '未交')

# 将更新后的DataFrame写回Excel文件，这里假设您想保留原文件名，进行覆盖保存
df.to_excel(excel_name, index=False)
print("已将提交情况保存至表格")
