import pandas as pd

# # 检查权限
# import ctypes, sys

# def is_admin():
#     try:
#         return ctypes.windll.shell32.IsUserAnAdmin()
#     except:
#         return False

# if not is_admin():
#     ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)



# 读取Excel文件
df = pd.read_excel("名单.xlsx")

# 初始化姓名字典
name_dic = {name: 0 for name in df['姓名']}

print(name_dic)

import os

# 初始化一个空的tuple用于存放作业文件中提取的姓名
work_list = tuple()

# 遍历当前目录下的所有文件
for filename in os.listdir('.'):
    # 假设文件名格式为“姓名其他信息.docx”，我们通过提取文件名来获取姓名
    # 这里使用的是简单的字符串分割方法，可能需要根据实际情况调整
    if filename.endswith('.docx'):  # 确保处理的是文档文件
        name_part = filename.split('230')[0]  # 使用姓名和学号之间的分隔符进行分割
        work_list += (name_part,)  # 将姓名加入到work_list中

print(work_list)

# 遍历name_dic中的每个姓名
for name in name_dic.keys():
    # 检查姓名是否在work_list中
    if name in work_list:
        name_dic[name] = 1  # 存在，更新为已提交
    else:
        name_dic[name] = 0  # 不存在，保持为未提交

print(name_dic)

# 将提交情况转换为“已交”或“未交”
df['提交情况'] = df['姓名'].map(lambda name: '已交' if name_dic[name] == 1 else '未交')

# 将更新后的DataFrame写回Excel文件，这里假设您想保留原文件名，进行覆盖保存
df.to_excel("名单.xlsx", index=False)
