# -*- coding: utf-8 -*-
# @Time : 2022/3/18 16:03
# @Author : lingz
# @Software: PyCharm


import os
import pandas as pd


input_path = "D:\docx2xlsx\classify_2"
num_1 = 0
num_2 = 0
filedict_list = []

dirs_list = os.listdir(input_path)
for index in range(0, len(dirs_list)):
    files_dirs = os.listdir(os.path.join(input_path+'/', dirs_list[index]))
    num_1 += 1
    num_in = 0
    data = {}
    for index_2 in range(0, len(files_dirs)):
        if os.path.splitext(files_dirs[index_2])[1] != ".xlsx":
            num_in += 1
            num_2 += 1
    data[dirs_list[index]] = num_in
    filedict_list.append(data)

print("sum_files:" + str(num_2))
print("sum_dirs:" + str(num_1))
print(filedict_list)
pf = pd.DataFrame(filedict_list)
file_path = pd.ExcelWriter('../check.xlsx')
pf.to_excel(file_path, encoding='utf-8', index=False)
file_path.save()
