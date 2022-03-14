# -*- coding: utf-8 -*-
# @Time : 2022/3/4 14:41
# @Author : lingz
# @Software: PyCharm

import os
import json
import summary_func as func
import time

if __name__ == '__main__':
    time_start = time.time()
    with open('./path.json', 'r', encoding='utf-8') as fp:
        json_data = json.load(fp)
        # 输入word文件集目录
        # input_path = json_data["input"]["path"]
        # 输出excel表格目录
        output_path = json_data["output"]["path"]
        # 分类目录
        classify_path = json_data["classify"]["path"]

    input_path = '../docx/ZYJY'
    output_xlsx = os.path.join(output_path+'/', 'summary.xlsx')
    func.mkdir_classify(classify_path)
    if os.path.exists(input_path):
        dirs_path, docx_files_path = func.get_docx_files_path(input_path)
        func.classify(output_xlsx, dirs_path, classify_path)
    else:
        print('Path not exist')

    time_end = time.time()
    time_sum = time_end - time_start
    print("一次分类结束，共花费时间：" + str(time_sum) + 's')
