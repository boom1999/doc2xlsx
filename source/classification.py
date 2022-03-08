# -*- coding: utf-8 -*-
# @Time : 2022/3/4 14:41
# @Author : lingz
# @Software: PyCharm

import os
import json
import summary_func as func

if __name__ == '__main__':
    with open('./path.json', 'r', encoding='utf-8') as fp:
        json_data = json.load(fp)
        # 输入word文件集目录
        input_path = json_data["input"]["path"]
        # 输出excel表格目录
        output_path = json_data["output"]["path"]
        # 匹配文件map目录
        map_tables_path = json_data["map"]["path"]
        # 分类目录
        classify_path = json_data["classify"]["path"]

    func.mkdir_classify(classify_path)
    if os.path.exists(input_path):
        dirs_path, docx_files_path = func.get_docx_files_path(input_path)
        func.classify(output_path, dirs_path, classify_path)
    else:
        print('Path not exist')

