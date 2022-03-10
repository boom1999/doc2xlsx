# -*- coding: utf-8 -*-
# @Time : 2022/2/24 9:28
# @Author : lingz
# @Software: PyCharm

import json
import summary_func as func
import os


if __name__ == '__main__':
    with open('./path.json', 'r', encoding='utf-8') as fp:
        json_data = json.load(fp)
        # 输入word文件集目录
        input_path = json_data["input"]["path"]
        # 输出excel表格目录
        output_path = json_data["output"]["path"]
        # 匹配文件map目录
        map_tables_path = json_data["map"]["path"]

    list_sum = []
    if os.path.exists(input_path):
        dirs_path, docx_files_path = func.get_docx_files_path(input_path)
        # 加文件汇集分类操作
        for docx_file_path in docx_files_path:
            print("Loading files:'"+str(docx_file_path)+"'......")
            tables = func.read_docx(input_path + str(docx_file_path))
            list_sum.extend(tables)
            print("Extend docx:' "+str(docx_file_path)+" 'successfully!\n")
    else:
        print('Path not exist')
    final_list = func.add_type_region(list_sum, map_tables_path)
    func.export2excel(final_list, output_path)
    func.cell_handling(output_path)
    print("Change successfully!")
