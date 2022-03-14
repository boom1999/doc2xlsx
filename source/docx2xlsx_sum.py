# -*- coding: utf-8 -*-
# @Time : 2022/2/24 9:28
# @Author : lingz
# @Software: PyCharm

import json
import summary_func as func
import os
import time
import pandas as pd

if __name__ == '__main__':
    time_start = time.time()
    with open('./path.json', 'r', encoding='utf-8') as fp:
        json_data = json.load(fp)
        # 输入word文件集目录
        input_path = json_data["input"]["path"]
        # 输出excel表格目录
        output_path = json_data["output"]["path"]
        # 匹配文件map目录
        map_tables_path = json_data["map"]["path"]

    input_path_subdir_list = os.listdir(input_path)
    for index in range(0, len(input_path_subdir_list)):
        merge_dir = os.path.join(input_path+'/', input_path_subdir_list[index]+'/')
        list_sum = []
        if os.path.exists(merge_dir):
            dirs_path, docx_files_path = func.get_docx_files_path(merge_dir)
            # 加文件汇集分类操作
            for docx_file_path in docx_files_path:
                print("Loading files:'" + str(docx_file_path) + "'......")
                tables = func.read_docx(str(docx_file_path))
                list_sum.extend(tables)
                print("Extend docx:' " + str(docx_file_path) + " 'successfully!\n")
        else:
            print('Path not exist')
        final_list = func.add_type_region(list_sum, map_tables_path)
        export_path = os.path.join(output_path+'/', input_path_subdir_list[index]+'.xlsx')
        func.export2excel(final_list, export_path)
        func.cell_handling(export_path)
        print("Exporting ' " + input_path_subdir_list[index] + " ' successfully!")

    func.merge_summary_xlsx(output_path, deleting=False)
    time_end = time.time()
    time_sum = time_end - time_start
    print("统计结束，共花费时间：" + str(time_sum) + 's')
