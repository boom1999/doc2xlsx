# -*- coding: utf-8 -*-
# @Time : 2022/3/10 10:29
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
        # 已分类目录
        classify_path = json_data["classify"]["path"]
        # 重新分类目录
        classify_2_path = json_data["classify_2"]["path"]
        # 分类对应表
        classify_map_path = json_data["classify_map"]["path"]

    # 1.互联网+, 2.生命健康, 3.新材料, 4.碳达峰碳中和, 5.海洋强省, 6.科技强农， 7.其他
    static_type_name_list = ['互联网+', '生命健康', '新材料', '碳达峰碳中和', '海洋强省', '科技强农']
    static_type_path_list = []
    dirs_path = []
    temp_num_list = []
    func.mkdir_classify(classify_2_path)
    for index in range(0, len(static_type_name_list)):
        temp_path = os.path.join(classify_2_path + '/', static_type_name_list[index])
        static_type_path_list.append(func.mkdir_classify(temp_path))
    static_type_path_list.append(func.mkdir_classify(os.path.join(classify_2_path + '/', str('其他'))))
    if os.path.exists(classify_path):
        dirs_list = os.listdir(classify_path)
        print("当前需进行二次分类次数：", len(dirs_list))
        for index in range(0, len(dirs_list)):
            dirs_path.append(os.path.join(classify_path + '/', dirs_list[index]))
            similar_ratio = {}
            auto = 0
            for index_2 in range(0, len(static_type_name_list)):
                temp_dict = {static_type_name_list[index_2]: func.string_similar(dirs_list[index], static_type_name_list[index_2])}
                similar_ratio.update(temp_dict)
                if max(similar_ratio.values()) == 1.0:
                    print(dirs_list[index], ' and ', static_type_name_list[index_2], ' similar_ratio is:', similar_ratio)
                    func.classify_2(dirs_path[index], static_type_path_list[index_2])
                    auto = 1
                    break
            if auto == 0:
                print("当前处理的基础类别为：", dirs_list[index])
                temp_num_list = func.get_classify_2_list(classify_map_path, str(dirs_list[index]))
                for temp_num in temp_num_list:
                    func.classify_2(dirs_path[index], static_type_path_list[temp_num - 1])
        func.merge_type_xlsx(classify_2_path)
    else:
        print('Path not exist')
    func.reorder_type(classify_2_path)

    time_end = time.time()
    time_sum = time_end - time_start
    print("二次分类结束，共花费时间：" + str(time_sum) + 's')
