# -*- coding: utf-8 -*-
# @Time : 2022/3/10 10:29
# @Author : lingz
# @Software: PyCharm
import copy
import os
import json
import summary_func as func


if __name__ == '__main__':
    with open('./path.json', 'r', encoding='utf-8') as fp:
        json_data = json.load(fp)
        # 已分类目录
        classify_path = json_data["classify"]["path"]
        # 重新分类目录
        classify_2_path = json_data["classify_2"]["path"]

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
                print("请输入您认为的归属类型号, 多个类型以空格隔开, 1.互联网+, 2.生命健康, 3.新材料, 4.碳达峰碳中和, 5.海洋强省, 6.科技强农， 7.其他")
                # TODO 如果要用已有字典自动匹配六大领域，在这里替换手动输入的列表
                temp_num_list = func.input_index()
                for temp_num in temp_num_list:
                    func.classify_2(dirs_path[index], static_type_path_list[temp_num - 1])
        func.merge_type_xlsx(classify_2_path)
    else:
        print('Path not exist')
