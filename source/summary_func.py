# -*- coding: utf-8 -*-
# @Time : 2022/3/4 14:43
# @Author : lingz
# @Software: PyCharm

import copy
from docx import Document
import xlrd
import pandas as pd
from pandas import DataFrame
import os
import re
import shutil


def get_docx_files_path(input_path):
    """
    input dirs of the docx_s, then listdir input_path-->dirs_path-->docx_files_path.
    :param input_path: the path containing the whole input files.
    :return: the path of the dirs and the docx_file which need to tidy up.
    """
    docx_files_path = []
    docx_files = []
    dirs_path = os.listdir(input_path)
    for index in range(0, len(dirs_path)):
        dirs_path[index] = str(os.path.join(input_path, dirs_path[index]))
        files = os.listdir(dirs_path[index])
        for file in files:
            if os.path.splitext(file)[1] == ".docx":
                docx_files.append(file)
        docx_files_path.append(os.path.join(dirs_path[index]+'/', docx_files[index]))
    return dirs_path, docx_files_path


def read_docx(read_path):
    """

    :param read_path: .docx files.
    :return: list_ -- contain Dicts = {":", ";"...}.
    """
    input_path = copy.deepcopy(read_path)
    docx = Document(input_path)
    table_s = docx.tables

    # 从文件路径拆分出文件名，从文件名中按照‘-’拆出编码 code_
    (filepath, filename) = os.path.split(input_path)
    code_list = filename.split('-')[0:3]
    code_ = code_list[0] + '-' + code_list[1] + '-' + code_list[2]

    list_ = []
    for table in table_s:
        for j in range(1, len(table.rows)):
            dict_ = {'序号': table.cell(j, 0).text,
                     '展品名称': table.cell(j, 1).text,
                     '所属领域': table.cell(j, 2).text,
                     '展品形式': table.cell(j, 3).text,
                     '展品单位': table.cell(j, 4).text,
                     '联系人': table.cell(j, 5).text,
                     '联系电话': table.cell(j, 6).text,
                     '编码': code_}
            if dict_['展品名称'] != '':
                list_.append(dict_)
                print("Append successfully: row " + str(j))
    return list_


def add_type_region(list_in, map_tables_path):
    """

    :param map_tables_path: path of the map, original list and their types and regions.
    :param list_in: putin original list which need to add type and region.
    :return: final list -- containing type and region.
    """

    # Use copy.deepcopy to avoid to change list_in outsider
    list_out = copy.deepcopy(list_in)
    # 这里需要拆分目录读取编码
    code_ = '编码'
    flat_ = '展品单位'
    # flat_simple = '单位简称'
    type_ = '类型'
    origin_ = '地区'
    # Attention! 'map_tables_path' can't have head index, pd.read_excel can't read.
    # Or set 'header= '
    map_data = pd.read_excel(map_tables_path, index_col='编码', header=1, sheet_name='Sheet 1')
    code_list = list(map_data.index)
    for j in range(0, len(list_out)):
        if list_out[j][code_] not in code_list:
            print(list_out[j][flat_] + "出现编码错误，请手动修改!")
            continue
        list_out[j][flat_] = map_data.loc[list_out[j][code_], flat_]
        # list_out[j][flat_simple] = map_data.loc[list_out[j][code_], flat_simple]
        list_out[j][type_] = map_data.loc[list_out[j][code_], type_]
        list_out[j][origin_] = map_data.loc[list_out[j][code_], origin_]

    return list_out


def export2excel(export, out):
    """

    :param export: list_sum -- contain Dicts = {":", ";"...} +  {":", ";"...} + ...
    :param out: outfile .xlsx file.
    :return:
    """
    pf = pd.DataFrame(list(export))

    # Redefine column labels
    order = ['编码', '序号', '展品名称', '所属领域', '展品形式', '展品单位', '类型', '地区', '联系人', '联系电话']
    pf = pf[order]

    file_path = pd.ExcelWriter(out)
    pf.fillna(' ', inplace=True)
    pf.to_excel(file_path, encoding='utf-8', index=False)
    file_path.save()


def mkdir_classify(classify_path):
    """

    :param classify_path: The path need to be created.
    :return: The classify_path.
    """
    classify_path = classify_path.strip()
    classify_path = classify_path.rstrip("\\")

    isExists = os.path.exists(classify_path)
    if not isExists:
        os.makedirs(classify_path)
        print("分类路径"+str(classify_path)+"创建成功")
        return classify_path
    else:
        print("分类路径"+str(classify_path)+"已存在")
        return classify_path


def classify(summary_path, dirs_path, classify_path):
    """

    :param summary_path: The original statistical table path.
    :param dirs_path: The file summary path.
    :param classify_path: The output path.
    :return:
    """
    path = copy.deepcopy(summary_path)
    data = pd.read_excel(path, index_col='所属领域')
    summary_type_list = list(set(data.index))
    for summary_type in summary_type_list:
        type_path = mkdir_classify(os.path.join(classify_path+'/', str(summary_type)))
        # 注意：这里data.loc[[summary_type], :])第一个参数必须要加[]，否则会解析成series.Series，需要使用frame.DataFrame保存
        target_path = str(type_path) + '/'
        temp_target_path = copy.deepcopy(target_path)
        data.loc[[summary_type], :].to_excel(target_path + summary_type + '汇总表.xlsx')
        print(summary_type + '汇总表.xlsx 创建成功')
        print("--------------------------------------------------------------------")
        print("Moving file structure......")
        temp_data = data.loc[[summary_type], :]
        for i in range(0, temp_data.shape[0]):
            code_ = temp_data.iloc[i, 0]
            num_ = temp_data.iloc[i, 1]
            name_ = temp_data.iloc[i, 4]
            for dir_path in dirs_path:
                f_path, f_name = os.path.split(dir_path)
                if re.match(code_ + '.+', f_name) is not None:
                    temp_dirs_path = os.listdir(dir_path)
                    for temp_dir_path in temp_dirs_path:
                        # 用正则从路径中拆出数字, like: 10、源牌电磁水表（实物） to ['10', '']
                        path_list = re.split("[^0-9]+", temp_dir_path)
                        for path in path_list:
                            if path == str(num_) and os.path.splitext(temp_dir_path)[1] != ".docx":
                                src_path = os.path.join(dir_path+'/', temp_dir_path)
                                target_path = os.path.join(target_path, '('+str(name_)+')'+temp_dir_path)
                                if not os.path.exists(target_path):
                                    shutil.copytree(src_path, target_path)
                                    target_path = copy.deepcopy(temp_target_path)
                                    print(str(src_path) + '-->' + str(target_path) + "移动成功!")
                                else:
                                    target_path = copy.deepcopy(temp_target_path)
                                    print("----------------目标路径%s已存在，请检查！------------------" % target_path)
                                break
                    break
        print("--------------------------------------------------------------------")
