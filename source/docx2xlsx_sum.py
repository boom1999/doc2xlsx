# -*- coding: utf-8 -*-
# @Time : 2022/2/24 9:28
# @Author : lingz
# @Software: PyCharm
import copy

from docx import Document
import xlrd
import pandas as pd
import os


def file_rename(file_path):
    """

    :param file_path: path of the docx root  file
    :return: rename docx files followed certain rules, null return
    """
    path = file_path
    files = os.listdir(path)
    for i, file in enumerate(files):
        OldName = os.path.join(path, file)
        NewName = os.path.join(path, "test_file_"+str(i)+"_.docx")
        os.rename(OldName, NewName)


def read_docx(read):
    """

    :param read: .docx files
    :return: list_ -- contain Dicts = {":", ";"...}
    """
    input_path = read
    docx = Document(input_path)
    table_s = docx.tables
    list_ = []
    for table in table_s:
        for j in range(1, len(table.rows)):
            dict_ = {'序号': table.cell(j, 0).text,
                     '编码': table.cell(j, 1).text,
                     '展品名称': table.cell(j, 2).text,
                     '所属领域': table.cell(j, 3).text,
                     '展品形式': table.cell(j, 4).text,
                     '展品单位': table.cell(j, 5).text,
                     '联系人': table.cell(j, 6).text,
                     '联系电话': table.cell(j, 7).text}
            if dict_['编码'] != '':
                list_.append(dict_)
                print("Append successfully: row " + str(j))
    return list_


def add_type_region(list_in, map_tables_path):
    """

    :param map_tables_path: path of the map, original list and their types and regions
    :param list_in: putin original list which need to add type and region
    :return: final list -- containing type and region
    """

    # Use copy.deepcopy to avoid to change list_in outsider
    list_out = copy.deepcopy(list_in)
    code_ = '编码'
    flat_ = '展品单位'
    flat_simple = '单位简称'
    type_ = '类型'
    origin_ = '地区'
    # Attention! 'map_tables_path' can't have head index, pd.read_excel can't read.
    map_data = pd.read_excel(map_tables_path, index_col='编码', sheet_name='Sheet 1')
    code_list = list(map_data.index)
    for j in range(0, len(list_out)):
        if list_out[j][code_] not in code_list:
            print(list_out[j][flat_]+"的一项展品编码出错，请手动修改!")
            continue
        list_out[j][flat_] = map_data.loc[list_out[j][code_], flat_]
        list_out[j][flat_simple] = map_data.loc[list_out[j][code_], flat_simple]
        list_out[j][type_] = map_data.loc[list_out[j][code_], type_]
        list_out[j][origin_] = map_data.loc[list_out[j][code_], origin_]
    return list_out


def export2excel(export, out):
    """

    :param export: list_sum -- contain Dicts = {":", ";"...} +  {":", ";"...} + ...
    :param out: outfile .xlsx file
    :return:
    """
    pf = pd.DataFrame(list(export))

    # Redefine column labels
    order = ['序号', '编码', '展品名称', '所属领域', '展品形式', '展品单位', '单位简称', '类型', '地区', '联系人', '联系电话']
    pf = pf[order]

    file_path = pd.ExcelWriter(out)
    pf.fillna(' ', inplace=True)
    pf.to_excel(file_path, encoding='utf-8', index=False)
    file_path.save()


if __name__ == '__main__':
    # 输出excel表格目录
    output_path = "../xlsx/out_doc2xl.xlsx"
    # 输入word文件集目录
    files_path = '../docx/'
    docx_files = os.listdir(files_path)
    num_docx = len(docx_files)
    map_tables_path = "../map/map.xlsx"
    list_sum = []

    if os.path.exists(files_path):
        file_rename(files_path)
        files = os.listdir(files_path)
        for i, file in enumerate(files):
            tables = read_docx(files_path+file)
            list_sum.extend(tables)
            print("Extend docx "+str(i+1)+" successfully!\n")
    else:
        print('Path not exist')
    final_list = add_type_region(list_sum, map_tables_path)
    export2excel(final_list, output_path)
    print("Change successfully!")
