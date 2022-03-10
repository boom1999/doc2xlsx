import pandas as pd
from pandas import DataFrame
import os

static_path = "../classify_2"
static_type_path_list = os.listdir(static_path)
for index in range(0, len(static_type_path_list)):
    file_name = static_type_path_list[index]
    static_type_path_list[index] = os.path.join(static_path+'/', static_type_path_list[index])
    DFs = []
    dirs_list = os.listdir(static_type_path_list[index])
    for dir_path in dirs_list:
        if os.path.splitext(dir_path)[1] == ".xlsx":
            file_path = os.path.join(static_type_path_list[index] + '/', dir_path)
            df = pd.read_excel(file_path)
            DFs.append(df)
            os.remove(file_path)
    writer = pd.ExcelWriter(os.path.join(static_type_path_list[index] + '/', file_name+'.xlsx'))
    pd.concat(DFs).to_excel(writer, index=False)
    writer.save()

