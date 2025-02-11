'''
    导出翻译文件
    从ts文件中导出翻译信息，并生成 excel 文件

    指令：
    python export.py -i={ts_dir} -o={excel_file}
'''

import argparse
import os

from common import *

PARSER = argparse.ArgumentParser(
    description="""
    """)

PARSER.add_argument('-i', '--input', type=str,
                    help='TS file directory, non-recursive search')
PARSER.add_argument('-o', '--output', type=str,
                    help='Export excel file path.')

args = PARSER.parse_args()

print(args)


if __name__ == '__main__':
    ts_dir = args.input
    excel_file_path = args.output

    # check excel_file_path is exists
    excel_file_path = CheckFileExists(excel_file_path)

    # Traverse all ts files under ts_dir
    ts_files = GetTsFiles(ts_dir)
    print(ts_files)

    # get all ts data
    ts_data_list = []
    for ts_file in ts_files:
        ts_data = TsData(ts_file)
        ts_data_list.append(ts_data)

    # out put excel
    excel_data = ExcelData(excel_file_path)
    excel_data.WriteData(ts_data_list)