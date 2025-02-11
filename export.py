'''
    导出翻译文件
    从ts文件中导出翻译信息，并生成 excel 文件

    指令：
    python export.py -i={ts_dir} -o={TS-DATA-excel_file}
'''

import argparse
import os
from datetime import datetime
from common import *
import concurrent.futures

PARSER = argparse.ArgumentParser(
    description="""
    """)

PARSER.add_argument('-i', '--input', type=str,
                    help='TS file directory, non-recursive search')
PARSER.add_argument('-o', '--output', type=str,
                    help='Export excel file path.')
PARSER.add_argument('-r', '--recursive', action='store_true', default=False,
                    help='Enable recursive search for TS files.')

args = PARSER.parse_args()

print(args)

def process_ts_file(ts_file):
    return TsData(ts_file)

if __name__ == '__main__':
    ts_dir = args.input
    excel_file_path = args.output
    recursive = args.recursive  # Get the recursive flag

    # TS-DATA-excel_file
    excel_file_path = f"TS-{datetime.now().strftime('%Y-%m%d-')}{excel_file_path}.xlsx"

    # check excel_file_path is exists
    excel_file_path = CheckFileExists(excel_file_path)

    # Traverse all ts files under ts_dir
    ts_files = GetTsFiles(ts_dir, recursive)  # Pass the recursive flag
    print(ts_files)

    # Use ThreadPoolExecutor to process TS files in parallel
    with concurrent.futures.ThreadPoolExecutor() as executor:
        ts_data_list = list(executor.map(process_ts_file, ts_files))

    # out put excel
    excel_data = ExcelData(excel_file_path)
    excel_data.WriteData(ts_data_list)