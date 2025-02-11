'''
    从excel中导入数据到ts文件中

    指令：
    python import.py -i={excel_file} -o={ts_dir}
'''

import argparse
import os
import concurrent.futures

from common import *

PARSER = argparse.ArgumentParser(
    description="""
    """)

PARSER.add_argument('-i', '--input', type=str,
                    help='Import excel file path.')

PARSER.add_argument('-o', '--output', type=str,
                    help='Modified ts file directory, non-recursive search')

PARSER.add_argument('-r', '--recursive', action='store_true', default=False,
                    help='Enable recursive search for TS files.')

args = PARSER.parse_args()

def process_ts_file(ts_file, excel_data):
    ts_data = TsData(ts_file)
    ts_data.Update(excel_data)
    ts_data.Save()

if __name__ == '__main__':
    excel_file_path = args.input
    ts_dir = args.output
    recursive = args.recursive  # Get the recursive flag

    if not os.path.exists(excel_file_path):
        print(f"Excel file not found: {excel_file_path}")
        exit(1)

    excel_data = ExcelData(excel_file_path)
    excel_data.Load()

    # load ts 文件，修改它
    ts_files = GetTsFiles(ts_dir, recursive)

    # Use ThreadPoolExecutor to process TS files in parallel
    with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(lambda ts_file: process_ts_file(ts_file, excel_data), ts_files)

    print("Import finished")