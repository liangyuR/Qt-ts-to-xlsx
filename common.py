import os
import xml.etree.ElementTree as ET
from tqdm import tqdm
import pandas as pd

__all__ = [
    'GetTsFiles',
    'TsData',
    'ExcelData',
    'CheckFileExists',
]

def GetTsFiles(ts_dir, recursive=False):
    ts_files = []
    for file in os.listdir(ts_dir):
        if file.endswith('.ts'):
            ts_files.append(os.path.join(ts_dir, file))
        elif recursive and os.path.isdir(os.path.join(ts_dir, file)):
            ts_files.extend(GetTsFiles(os.path.join(ts_dir, file), recursive))
    return ts_files

def CheckFileExists(file_path):
    if os.path.exists(file_path):
        print(f"File already exists: {file_path}")
        # question user using new file name
        new_file_name = input("Please input new file name: ")
        new_file_path = os.path.join(os.path.dirname(file_path), new_file_name)
        print(f"New file path: {new_file_path}")
        return CheckFileExists(new_file_path)
    return file_path

class TsData:
    def __init__(self, ts_file):
        self.ts_file = ts_file
        if not os.path.exists(ts_file):
            raise FileNotFoundError(f"File not found: {ts_file}")
        self.tree = ET.parse(ts_file)
        self.data = {}
        self.lang_type = ""
        self.Load()

    def Load(self):
        self.lang_type = self.tree.getroot().get('language', '')
        if self.lang_type == '':
            raise ValueError(f"Language type not found in ts file: {self.ts_file}")

        for context in tqdm(self.tree.findall('context'), desc='Processing', unit='context'):
            class_name = context.find('name').text
            for message in context.findall('message'):
                source = message.find('source').text
                comment_element = message.find('comment')
                comment = comment_element.text if comment_element is not None else ''
                translation = message.find('translation')

                self.data[(class_name, source, comment)] = translation.text

    def Update(self, excel_data):
        for key, value in excel_data.data.items():
            if key in self.data:
                self.data[key] = value[self.lang_type]

    def Save(self):
        for context in tqdm(self.tree.findall('context'), desc='Processing', unit='context'):
            class_name = context.find('name').text
            for message in context.findall('message'):
                source = message.find('source').text

                comment = message.find('comment')
                comment_text = comment.text if comment is not None else ''

                if (class_name, source, comment_text) in self.data:
                    translation = message.find('translation')
                    translation.text = self.data[(class_name, source, comment_text)]

                    if self.data[(class_name, source, comment_text)] != '':
                        if 'type' in translation.attrib:
                            del translation.attrib['type']
                    else:
                        translation.set('type', 'unfinished')
                        translation.text = ''

        # Save to file with proper XML declaration, DOCTYPE, and encoding
        with open(self.ts_file, 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="utf-8"?>\n')
            f.write('<!DOCTYPE TS>\n')
            self.tree.write(f, encoding='unicode')

    # print data
    def PrintData(self):
        for key, value in self.data.items():
            print(f"{key}: {value}")

class ExcelData:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        if not excel_file.endswith('.xlsx'):
            self.excel_file = excel_file + '.xlsx'
        self.data = {}

    def Load(self):
        # 读取 Excel 文件
        df = pd.read_excel(self.excel_file)
        
        # 将 DataFrame 转换为字典格式，添加进度条
        for _, row in tqdm(df.iterrows(), desc='Loading Excel data', total=len(df)):
            class_name = row['ClassName']
            source = row['Source']
            comment = row['Comment']
            if pd.isna(comment):
                comment = ''

            # 初始化该key的字典
            if (class_name, source, comment) not in self.data:
                self.data[(class_name, source, comment)] = {}
            
            # 对每种语言类型的翻译进行处理
            for lang in ['zh_CN', 'en_US', 'ja_JP', 'ko_KR']:
                if pd.notna(row[lang]):  # 检查翻译是否存在
                    self.data[(class_name, source, comment)][lang] = row[lang]
                else:
                    self.data[(class_name, source, comment)][lang] = ''

    def PrintData(self):
        for key, value in self.data.items():
            print(f"{key}: {value}")

    def WriteData(self, ts_data_list):
        rows = {}
        for ts_data in tqdm(ts_data_list, desc='Processing TS files'):
            for key, value in ts_data.data.items():
                if key not in rows:
                    rows[key] = {
                        'zh_CN': '',
                        'en_US': '',
                        'ja_JP': '',
                        'ko_KR': ''
                    }
                rows[key][ts_data.lang_type] = value

        df_rows = []
        for (class_name, source, comment), translations in tqdm(rows.items(), desc='Preparing Excel rows'):
            df_rows.append({
                'ClassName': class_name,
                'Source': source,
                'Comment': comment,
                'zh_CN': translations['zh_CN'],
                'en_US': translations['en_US'],
                'ja_JP': translations['ja_JP'],
                'ko_KR': translations['ko_KR']
            })

        self.data = pd.DataFrame(df_rows, columns=['ClassName', 'Source', 'Comment', 'zh_CN', 'en_US', 'ja_JP', 'ko_KR'])
        print('Saving Excel file...')
        
        if not self.excel_file.endswith('.xlsx'):
            self.excel_file += '.xlsx'
        self.data.to_excel(self.excel_file, index=False, engine='openpyxl')
