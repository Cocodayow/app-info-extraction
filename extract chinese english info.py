import os
import docx2txt
import pandas as pd
from docx import Document

has_been_called = False

docxDir = input("请输入要转换的docx文件目录: ")

# 获取用户输入的没有页眉的docx文件目录
wdocxDir = input("请输入没有页眉的docx文件目录: ")

# 获取用户输入的转换成功的xlsx文件目录
xlsxDir = input("导出为xlsx到: ")

# get word file
def getfile():
    DEBUG = False
    # 文件总数
    fileTotal = 0
    # 转换正确的文件总数
    success_list = []
    # 装换错误的文件总数
    error_list = []
    data = []
    # 去掉页眉
    for root_dir, sub_dir, files in os.walk(docxDir):
        for file in files:
            if file.endswith('.docx'):
                file_name = os.path.join(root_dir, file)
                document = Document(file_name)
                for section in document.sections:
                    section.different_first_page_header_footer = False
                    section.header.is_linked_to_previous = True
                try:
                    document.save(os.path.join(wdocxDir, file))
                except Exception as e:
                    print(f"无法去掉文件页眉 Error processing file: {file_name}\n{str(e)}")
    # 读取去掉页眉的文件，并将其转换为txt,然后转化为dataframe
    for root_dir, sub_dir, files in os.walk(wdocxDir):
        for file in files:
            if file.endswith('.docx'):
                fileTotal += 1
                file_name = os.path.join(root_dir, file)
                try:
                    text = docx2txt.process(file_name)
                    lines = text.split('\n')
                    for line in lines:
                        if line.strip() != "":
                            data.append({'Text': line})
                    success_list.append(file_name)
                except Exception as e:
                    error_list.append(file_name)
                    print(f"读取出错 Error processing file: {file_name}\n{str(e)}")
    global has_been_called
    if not has_been_called:
        print("================================================================")
        print(f"读取完成，共有{fileTotal}个docx文件，成功读取{len(success_list)}个文件，共有{len(error_list)}"
              f"个文件读取出错")
        print(f"读取出错的文件为：{error_list}")
        print("================================================================")
        has_been_called = True
    df = pd.DataFrame(data)
    if DEBUG:
        print(df)
    return df




# read every production index
def get_line_numbers_starting_with():
    df = getfile()
    DEBUG = False
    line_numbers = []
    for idx, text in enumerate(df['Text']):
        if text.strip().startswith('..'):
            line_numbers.append(idx)
    if DEBUG:
        print(line_numbers)
    return line_numbers
#

# check if it's english or not
def isEnglish(checkStr):
    for ch in checkStr:
        if u'\u4e00' <= ch <= u'\u9fff':
            return False
    return True

def extract_lines_between():
    df = getfile()
    line_numbers = get_line_numbers_starting_with()
    DEBUG = False


    for i in range(len(line_numbers)):
        chinese_lines = []
        english_lines = []
        if i == len(line_numbers) - 1:
            newdf = df.loc[line_numbers[i] + 1:]
        else:
            newdf = df.loc[line_numbers[i] + 1:line_numbers[i + 1] - 1]
        index_col = []
        for index, row in newdf.iterrows():

            if isEnglish(row['Text']):
                english_lines.append(row['Text'])
            else:
                chinese_lines.append(row['Text'])
        max_size = max(len(english_lines), len(chinese_lines))
        english_lines += [''] * (max_size - len(english_lines))
        chinese_lines += [''] * (max_size - len(chinese_lines))
        step = df.loc[line_numbers[i]].get('Text')
        index_col += max_size * [step]
        # 创建一个dataframe，一列是英文，一列是中文
        data = {'Index': index_col, 'English': english_lines, 'Chinese': chinese_lines}
        result_df = pd.DataFrame(data)
        result_df.set_index('Index', inplace = True)
        merged_df = result_df.groupby('Index').apply(lambda x: pd.Series({
            'English': '\n'.join(x['English']),
            'Chinese': '\n '.join(x['Chinese'])
        }))
        merged_df.reset_index(inplace=True)
        output_path = os.path.join(xlsxDir, '中英文对照工部.xlsx')
        # Check if the output file already exists
        if os.path.isfile(output_path):
            if i == 0:
                # Clear the file if it's the first iteration
                with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
                    merged_df.to_excel(writer, index=False)
            else:
                existing_df = pd.read_excel(output_path)
                merged_df = pd.concat([existing_df, merged_df], ignore_index=True)
                merged_df.to_excel(output_path, index=False)
        else:
            # Save the DataFrame to the output file
            merged_df.to_excel(output_path, index=False)

getfile()
get_line_numbers_starting_with()
extract_lines_between()

