# This is a script for collecting specific columns in excel files
# from the folder where the script is placed in


import pandas as pd
import numpy
import os
import msvcrt

DEBUG = True


def Col2Int(s: str) -> int:
    assert (isinstance(s, str))
    for i in s:
        if not 64 < ord(i) < 91:
            raise ValueError('Excel Column ValueError')
    return sum([(ord(n) - 64) * 26 ** i for i, n in enumerate(list(s)[::-1])]) - 1


def getCol():
    while True:
        col = input("请输入提取列数: ")
        col_num = Col2Int(col)
        while True:
            isCol = input(f"确认列数<{col}>(y/n): ")
            if isCol.lower() == "y":
                return col_num
            elif isCol.lower() == "n":
                break
            else:
                next


def getColumns(path):
    target = getCol()
    data = set()

    for root_dir, sub_dir, files in os.walk(path):
        for file in files:
            if file.endswith('xlsx'):
                if DEBUG:
                    print(file)
                file_name = os.path.join(root_dir, file)
                try:
                    df_1 = readExcel(file_name, target)

                    print(f"读取文件：{file_name}")
                    if DEBUG:
                        print(df_1)
                    data = data | df_1
                except PermissionError:
                    print("================================================================")
                    print(f"{file_name}读取出错,请确认数据是否写入!")
                    print("================================================================")
                    print()
                    next
                except IndexError:
                    print("================================================================")
                    print(f"{file_name}读取失败,请确认该文件所选列是否有数据")
                    print("================================================================")
                    print()
                    next

    out_path = os.path.join(path, 'result.xlsx')
    dataframe = pd.DataFrame(data)

    if DEBUG:
        print(dataframe)

    try:
        print()
        print("注意,如果result.xlsx已存在于目录中,该文件会被覆盖!")

        noWrite = False
        isError = False
        while True:
            isWriteData = input("是否继续? (y/n): ")
            if isWriteData.lower() == "y":
                dataframe.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=False)
                break
            elif isWriteData.lower() == "n":
                noWrite = True
                break
    except PermissionError:
        print("================================================================")
        print("result.xlsx 文件似乎被占用,输出文件改为 final_res.xlsx")
        print("================================================================")
        out_path = os.path.join(path, 'final_res.xlsx')
        dataframe.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=False)
        isError = True

    if noWrite:
        return
    if isError == False:
        print()
        print("汇总数据输出到result.xlsx文件")


def readExcel(file_name, target):
    df = set(pd.read_excel(file_name, keep_default_na=False).iloc[:, target])
    if "" in df:
        df.remove("")
    s = set()
    for i in df:
        l = str(i).split("@\n")
        for j in l:
            s.add(j)
    return s  # return a set of each file's column's data


if __name__ == "__main__":
    dirc = os.getcwd()
    getColumns(dirc)
    os.system('pause')



