'''
    This is a script for collecting specific names of excel files
    from the folder where the script is placed in
    The column number can be specified by user.
'''

import pandas as pd
import numpy
import os
import msvcrt

DEBUG = False


def getData(path: str) -> bool:
    end_col = getUserInput()
    part_num = getPartNum()
    head = []
    data = []
    error_list = []
    flag = False
    part_exists = True
    for root_dir, sub_dir, files in os.walk(path):
        for file in files:
            if file.endswith('xlsx') and part_num in file and not "result" in file:
                if DEBUG:
                    print(file)
                file_name = os.path.join(root_dir, file)
                try:
                    if flag == False:
                        df, header = readExcel(file_name, end_col, flag)
                        if DEBUG:
                            print(
                                "==========================================================================================")
                            print(header)
                            print(
                                "==========================================================================================")
                        flag = True
                        print(f"读取文件：{file_name}")
                    else:
                        df = readExcel(file_name, end_col, flag)

                        print(f"读取文件：{file_name}")
                    if DEBUG:
                        print(type(df))
                        print(df)
                    data.extend(df)
                    if DEBUG:
                        print()
                        print(data)
                        print()
                except PermissionError:
                    print("==========================================================================================")
                    print(f"{file_name}读取出错,请确认数据是否写入!")
                    print("==========================================================================================")
                    print()
                    next
                except IndexError:
                    print("==========================================================================================")
                    print(f"{file_name}读取失败,请确认该文件所选列是否有数据")
                    print("==========================================================================================")
                    print()
                    error_list.append(file_name)
                    next

    out_path = os.path.join(path, f"{part_num}-result.xlsx")
    if len(data) == 0:
        print(f"不存在此件号:<{part_num}>!")
        part_exists = False
        return False

    dataframe = pd.DataFrame(data)
    dataframe.dropna(axis=0, how='all', inplace=True)
    if DEBUG:
        print(dataframe)

    try:
        print()
        print(f"注意,如果{part_num}-result.xlsx已存在于目录中,该文件会被覆盖!")

        noWrite = False
        isError = False
        while True:
            isWriteData = input("是否继续? (y/n): ")
            if isWriteData.lower() == "y":
                dataframe.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=header)
                break
            elif isWriteData.lower() == "n":
                noWrite = True
                break
    except PermissionError:
        print("================================================================")
        print("result.xlsx 文件似乎被占用,输出文件改为 final_res.xlsx")
        print("================================================================")
        out_path = os.path.join(path, f"{part_num}-final_res.xlsx")
        dataframe.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=header)
        isError = True

    if noWrite:
        return
    if isError == False and part_exists:
        print()
        print(f"汇总数据输出到{part_num}result.xlsx文件")
        return True
        if len(error_list) != 0:
            print("下列文件读取出错")
            for i in error_list:
                print(f"错误文件: {i}")


def readExcel(file_name: str, end_col: int, flag: bool) -> list:
    '''
        prams : file_name = 所读文件的绝对地址; end_col = 所读的结束列数;
                flag = 是否需要每一列的Title
        returns : list of data, list of titles
    '''
    if flag == True:
        df = pd.read_excel(file_name, keep_default_na=True).iloc[:, : end_col].values.tolist()
        return df
    else:
        df = pd.read_excel(file_name, keep_default_na=True).iloc[:, : end_col]
        header = df.columns.tolist()
        df = df.values.tolist()
        if DEBUG:
            print(header)
        return df, header


def getUserInput() -> int:
    while True:
        end_col = input("请输入提取结束列数(1为A列,AA为27,以此类推): ")
        try:
            end_col = int(end_col)
            if end_col <= 0:
                raise Exception
            else:
                break
        except:
            print("输入格式错误,必须为大于等于1的整数")
    if DEBUG:
        print(f"end_col: {end_col}")

    return end_col


def getPartNum() -> int:
    while True:
        part_num = input("请输入提取的件号(请务必确认件号的准确性): ")
        while True:
            isPartNum = input("确认件号? (y/n): ")
            if isPartNum.lower() == "y":
                return part_num
            elif isPartNum.lower() == "n":
                break
            else:
                next


if __name__ == "__main__":
    dirc = os.getcwd()  # Get the current working directory
    getData(dirc)  # Start the program
    os.system('pause')  # Command line pause

