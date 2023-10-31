# This is a script for collecting specific columns in excel files
# from the folder where the script is place in
# The column numbers and row numbers can be specified by user.



import pandas as pd
import numpy
import os
import msvcrt

DEBUG = False

AL = "al"
AJ = "aj"


def getData(path):
    end_col = getUserInput()
    head = []
    data = []
    error_list = []
    flag = False
    for root_dir, sub_dir, files in os.walk(path):
        for file in files:
            if file.endswith('xlsx'):
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
                            print(f"Title: {len(header)}")
                            print(
                                "==========================================================================================")
                        flag = True
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
                dataframe.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=header)
                break
            elif isWriteData.lower() == "n":
                noWrite = True
                break
    except PermissionError:
        print("================================================================")
        print("result.xlsx 文件似乎被占用,输出文件改为 final_res.xlsx")
        print("================================================================")
        out_path = os.path.join(path, 'final_res.xlsx')
        dataframe.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=header)
        isError = True

    if noWrite:
        return
    if isError == False:
        print()
        print("汇总数据输出到result.xlsx文件")
        if len(error_list) != 0:
            print("下列文件读取出错")
            for i in error_list:
                print(f"错误文件: {i}")


def readExcel(file_name, end_col, flag):
    if flag == True:
        df = pd.read_excel(file_name, keep_default_na=False).iloc[:, : end_col].values.tolist()
        return df
    else:
        df = pd.read_excel(file_name, keep_default_na=False).iloc[:, : end_col]
        header = df.columns.tolist()
        df = df.values.tolist()
        if DEBUG:
            print(header)
        return df, header


def getUserInput():
    while True:
        end_col = input("请输入提取结束列数(1为A列,AA为27,以此类推): ")
        try:
            end_col = int(end_col)
            if end_col < 0:
                raise Exception
            else:
                end_col = end_col
                break
        except:
            print("输入格式错误,必须为大于等于0的整数")
    if DEBUG:
        print(f"end_col: {end_col}")

    return end_col


if __name__ == "__main__":
    dirc = os.getcwd()
    getData(dirc)
    os.system('pause')


