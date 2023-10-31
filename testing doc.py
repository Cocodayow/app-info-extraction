import pandas as pd
import numpy
import os
import msvcrt

DELETE_HEADER = ["Material", "Type", "Rev"]
HEADER = ["Address", "Row_Num", "Col_Num"]
DEBUG = False
BCBDBE = [54, 55, 56]


def deleteBCBDBE(file):
    '''
    删除物料三列的内容
    '''
    df = pd.read_excel(file, keep_default_na=True).iloc[:, : 67]
    # print(df)
    new_df = df.drop('Material', axis=1)
    new_df.drop('Type', axis=1, inplace=True)
    new_df.drop('Rev', axis=1, inplace=True)

    new_df.insert(loc=54, column="Material", value=numpy.nan)
    new_df.insert(loc=55, column="Type", value=numpy.nan)
    new_df.insert(loc=56, column="Rev", value=numpy.nan)
    new_df.to_excel(file, na_rep='', sheet_name='Sheet1', index=False)


def checkGeneralOP(file):
    '''
    检查若工步为通用工步，后面两列有没有数据，若没有则报错
                       两列后不允许有数据，若有则需报错
    '''
    df = pd.read_excel(file, keep_default_na=True).iloc[:, 21: 67]
    ar = df.to_numpy()
    row = 2
    row_list = []
    flag = True
    for i in ar:
        if i[0] == '通用工步':
            if i[1] != i[1] or i[2] != i[2]:
                # check if 通用工步紧接后面两列有没有数据
                # 根据nan != nan的方法测试是否是nan
                flag = False
                row_list.append(row)
                row += 1
                continue

            for j in i[3:]:
                if j == j:
                    flag = False
                    row_list.append(row)
                    continue
        row += 1
    return flag, list(set(row_list))


def readExcel(file_name, start_col, end_col):
    df = pd.read_excel(file_name, keep_default_na=True).iloc[:, start_col:end_col]
    return df


def removeFirstLines(file):
    '''
    检查有没有备注的两行:title + 备注,有则删除
    '''
    df = pd.read_excel(file, keep_default_na=True).iloc[:, : 67]

    # print(df)
    ar = df.to_numpy()

    if "1.Run" in ar[0][0] and "Product" in ar[1][0]:
        df.drop(index=[0, 1], inplace=True)
        df.to_excel(file, na_rep='', sheet_name='Sheet1', index=False)
    elif "1.Run" in ar[0][0]:
        df.drop(index=[0, 1], inplace=True)
        df.to_excel(file, na_rep='', sheet_name='Sheet1', index=False)
    elif "Product" in ar[0][0]:
        df.drop(index=[0], inplace=True)
        df.to_excel(file, na_rep='', sheet_name='Sheet1', index=False)


def checkPartNum(file):
    '''
    检查件号是否一致
    True = 件号数量为一
    False = 件号数量异常
    '''
    df = readExcel(file, 1, 2)
    ar = df.to_numpy()
    s = set()
    for i in ar:
        s.add(i[0])

    if DEBUG:
        print(s)
    return len(s) == 1


def checkTU(file):
    '''
    检查TU两列是否有数据，且同一行数据只能有一列存在
    '''
    flag = True
    df = readExcel(file, 19, 21)
    ar = df.to_numpy()
    row_list = []
    row = 2
    for pair in ar:
        if pair[0] == "Op":
            row += 1
            continue
        if pair[0] == pair[0] and pair[1] == pair[1]:
            # print(pair[0])
            flag = False
            row_list.append(row)
        if pair[0] != pair[0] and pair[1] != pair[1]:
            flag = False
            row_list.append(row)
        row += 1
    return flag, row_list


def checkV(file):
    '''
    检查V列是否有数据，且必须为“专用工步”或“通用工步”
    '''
    flag = True
    df = readExcel(file, 21, 22)
    ar = df.to_numpy()
    row_list = []
    row = 2
    for i in ar:
        # print(i[0])
        if i[0] != "专用工步" and i[0] != "通用工步":
            flag = False
            row_list.append(row)
        row += 1
    return flag, row_list


def checkS(file):
    """
    检查S列的工序号是否从10开始且以10递进
    """
    df = readExcel(file, 18, 19)
    ar = df.to_numpy()
    temp = 0
    row = 2
    row_list = []
    flag = True
    for i in ar:

        if i[0] == temp + 10:
            temp += 10
            row += 1
            continue
        else:
            flag = False
            row_list.append(row)
            row += 1
    return flag, row_list


def checkFile(path):
    partNum_error = []
    TU_error = []
    V_error = []
    S_error = []
    GOP_error = []
    error_list = []
    exceptionList = []

    for root_dir, sub_dir, files in os.walk(path):
        for file in files:
            if file.endswith('xlsx'):
                flag = False
                print(f">CHECKING FILE: {file}")
                file_name = os.path.join(root_dir, file)
                if DEBUG:
                    removeFirstLines(file_name)

                    deleteBCBDBE(file_name)

                    GOP_status, GOP_row = checkGeneralOP(file_name)

                    partNum_status = checkPartNum(file_name)

                    TU_status, TU_row = checkTU(file_name)

                    V_status, V_row = checkV(file_name)

                    S_status, S_row = checkS(file_name)

                    if partNum_status == False:
                        error_list.append((file_name, "N/A", "B"))
                        partNum_error.append(file_name)

                    if TU_status == False:
                        error_list.append((file_name, TU_row, "TU"))
                        TU_error.append((file_name, TU_row, "TU"))

                    if V_status == False:
                        V_error.append((file_name, V_row, "V"))
                        error_list.append((file_name, V_row, "V"))

                    if S_status == False:
                        S_error.append((file_name, S_row, "S"))
                        error_list.append((file_name, S_row, "S"))

                    if GOP_status == False:
                        GOP_error.append((file_name, GOP_row, "通用工步后有数据"))
                        error_list.append((file_name, GOP_row, "通用工步后有数据"))
                else:
                    try:
                        removeFirstLines(file_name)

                        deleteBCBDBE(file_name)

                        partNum_status = checkPartNum(file_name)

                        GOP_status, GOP_row = checkGeneralOP(file_name)

                        TU_status, TU_row = checkTU(file_name)

                        V_status, V_row = checkV(file_name)

                        S_status, S_row = checkS(file_name)

                        if partNum_status == False:
                            flag = True
                            error_list.append((file_name, "N/A", "B"))
                            partNum_error.append(file_name)

                        if TU_status == False:
                            flag = True
                            error_list.append((file_name, TU_row, "TU"))
                            TU_error.append((file_name, TU_row, "TU"))

                        if V_status == False:
                            flag = True
                            V_error.append((file_name, V_row, "V"))
                            error_list.append((file_name, V_row, "V"))

                        if S_status == False:
                            flag = True
                            S_error.append((file_name, S_row, "S"))
                            error_list.append((file_name, S_row, "S"))

                        if GOP_status == False:
                            flag = True
                            GOP_error.append((file_name, GOP_row, "通用工步后有数据"))
                            error_list.append((file_name, GOP_row, "通用工步后有数据"))
                        # if not flag:
                        # print("PASS")
                    except:
                        exceptionList.append(file_name)

    return partNum_error, TU_error, V_error, S_error, GOP_error, error_list, exceptionList


if __name__ == "__main__":
    dirc = os.getcwd()  # working dirctory
    partNum_error, TU_error, V_error, S_error, GOP_error, error_list, exceptionList = checkFile(dirc)
    out_path = os.path.join(dirc, 'error_list.xlsx')
    df = pd.DataFrame(error_list, columns=HEADER)
    # print(df)
    df.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False, header=HEADER)
    print("+++++++++++++++++++++++++++++++++++++++++++++")
    if len(error_list) < 10:
        if len(partNum_error) > 0:
            print(f"件号错误: ")
            for i in partNum_error:
                print(i)

        if len(TU_error) > 0:
            print("+++++++++++++++++++++++++++++++++++++++++++++")
            print(f"TU列错误: ")
            for i in TU_error:
                print(i)

        if len(V_error) > 0:
            print("+++++++++++++++++++++++++++++++++++++++++++++")
            print(f"V列错误: ")
            for i in V_error:
                print(i)

        if len(S_error) > 0:
            print("+++++++++++++++++++++++++++++++++++++++++++++")
            print(f"S列错误: ")
            for i in S_error:
                print(i)
        if len(GOP_error) > 0:
            print("+++++++++++++++++++++++++++++++++++++++++++++")
            print(f"通用工步错误: ")
            for i in GOP_error:
                print(i)

    if len(error_list) > 0:
        print(f"校验数据输出到 {out_path}")
    if len(exceptionList) > 0:
        print("+++++++++++++++++++++++++++++++++++++++++++++")
        print(f"下列文件读取出错")
        for i in exceptionList:
            print(i)
        print()
        out_path = os.path.join(dirc, 'exception_list.xlsx')
        df = pd.DataFrame(exceptionList, columns=["地址"])
        df.to_excel(out_path, na_rep='', sheet_name='Sheet1', index=False)
    os.system('pause')