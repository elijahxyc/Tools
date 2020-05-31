""" 
2020.05.22
用来计算表格中的业务数据也总数
"""

import xlrd
import xlwt
from decimal import Decimal

function_name = []
function_typeOne = "业务功能"
function_typeTwo = "IT功能"
dict_function = {}
sum_all = Decimal(0.00)



def readFile():
    # 初始化
    wb = xlrd.open_workbook("sum.xlsx") #打开文件
    data_sheet = wb.sheet_by_index(1)
    nrows = data_sheet.nrows #列数
    ncols = data_sheet.ncols #行数

    function_value = data_sheet.col_values(5, start_rowx=2, end_rowx = None) #功能列
    sumOneTwo_value = data_sheet.col_values(8, start_rowx=2, end_rowx= None) #业务/IT功能
    sumThree_value = data_sheet.col_values(17, start_rowx=2, end_rowx= None) #功能点数

    for x in function_value:
        if x not in function_name:
            function_name.append(x)
    
    #for x in sumThree_value:
    #   print(x)

    
    for i in range(len(function_name) - 1):
        sumOne = 0 # 计算业务功能点数
        sumTwo = 0 # 计算 IT 功能点数
        sumThree = 0.00 # 计算总功能点数
        for j in range(len(function_value) - 1):    
            if function_name[i] == function_value[j] and function_value[j] == function_value[j+1]:
                sumThree = Decimal(sumThree) + Decimal(sumThree_value[j])
                if sumOneTwo_value[j] == function_typeOne:
                    sumOne = sumOne + 1
                elif sumOneTwo_value[j] == function_typeTwo:
                    sumTwo = sumTwo + 1
            elif function_name[i] == function_value[j] and function_value[j] != function_value[j+1]:
                sumThree = Decimal(sumThree) + Decimal(sumThree_value[j])
                if sumOneTwo_value[j] == function_typeOne:
                    sumOne = sumOne + 1
                elif sumOneTwo_value[j] == function_typeTwo:
                    sumTwo = sumTwo + 1
                #sum_all = sum_all + sumThree
                dict_function[function_name[i]] = str(sumOne) + "_"  + str(sumTwo) + "_" + str(round(sumThree,2))
    
    for k in dict_function.keys():
        print(k)
        #print(k + ":" + dict_function[k])

def readFile2():
    
 
    sum_value = data_sheet.col_values(0, start_rowx=0, end_rowx= None) #功能点数

    sum = 0.00

    for x in sum_value:
        sum = sum + x
    
    print(sum)



if __name__ == '__main__':
    readFile()
    #readFile2()
