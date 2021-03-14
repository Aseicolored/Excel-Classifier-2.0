# coding:utf-8
import os
import xlwings

def adder(nul,nun):
    return nul + str(nun)

print("Excel-Classifier 2.0")
print("")

source_path = input("Source/Path: ") #输入工作簿
output_path = input("Output/Path: ") #输出工作簿

wbs = xlwings.Book(source_path)
wbo = xlwings.Book(output_path)

sts = input("Worksheet/Source: ")#输入工作表
sto = input("Worksheet/Output: ")#输出工作表

shts = wbs.sheets[sts]#实例化输入工作表
shto = wbo.sheets[sto]#实例化输出工作表

col_index = input("Source/Ground: ")#判断依据所在列

col_left = input("Column/Left: ")#范围的最左列
col_right = input("Column/Right: ")#范围的最右列

row_begin_source = int(input("Source/Row/Begin: "))#输入的起始行(包括此行)
row_end_source = int(input("Source/Row/End: "))#输入的结束行(包括此行)

row_begin_output = int(input("Output/Row/Begin: "))#输出的起始行(包括此行)
row_end_output = int(input("Output/Row/End: "))#输出的结束行(包括此行)

row_begin_source_loop = row_begin_source#外层循环
row_end_source_loop = row_end_source#外层循环

row_begin_output_loop = row_begin_output#内层循环
row_end_output_loop = row_end_output#内层循环

print("")
times = 1

for i in range(row_begin_source_loop , row_end_source_loop + 1):#在输入文件中遍历
    index_ground_source = adder(col_index , row_begin_source)#获取判断依据在输入文件中的坐标
    ground_source = shts.range(index_ground_source).value#获取在输入文件中的判断依据

    row_begin_output_process = row_begin_output#重置内循环的坐标
    for j in range(row_begin_output_loop , row_end_output_loop + 1):#在输出文件中遍历
        index_ground_output = adder(col_index , row_begin_output_process)#获取判断依据在输出文件中的坐标
        ground_output = shto.range(index_ground_output).value#获取在输出文件中的判断依据

        if ground_output == ground_source:
            print("[Time " , end = '')
            print(times , end == '')
            print("]")
            
            output_range_left = adder(col_left , row_begin_output_process)
            output_range_right = adder(col_right , row_begin_output_process)
            output_range = output_range_left + ":" + output_range_right#获得输出范围
            print("[Output Range] " , end = '')
            print(output_range)

            source_range_left = adder(col_left , row_begin_source)
            source_range_right = adder(col_right , row_begin_source)
            source_range = source_range_left + ":" + source_range_right#获得输入范围
            print("[Source range] " , end = '')
            print(source_range)
            print("")

            shto.range(output_range).value = shts.range(source_range).value#输出数据
            times = times + 1
            break

        row_begin_output_process = row_begin_output_process + 1#输出文件的换行

    row_begin_source = row_begin_source + 1#输入文件的换行
print("")
print("Finish!")
