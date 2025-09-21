import os
import pandas as pd
import openpyxl

########### todo: 将大写英文字母char转换成对应的数字，如'A'转换成1，'C'转换成3
def alphabet2number(char):
    result = ord(char) - ord('A') + 1 # ord('A')将字母“A”转换成ascii 65
    return result

# test
if __name__ == '__main__':
    excel_file_path = './other_files/test_excel.xlsx'
    excel_sheet_name = 'Sheet2'
    column_char = 'J'

    # todo：从Excel文件中，读取指定的表单sheet
    excel_file = pd.read_excel(excel_file_path, excel_sheet_name)
    print('打印excel_file的数据类型', type(excel_file))  # <class 'pandas.core.frame.DataFrame'>
    print('打印excel_file：\n', excel_file)
    print('\n\n')
    # todo: 获取Excel中指定列的内容，如'J'列内容
    # 从表单中，获取指定列（这里是‘J’列）的所有单元格内容
    num = alphabet2number(column_char)  # 获取大写字母的序号，如'A'就是1，'C'就是3
    column_data = excel_file.iloc[:, num - 1]  # column_data中存放了指定列的所有单元格内容
    print('打印column_data的数据类型（里面存放了所有的列单元格数据）：', type(column_data))  # <class 'pandas.core.series.Series'>
    print('打印column_data中的所有单元格数据：\n', column_data)
    print('打印column_data中的单元格个数：', len(column_data))  # 注意，仅有9个，而不是10个！！！

    print('\n\n')

    # 方式一：直接遍历 该列中的所有单元格内容
    for cell_data in column_data:
        print(cell_data)
        print('打印column_data中的单元格cell_data数据类型：', type(cell_data))

    print('\n\n')

    # 方式二：索引遍历 该列中的所有单元格内容
    for index in range(0, len(column_data), 1):
        print(column_data[index])
        print('打印column_data中的单元格数据类型：', type(column_data[index]))

    print('\n\n')
    print(type(column_data[0]))
    print(column_data[0])

    print('\n\n')

    # 获取Excel文件中指定列的第3行至第8行的数据内容
    start_row = 3
    end_row = 8
    start_row_truth = start_row - 2 # 为什么要减去2，因为Excel表格中的第一行是列名称，且行也是从索引0开始的！
    end_row_truth = end_row - 2
    for row in range(start_row_truth, end_row_truth + 1, 1):
        print('打印Excel中指定列的第', row + 2, '行的内容：', column_data[row])

