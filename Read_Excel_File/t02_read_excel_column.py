import os
import pandas as pd
import openpyxl


########### todo: 将大写英文字母char转换成对应的数字，如'A'转换成1，'C'转换成3
def alphabet2number(char):
    result = ord(char) - ord('A') + 1 # ord('A')将字母“A”转换成ascii 65
    return result

##########todo：读取Excel文件里面"J"列中从第3行至第8行的单元格内容
def read_excel_column_ppxie(excel_file_path, excel_sheet_name, column_char, start_row, end_row):
    # todo：从Excel文件中，读取指定的表单sheet
    excel_file = pd.read_excel(excel_file_path, excel_sheet_name)
    print('打印excel_file的数据类型', type(excel_file)) # <class 'pandas.core.frame.DataFrame'>
    print('打印excel_file：\n', excel_file)
    print('\n\n')
    # todo: 获取Excel中指定列的内容，如第'J'列内容
    num = alphabet2number(column_char)  # 获取大写字母的序号，如'A'就是1，'C'就是3
    column_data = excel_file.iloc[:, num - 1] # 切片，column_data里面存放了"J"列中所有行的单元格内容
    # 打印
    print('打印column_data的数据类型（里面存放了“J”列中所有行的单元格数据）：', type(column_data)) # <class 'pandas.core.series.Series'>
    print('打印column_data中的所有单元格数据：\n', column_data) # column_data里面存放了“J”列中所有行的单元格数据
    print('打印column_data中的单元格个数：', len(column_data)) # 9
    print('\n\n')

    # 两种方式遍历Excel里面 该列中的所有单元格内容
    print('以下是遍历该列的所有单元格内容（索引遍历）：')
    for index in range(0, len(column_data), 1):
        cell_data = column_data[index]
        print(cell_data)
        print('打印column_data中的单元格cell_data数据类型：', type(cell_data), '在原始Excel文件中的行号：', (index + 2))
    print('\n\n')
    print('以下是遍历该列的所有单元格内容（直接遍历）：')
    for cell_data in column_data:
        print(cell_data)
        print('打印column_data中的单元格cell_data数据类型：', type(cell_data))
    print('\n\n')


    # 索引遍历Excel里面“J”列中第3行至第8行的所有单元格内容
    print('以下是遍历J列中第3行至第8行的所有单元格内容：')
    for index in range(start_row - 2, (end_row -2) + 1, 1):
        cell_data = column_data[index]
        print(cell_data)
        print('打印column_data中的单元格cell_data数据类型：', type(cell_data), '在原始Excel文件中的行号：', (index + 2))




# test
if __name__ == '__main__':
    excel_file_path = './other_files/test_excel.xlsx'
    excel_sheet_name = 'Sheet2'
    column_char = 'J'
    start_row, end_row = 3, 8 # Excel文件中的开始行和结束行
    read_excel_column_ppxie(excel_file_path, excel_sheet_name, column_char, start_row, end_row)