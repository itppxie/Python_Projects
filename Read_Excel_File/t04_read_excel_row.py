import pandas as pd

def read_excel_row_ppxie(excel_file_path, excel_sheet_name, row_num, start_column, end_column):
    # todo：从Excel文件中，读取指定的表单sheet
    excel_file = pd.read_excel(excel_file_path, excel_sheet_name)
    print('打印excel_file的数据类型', type(excel_file))  # <class 'pandas.core.frame.DataFrame'>
    print('打印excel_file：\n', excel_file)
    print('\n\n')
    # todo: 获取Excel中指定行的内容，如第6行的所有单元格内容
    row_data = excel_file.iloc[(row_num - 2), :] # 切片，row_data直接存放了第6行中所有列的单元格内容
    #
    print('打印row_data的数据类型（里面存放了第6行中的所有列单元格内容）：', type(row_data)) # <class 'pandas.core.series.Series'>
    print('打印row_data中的所有单元格数据：\n', row_data) # 里面存放了第6行中的所有列单元格内容
    print('打印row_data中的单元格个数：', len(row_data)) # 10
    print('\n\n')

    # 两种方式遍历Excel里面 该行的所有列的单元格
    print('以下是遍历该行的所有单元格内容（索引遍历）：')
    for index in range(0, len(row_data), 1):
        cell_data = row_data[index]
        print('打印row_data中的单元格cell_data数据类型：', type(cell_data), '在原始Excel文件中的列号：', (index + 1))
    print('\n\n')
    print('以下是遍历该行的所有单元格内容（直接遍历）：')
    for cell_data in row_data:
        print(cell_data)
        print('打印row_data中的单元格cell_data数据类型：', type(cell_data))
    print('\n\n')

    # 索引遍历Excel里面第6行中从第3列至第8列的所有单元格
    print('以下是遍历第6行的从第3列至第8列的所有单元格：')
    for index in range((start_column - 1), (end_column - 1) + 1, 1):
        cell_data = row_data[index]
        print(cell_data)
        print('打印row_data中的单元格cell_data数据类型：', type(cell_data), '在原始Excel文件中的列号：', (index + 1))





# test
if __name__ == '__main__':
    pass
    excel_file_path = './other_files/test_excel_2.xlsx'
    excel_sheet_name = 'Sheet2'
    row_num = 6 # Excel文件中的6行
    start_column, end_column = 3, 8 # Excel文件中的开始列和结束列
    read_excel_row_ppxie(excel_file_path, excel_sheet_name, row_num, start_column, end_column)
