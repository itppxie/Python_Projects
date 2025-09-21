import pandas as pd

def read_excel_row_ppxie(excel_file_path, excel_sheet_name, row_num, start_column, end_column):
    # todo：从Excel文件中，读取指定的表单sheet
    excel_file = pd.read_excel(excel_file_path, excel_sheet_name)
    print('打印excel_file的数据类型', type(excel_file))  # <class 'pandas.core.frame.DataFrame'>
    print('打印excel_file：\n', excel_file)
    print('\n\n')
    # todo: 获取Excel中指定行的内容，如第6行的所有单元格内容
    row_data = excel_file.iloc[(row_num - 2), (start_column - 1):((end_column - 1) + 1)] # 切片，row_data直接存放了第6行中从第3列至第8列的所有单元格内容
    #
    print('打印row_data的数据类型（里面存放了第6行中从第3列至第8列的所有单元格内容）：', type(row_data)) # <class 'pandas.core.series.Series'>
    print('打印row_data中的所有单元格数据：\n', row_data) # 里面存放了第6行中从第3列至第8列的所有单元格内容
    print('打印row_data中的单元格个数：', len(row_data)) # 6
    print('\n\n')

    # 直接遍历该行的从第3列至第8列的所有单元格
    print('以下是遍历该行的从第3列至第8列的所有单元格：')
    for cell_data in row_data:
        print(cell_data)
        print('打印column_data中的单元格cell_data数据类型：', type(cell_data))





# test
if __name__ == '__main__':
    pass
    excel_file_path = './other_files/test_excel_2.xlsx'
    excel_sheet_name = 'Sheet2'
    row_num = 6 # Excel文件中的6行
    start_column, end_column = 3, 8 # Excel文件中的开始列和结束列
    read_excel_row_ppxie(excel_file_path, excel_sheet_name, row_num, start_column, end_column)
