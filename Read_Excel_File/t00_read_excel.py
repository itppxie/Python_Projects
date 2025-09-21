import pandas as pd


########### todo: 将大写英文字母char转换成对应的数字，如'A'转换成1，'C'转换成3
def alphabet2number(char):
    result = ord(char) - ord('A') + 1 # ord('A')将字母“A”转换成ascii 65
    return result


# Excel文件路径
excel_file_path = './other_files/test_excel.xlsx'
excel_sheet_name = 'Sheet2'

# todo 1：读取Excel文件中【第3行第“J”列】单元格内容
excel_row_num = 3
excel_column_num = alphabet2number('J')
# 读取Excel文件
excel_file = pd.read_excel(excel_file_path, excel_sheet_name) # 数据类型：<class 'pandas.core.frame.DataFrame'>
cell_data = excel_file.iloc[excel_row_num - 2, excel_column_num - 1] # 一个元素即为单元格内容，数据类型<class 'str'>
print('cell_data的长度：', len(cell_data)) # 2，即为‘李四’
print('cell_data的数据类型：', type(cell_data)) # <class 'str'>
print('打印cell_data：', cell_data) # ‘李四’



# test
if __name__ == '__main__':
    pass
