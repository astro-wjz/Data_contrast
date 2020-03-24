# 以身份证号码为基准对比不同工作表数据，并将不重复数据输出
# --* coding: UTF-8 *--
# 作者：王荐钊  2020.03.05

# 说明：
# 1. 对数据格式的要求：身份证号码列列名必须为“身份证号码”，姓名列列名必须为“姓名”
# 2. 数据文件名必须以数字开头，文件后缀名必须为 .xls 或 .xlsx
# 3. 请将要处理的文件和此程序文件一起放入同一个文件夹，并保证其中没有其他的Excel文件
# 4. 此程序只能处理2个文件，若文件夹内有多个文件则只能处理前两个文件

import pretty_errors
import pandas as pd 
import os
import xlrd,openpyxl
# import win32api,win32con

# 获取要处理的文件并读取数据
# 身份证号码存放在 idNum_i 中，姓名存放在 prName_i 中
##确认信息框
# win32api.MessageBox(0, '请只保留要对比的两个数据文件和程序文件，移除其他的数据文件', "提醒",win32con.MB_OKCANCEL)
file_dir = os.path.dirname(os.path.realpath(__file__))
file_names = os.listdir(file_dir)
filenames = []
file_num = 0
for filename in file_names:
    if os.path.splitext(filename)[1] == '.xls' or os.path.splitext(filename)[1] == '.xlsx':
        filenames.append(filename)
        file_num = file_num + 1
filenames.sort(key=lambda x:int(x[:6]))
if(file_num < 2):
    print('可处理文件不足，请重试')
if(file_num > 1):
    print('数据读取中...')
    path_temp = filenames[0]
    data1 = pd.read_excel(os.path.join(file_dir, path_temp))
    path_temp = filenames[1]
    data2 = pd.read_excel(os.path.join(file_dir, path_temp))
    idNum_1 = data1['身份证号码']
    prName_1 = data1['姓名']
    idNum_2 = data2['身份证号码']
    prName_2 = data2['姓名']
    if(file_num > 2):
        print('处理文件太多，请选择不多于2个文件处理')
print('待处理的文件有：', '\n', filenames)
print('文件总数：', file_num)

# 处理数据
print('数据处理中，请稍等...')
nochg_index = []
del_index = []
add_index = []
line_num = 0
for i in idNum_1.values:
    if i in idNum_2.values:
        nochg_index.append(line_num)
    else:
        del_index.append(line_num)
    line_num = line_num + 1
line_num = 0
for i in idNum_2.values:
    # temp_num = data2[(data1['身份证号码']==i)].index.tolist()[0]
    if i not in idNum_1.values:
        add_index.append(line_num)
    line_num = line_num + 1

# 输出数据
new_added = data2[add_index[0]:add_index[0]+1]
for i in add_index[1:]:
    newData = data2[i:i+1]
    new_added = new_added.append(newData)
new_added.to_excel('新增人员统计.xls', index=False)

new_del = data1[del_index[0]:del_index[0]+1]
for i in del_index[1:]:
    newData = data1[i:i+1]
    new_del = new_del.append(newData)
new_del.to_excel('删减人员统计.xls')

new_nochg = data1[nochg_index[0]:nochg_index[0]+1]
for i in nochg_index[1:]:
    newData = data1[i:i+1]
    new_nochg = new_nochg.append(newData)
new_nochg.to_excel('未变动人员统计.xls')

print('处理完成！')
print('未变动数据共', len(nochg_index), '条，', end='')
print('新增数据', len(add_index), '条，', end='')
print('删减数据', len(del_index), '条.')

mylog = open('recode.log', mode = 'a',encoding='utf-8')
print('未变动数据共', len(nochg_index), '条，', end='', file=mylog)
print('新增数据', len(add_index), '条，', end='', file=mylog)
print('删减数据', len(del_index), '条.', file=mylog)
mylog.close()
# print(idNum_1[1:2], prName_1[1:2])
