#coding:utf-8
#此程序功能，将统一文件夹下部员课表(pdf格式)先转成0-1excel空闲表，再合成为一个txt文件
#0.08版本：0.仍要用现在的排班软件(新版)，此程序只是个提取pdf生成0-1空闲表的工具
#         1.仅适用于本科生课表(研究生博士生课表有奇怪读取问题，读出来全是None)
#         2.所用课表必须为直接从教学信息网直接导出pdf，不能用截图再转pdf的形式
#使用说明：1.做好文件管理，新建一个文件夹，其下再新建两个文件夹命名为'PDF'和'EXCEL',此外也可存放其它文件和文件夹，不影响程序运行
#        2.将且仅将部员课表(pdf格式)放入PDF文件夹下
#        3.修改主函数中的folder_name路径为1中文件夹路径
#        4.运行后将在EXCEL文件夹中生成对应的0-1excel空闲表

import numpy as np
import pandas as pd
import os
import openpyxl
import pdfplumber # 导入pdfplumber(读取pdf的工具包)

#将pdf课表转化成excel格式的0-1空闲表(0:空闲 ; 1:有课)
def pdf_01excel(path,Path_EXCEL,filename):#path:文件路径 Path_EXCEL:存放转化后的文件夹根目录路径 filename:不带扩展名的文件名(即部员姓名)
    # 读取pdf文件，保存为pdf实例
    pdf = pdfplumber.open(path)
    #print(pdf)
    PageNumber=len(pdf.pages)

    # 访问pdf所有页并拼接在一起
    first_page = pdf.pages[0]
    all = np.array(first_page.extract_table())
    for i in range (1,PageNumber):
        now_page = pdf.pages[i]
        now_table = np.array(now_page.extract_table())
        if (now_page.extract_table()==None):
            break
        else:
            all = np.concatenate((all, now_table))
    all_valid = all[2:,1:9]#删去前两行，列只保留课程节数和周一到周日的课

    #清洗表格，删去无效行
    #print(len(all_valid))#得出数组长度
    all_valid_final=all_valid#构造出一个第二维度与原数组相同的空数组(即两数组的一维数组的大小相同)
    times=0
    for j in range(len(all_valid)):
        if all_valid[j,0]==''or all_valid[j,0]==None:
            all_valid_final=np.concatenate((all_valid_final[:(j-times)],all_valid_final[(j-times+1):]))
            times=times+1

    os.chdir(Path_EXCEL)  # 修改工作路径至存放转化后的excel的文件夹
    workbook = openpyxl.Workbook()#新建xlsx文件
    sheet = workbook.active # 选中活动工作表
    sheet['A1']=filename
    for p in range(6):
        for q in range(7):
            if all_valid_final[p*2][q+1]=='' and all_valid_final[p*2+1][q+1]=='':
                col_index=chr(ord('A')+q)   #列索引
                #print(col_index)
                sheet[col_index+str(p+2)]=0
            else:
                col_index = chr(ord('A') + q)  # 列索引   遍历字母方法：先转ascii码遍历再转回字符
                sheet[col_index+str(p+2)]=1
    sheet.title = filename
    workbook.save(filename+'.xlsx')#另存为(实际上是重命名)



if __name__ == '__main__':#如果要调用则需要删去
    folder_name=r'C:\Users\86198\Desktop\convert'
    folder_PDF_name = os.path.join(folder_name,'PDF')  # 存放课表的文件夹根目录
    folder_EXCEL_name = os.path.join(folder_name,'EXCEL') #转换后的存放excel的文件夹根目录
    list_path = os.listdir(folder_PDF_name)  # 读取文件夹里面的全部文件名
    # print(len(list_path))

    for index in list_path:  #list_path返回的是一个列表   通过for循环遍历提取元素
        name = index.split('.')[0]   #split字符串分割的方法 , 分割之后是返回的列表 索引取第一个元素[0],分离出姓名
        print(index.split('.')[0])
        path = os.path.join(folder_PDF_name,index) #路径拼接，此处必须用完整路径，是因为不在明确的根目录下，不能直接用文件名
        #print(path)
        filename = name
        pdf_01excel(path,folder_EXCEL_name,filename)

print('处理完成')
