# coding:utf-8
import pandas as pd
import os
import openpyxl
#此程序功能:把不能导入的txt写入一个excel里，再把excel粘贴到另一个txt中则转化为可导入的格式(实测有效)
path=r'C:\Users\86198\Desktop\convert\EXCEL\schedule.txt'
Path=r'C:\Users\86198\Desktop\convert\EXCEL'
os.chdir(Path)
workbook=openpyxl.Workbook()
sheet = workbook.active  # 选中活动工作表
sheet.title = '01空闲表'
workbook.save('schedule.xlsx')  # 另存为(实际上是重命名)
new_path=os.path.join(Path,'schedule.xlsx')
test1 = pd.read_table(path,sep=' ',index_col=None,header=None)
#print(test1)
df = pd.DataFrame(test1)
print(df)
a=df.to_excel(new_path, startcol=0,index=None)

