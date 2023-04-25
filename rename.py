# coding:utf-8
import pandas
import os
import openpyxl
#此程序功能:根据部门管家名单重命名pdf文件()


# 将excel写入该txt文件
def rename(filename):#同一文件夹下文件名有唯一性，将文件名作为参数传入也问题不大
        src_path = os.path.join(folder_name, filename)  # 路径拼接
        name = filename.split('.')[0]  # split字符串分割的方法 , 分割之后是返回的列表 索引取第一个元素[0],分离出姓名
        for i in range(1,sheet.max_row):
            cell = sheet["A%d" % i].value  # 从excel读取数据
            if cell in name :
                print(name)
                dst_path=os.path.join(folder_name, cell)
                dst_path=dst_path+'.pdf'
                os.rename(src_path,dst_path)


if __name__ == '__main__':
    folder_name=r'C:\Users\86198\Desktop\convert\PDF'
    path = r'C:\Users\86198\Desktop\convert\department manager.xlsx'  # 改成部门管家的excel路径
    workbook = openpyxl.load_workbook(path)  # 返回一个workbook数据类型的值
    sheet = workbook.active  # 获取活动表
    os.chdir(folder_name)  # 修改工作路径
    filenames = os.listdir(folder_name)  # 读取文件夹里面的全部文件名
    for filename in filenames:
        rename(filename)
    print('修改完成')
