# coding:utf-8
import pandas
import os
#此程序功能:将多个excel写入一个txt文件(但不符合现有软件的格式要求)


# 将excel写入该txt文件
def excel_into_txt(input_path,output_path):#input_path和output_path既可以用完整路径，在明确的根目录下用文件名也可
    df = pandas.read_excel(input_path, header=None)
    df.to_csv(output_path, mode='a', float_format='%.0f', header=None, sep=' ', index=False)  # sep指定分隔符，分隔单元格;mode='a'/'w':追加模式/写入模式

if __name__ == '__main__':
    folder_name=r'C:\Users\86198\Desktop\convert\EXCEL'
    os.chdir(folder_name)  # 修改工作路径
    filenames = os.listdir(folder_name)  # 读取文件夹里面的全部文件名
    #新建txt文件
    txt = open('schedule.txt', 'w')#如有该文件则打开文件，如无该文件则会新建一个名为此的文件
    for filename in filenames:
        excel_into_txt(filename,txt)
