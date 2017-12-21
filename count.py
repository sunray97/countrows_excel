import xlrd
import os
import sys
# rootdir = 'D:/工作/code/electric/'
rootdir = sys.argv[1]
xlrd.Book.encoding = "gbk"
sumnum=0
filenum = 0
list = os.listdir(rootdir) #列出文件夹下所有的目录与文件
for i in range(0,len(list)):
    path = os.path.join(rootdir,list[i])
    if os.path.isfile(path):
        print('正在处理：'+path)
        data = xlrd.open_workbook(path)
        table = data.sheet_by_index(0)
        # table = data.sheet_by_name(u'Sheet1')
        nrows = table.nrows
        data.release_resources()
        sumnum=sumnum+nrows
        filenum=filenum+1
        print('-------------------------------------------------------------------------')
print('共有%d个文件'%filenum)
print('共有%d行记录'%sumnum)
