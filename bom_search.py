#coding=utf-8
import os
import pandas as pd
import re

class Bomsearch(object):
    def __init__(self,model):
        self.model=model
        self.ar=[]
        self.rootdir=u'J:\\PIE Process Manual\\新工序文件(CMP)\\工序手冊\\{}\\'.format(model)

    #获取目录下xls文件
    def get_file(self,rootdir):
        ar=[]
        x=os.listdir(rootdir)
        for fil in x:
            if os.path.splitext(fil)[1]!='.pdf':
                ar.append(fil)
        return ar

    #获取xls内容
    def openxls(self,rootdir,xlsworkbook):
        path=os.path.join(rootdir,xlsworkbook)
        xls=pd.ExcelFile(path)
        data=pd.read_excel(xls,sheetname=None,header=None,index=None,encoding='gb18030')
        for v in data.values():
            v.dropna(axis=0,how='all',inplace=True)
            v.insert(0,'motor_model',xlsworkbook[:-4])  #插入一列，并剔除文件后綴
            v.fillna(method='pad',inplace=True)
            v.fillna(method='bfill',inplace=True)
            v.fillna(value='Null',axis=0,inplace=True)
            v.to_csv('t.csv',mode='a',header=None,index=None,encoding='gb18030')

if __name__=='__main__':
    model=raw_input('Enter model series:>> ')
    fil=Bomsearch(model)
    files=fil.rootdir
    ar=fil.get_file(files)
    for a in ar:
        print 'appending>>',a
        try:
            fil.openxls(files,a)
        except:
            with open('error.txt','a') as f:
                f.write(a)

    #fil.openxls(files,'UK-30R-0005.xls')
'''
#d=v[v[1].str.contains(u'罩').astype(bool)]
    for a in ar:
        t=multiprocessing.Process(target=fil.openxls,args=(a,))
        th.append(t)
    for thread in th:
        thread.start()
    for thread in th:
        thread.join()

import os
import pandas as pd
df=pd.DataFrame({'one':[u'我的',u'中国',1,12,13,24,5,6],'two':[u'安徽',7,8,9,10,11,12,13]},index=None)
#print df
d=df['one'].str.contains(u'的')  #模糊匹配
'''