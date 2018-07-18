#coding=utf-8
import os
import pandas as pd
import re
from multiprocessing import Process,Pool

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
def openxls(rootdir,xlsworkbook):
    print 'Appending >>> ',xlsworkbook
    path=os.path.join(rootdir,xlsworkbook)
    xls=pd.ExcelFile(path)
    data=pd.read_excel(xls,sheetname=None,header=None,index=None,encoding='gb18030')
    for v in data.values():
        v.dropna(axis=0,how='all',inplace=True)
        v.insert(0,'motor_model',xlsworkbook[:-4])  #插入一列，并剔除文件后綴
        v.fillna(method='pad',inplace=True)
        v.fillna(method='bfill',inplace=True)
        v.fillna(value='Null',axis=0,inplace=True)
        v.to_csv('results.csv',mode='a',header=None,index=None,encoding='gb18030')

if __name__=='__main__':
    model=raw_input('Enter model series:>> ')
    fil=Bomsearch(model)
    files=fil.rootdir
    ar=fil.get_file(files)
    p=Pool(4)
    th=[]
    for a in ar:
        p.apply(openxls,args=(files,a,))
    p.close()
    p.join()