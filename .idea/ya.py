import os
import re
import xlrd
import datetime
import getopt
import sys
from collections import OrderedDict


os.chdir("/Users/Andy/Desktop/Hpertension-project/auto-script of hpertension")

D = {}  #药物分类和（基因型、遗传评估）的关系   地尔硫䓬 --> {'TT': '患者使用地尔卓硫的疗效正常', 'AT': '患者使用地尔卓硫的疗效正常', 'AA': '患者使用地尔卓硫的疗效可能较差'}
ID = {} # 药物分类和基因   氨氯地平 --> ACE
locus = {} #药物分类和位点   氨氯地平 --> rs4291
classify = {
            '钙离子拮抗剂':['氨氯地平','维拉帕米','地尔硫䓬','硝苯地平'],
            '血管紧张素Ⅱ受体拮抗（ARB)':['坎地沙坦','厄贝沙坦','替米沙坦','氯沙坦'],
            '利尿剂':['氯噻酮','氢氯噻嗪'],
            '血管紧张素转化酶抑制剂（ACEI)':['赖诺普利','群多普利','贝那普利','卡托普利','咪达普利'],
            'β-受体阻滞剂':['美托洛尔','比索洛尔']
            }

rbook1 = xlrd.open_workbook('Database.xlsx')
rsheet1 = rbook1.sheet_by_index(0)
for row in range(1, rsheet1.nrows):
    if rsheet1.row_values(row)[0]:
        ID[rsheet1.row_values(row)[0]] = rsheet1.row_values(row)[1]
        locus[rsheet1.row_values(row)[2]] = rsheet1.row_values(row)[0]
        name = rsheet1.row_values(row)[0]
        D[name] = {rsheet1.row_values(row)[3]:rsheet1.row_values(row)[4]}
    else:
        D[name][rsheet1.row_values(row)[3]] = rsheet1.row_values(row)[4]

def mkdir(path):
    folder=os.path.exists(path)
    if not folder:
        os.mkdir(path)
    else:
        print('There has been this folder')
current=os.getcwd()+'/Output'
mkdir(current)



rbook2= xlrd.open_workbook('result1.xlsx')
rsheet2 = rbook2.sheet_by_index(0)
with open('gao.tex', 'r') as fr:
   con = fr.read()
   for row in range(1, rsheet2.nrows):
        age=str(2019-int(re.findall(r'^\d+',rsheet2.row_values(row)[4])[0])) #年龄
        gender=rsheet2.row_values(row)[3] #性别
        samplecode=rsheet2.row_values(row)[0] #样本编号
        tel=rsheet2.row_values(row)[5] #联系电话
        dt=datetime.date.today()
        reportdate=str(dt.year)+'年'+str(dt.month)+'月'+str(dt.day)+'日' #报告日期
        submissiondate=rsheet2.row_values(row)[1]
        submissiondatelist=re.findall(r'\d+-\d+-\d+',submissiondate)[0].split('-')
        subdate=submissiondatelist[0]+'年'+submissiondatelist[1]+'月'+submissiondatelist[2]+'日'#送检日期
   
        file_name = rsheet2.row_values(row)[2]
        val = OrderedDict(zip(rsheet2.row_values(0)[15:], rsheet2.row_values(row)[15:]))
        list_name = []
        for k in val.keys():
            list_name.append(k)
        with open(current+'/'+file_name+'.tex', 'w') as fw:
            new_con = re.sub(r'第一', val[list_name[0]], con)
            new_con = re.sub(r'第二', D['氨氯地平'][val[list_name[0]]],new_con)
            new_con = re.sub(r'第三', val[list_name[1]], new_con)
            new_con = re.sub(r'第四', D['维拉帕米'][val[list_name[1]]],new_con)
            new_con = re.sub(r'第五', val[list_name[2]], new_con)
            new_con = re.sub(r'第六', D['地尔硫䓬'][val[list_name[2]]],new_con)
            new_con = re.sub(r'第七', val[list_name[3]], new_con)
            new_con = re.sub(r'第八', D['硝苯地平'][val[list_name[3]]],new_con)
            new_con = re.sub(r'第九', val[list_name[4]], new_con)
            new_con = re.sub(r'第拾', D['坎地沙坦'][val[list_name[4]]],new_con)
            new_con = re.sub(r'十一', val[list_name[5]], new_con)
            new_con = re.sub(r'十二', D['厄贝沙坦'][val[list_name[5]]],new_con)
            new_con = re.sub(r'十三', val[list_name[6]], new_con)
            new_con = re.sub(r'十四', D['替米沙坦'][val[list_name[6]]],new_con)
            new_con = re.sub(r'十五', val[list_name[7]], new_con)
            new_con = re.sub(r'十六', D['氯沙坦'][val[list_name[7]]],new_con)
            new_con = re.sub(r'十七', val[list_name[0]], new_con)
            new_con = re.sub(r'十八', D['氯噻酮'][val[list_name[0]]],new_con)
            new_con = re.sub(r'十九', val[list_name[8]], new_con)
            new_con = re.sub(r'贰十', D['氢氯噻嗪'][val[list_name[8]]],new_con)
            new_con = re.sub(r'二拾一', val[list_name[0]], new_con)
            new_con = re.sub(r'二拾二', D['赖诺普利'][val[list_name[0]]],new_con)
            new_con = re.sub(r'二拾三', val[list_name[1]], new_con)
            new_con = re.sub(r'二拾四', D['群多普利'][val[list_name[1]]],new_con)
            new_con = re.sub(r'二拾五', val[list_name[9]], new_con)
            new_con = re.sub(r'二拾六', D['贝那普利'][val[list_name[9]]],new_con)
            new_con = re.sub(r'二拾七', val[list_name[10]], new_con)
            new_con = re.sub(r'二拾八', D['卡托普利'][val[list_name[10]]],new_con)
            new_con = re.sub(r'二拾九', val[list_name[4]], new_con)
            new_con = re.sub(r'叁十', D['咪达普利'][val[list_name[4]]],new_con)
            new_con = re.sub(r'三拾一', val[list_name[11]], new_con)
            new_con = re.sub(r'三拾二', D['美托洛尔'][val[list_name[11]]],new_con)
            new_con = re.sub(r'三拾三', val[list_name[12]], new_con)
            new_con = re.sub(r'三拾四', D['比索洛尔'][val[list_name[12]]],new_con)
            new_con=re.sub(r'李明祝',file_name,new_con)
            new_con=re.sub(r'name',file_name,new_con)
            new_con=re.sub(r'gender',gender,new_con)
            new_con=re.sub(r'龙',age,new_con)
            new_con=re.sub(r'鼠',tel,new_con)
            new_con=re.sub(r'马',samplecode,new_con)
            new_con=re.sub(r'羊',subdate,new_con)
            new_con=re.sub(r'猴',reportdate,new_con)
            fw.write(new_con)









