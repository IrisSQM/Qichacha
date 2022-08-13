#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据处理可视化-不爬网

Created on Sun Mar  6 11:41:02 2022

@author: shiqimeng
"""
# ------------------------- Data Processing -----------------------------
import pandas as pd
import os
focused = ['红杉','高瓴','蓝驰','洪泰','正心谷','元禾重元','元禾原点','金雨茂物','金茂','毅达',
           '耀途','索道','黑蚁','琢石','风物','薄荷资本','腾讯投资','启明','经纬','鼎晖',
           '景林','淡水泉','重阳','高毅','星石','乐瑞','千合','苏高新创投','国发创投']			

date = '20220221-0227'

# Save path and output file names
path = '<path>' + os.sep + date

fe = pd.read_excel(path + os.sep + '一周融资事件' + date + '.xls')  # 去头企查查导出

index_list = []
focused_in = []

for index, value in fe['投资机构'].items():
    investor = str(value).split(',')
    find = 0
    for i in investor:
        if find == 1:
            break
        for j in focused:
            if j in i:
                index_list.append(index)
                focused_in.append(i.strip())
                find = 1
                break


output = fe.iloc[index_list].reset_index(drop = True)
focused_df = pd.DataFrame(focused_in, columns=['关注机构'])
out = pd.concat([focused_df,output],axis = 1)
out = out.sort_values(by = ['关注机构'])

with pd.ExcelWriter(path + os.sep + '一周融资事件处理后' + date + '.xlsx') as writer:
     out.to_excel(writer)
     
