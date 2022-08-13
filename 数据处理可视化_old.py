#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Qichacha data handling & visualization

Created on Mon Dec  6 10:52:04 2021

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

fe = pd.read_csv(path + os.sep + '一周融资事件' + date + '.csv')

index_list = []
focused_in = []

for index, value in fe['投资机构'].items():
    investor = value.split(',')
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
     
# with pd.ExcelWriter(path + os.sep + comp_name + '.xlsx', mode = 'a', engine='openpyxl') as writer:
#      df_kp.to_excel(writer, sheet_name = '主要人员')
# ------------------------- Visualization - Heatmap [pyecharts]-----------------------------
# Install packages
# pip install echarts-china-cities-pypkg
# pip install pyecharts

import pyecharts
from pyecharts.charts import Geo
from pyecharts import options as opts
# from pyecharts.faker import Faker
from pyecharts.globals import ChartType
import os
import pandas as pd

date = '20220221-0227'

# Save path and output file names
save_path = '<path>' + os.sep + date

fe = pd.read_csv(save_path + os.sep + '一周融资事件' + date + '.csv')

# Get location
loc = fe[fe['所属地']!='-']['所属地']
city_dist = loc.value_counts(sort = False).to_dict()
city = list(city_dist.keys())
values = list(city_dist.values())

ch = (Geo()
      .add_schema(maptype = 'china',
                  itemstyle_opts=opts.ItemStyleOpts(color="#ddeded", border_color="#111"))
      .add('融资事件数量',[list(z) for z in zip(city,values)],
           type_=ChartType.HEATMAP)
      .add('',[list(z) for z in zip(city,values)],
           type_=ChartType.EFFECT_SCATTER)
      .set_series_opts(label_opts = opts.LabelOpts(is_show = False))
      .set_global_opts(
          visualmap_opts = opts.VisualMapOpts(max_ = max(values),min_ = min(values)),
          title_opts = opts.TitleOpts(title = '一周融资事件地区分布图')
          )
      .render(save_path + os.sep + '一周融资事件地区分布图' + date +'.html')
      )
