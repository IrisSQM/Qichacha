#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
投资项目词云

Created on Wed Apr  6 11:58:28 2022

@author: shiqimeng

参考博客：
jieba: https://www.cnblogs.com/jackchen-net/p/8207009.html
词云: https://cloud.tencent.com/developer/article/1811219
"""
from PIL import Image
import jieba
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os

path = '<path>' 

# 读取停用词列表
with open(path + os.sep + '词云/cn_stopwords.txt','r',encoding='utf-8') as f:
    stwords = [word.strip('\n') for word in f.readlines()]

# 读取月度主营业务
txt = pd.read_excel(path + os.sep + '月度主营业务汇总/月度主营业务汇总_202205.xlsx')
txt = txt['主营业务'].to_list()

# 去除停用词的函数
def remove_stop(seg):
    '''
    Parameters
    ----------
    seg : 分好词的list

    Returns
    -------
    no_stop 去除停用词的list
    '''
    no_stop = []
    for wd in seg:
        if wd not in stwords:
            no_stop.append(wd)
    return no_stop

text = []
for para in txt:
    seg_ls = jieba.lcut(para, cut_all = False)  # 精确模式，返回list
    text.extend(remove_stop(seg_ls))
    
wc = WordCloud(background_color="white",# 设置背景颜色
           max_words = 100, # 词云显示的最大词数
           max_font_size = 90, #最大字体     
           stopwords = stwords, # 设置停用词
           mask=np.array(Image.open(path + os.sep + '词云/unicorn.jpg')), # 选择背景图片
           font_path = path + os.sep + '词云/SourceHanSerifSC-Regular.otf', # 兼容中文字体，不然中文会显示乱码
           )
# 生成词云 
wc.generate(" ".join(text)) 

# 生成的词云图像保存到本地
wc.to_file(path + os.sep + '词云.png')

# 显示图像
# plt.imshow(wc, interpolation='bilinear')
# interpolation='bilinear' 表示插值方法为双线性插值
# plt.axis("off")# 关掉图像的坐标
# plt.show()
