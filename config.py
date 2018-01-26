#!/usr/bin/python
# -*- coding: utf-8 -*-
# DIZHI = 'F:/2017/老余杭/1/设计结果'
import configparser
import sys
DZ = sys.path[0]
DZ = DZ.strip('base_library.zip')
config = configparser.ConfigParser()
config.read(f'{DZ}\YJKdizhi.ini', encoding='utf-8-sig')
DIZHI = config.get('dizhi', 'yjk')
# DIZHI = ('F:/2017/老余杭/1')
# print(DZ)
WMASSPATTERN = (
    # 基本信息1[0 1 2 3]
    '结构体系:(.*?)\n.*?结构材料信息:(.*?)\n.*?地下室层数:(.*?)\n.*?嵌固端所在层号.*?:(.*?)\n.*?' +
    # 基本信息2[4 5]
    '裙房层数:(.*?)\n.*?P-Delt 效应:(.*?)\n.*?' +
    # 风荷载1[6 7 8]
    '粗糙程度.*?:(.*?)\n.*?基本风压.*?:(.*?)\n.*?X向基本周期.*?:(.*?)\n.*?' +
    # 风荷载2[9 10]
    'Y向基本周期.*?:(.*?)\n.*?风荷载效应放大系数:(.*?)\n.*?' +
    # 地震信息1[11 12 13 14]
    '设计地震分组:(.*?)\n.*?地震烈度:(.*?)\n.*?场地类别:(.*?)\n.*?特征周期:(.*?)\n.*?' +
    # 地震信息2[15 16 17]
    '周期折减系数:(.*?)\n.*?框架的抗震等级:(.*?)\n.*?剪力墙的抗震等级:(.*?)\n.*?' +
    # 地震信息3[18 19 20]
    '是否考虑偶然偏心:(.*?)\n.*?双向地震扭转效应:(.*?)\n.*?最不利地震方向的作用:(.*?)\n.*?' +
    # 活荷载不利布置[21 22]
    '活荷不利布置的最高层号:(.*?)\n.*?梁活荷载内力放大系数:(.*?)\n.*?' +
    # 保护层[23 24]
    '梁保护层厚度.*?:(.*?)\n.*?柱保护层厚度.*?:(.*?)\n.*?' +
    # 包络[25 26 27]
    '分别计算，并取大:(.*?)\n.*?抗震墙模型计算大值:(.*?)\n.*?模型进行包络取大:(.*?)\n.*?')
# 最不利地震方向
WMASSPATTERN2 = ('自动计算最不利地震方向的作用:\s*(.*?)\n')

# 周期信息[0 1 2 3]
WZQPATTERN = ('(第1扭转周期.*?)=(.*?)\n.*?(地震作用最大的方向).*?=(.*?)°')
# 地震剪力信息[0 1]
WZQPATTERN2 = ('(\d*?.\d*?)\s*?(\d*?.\d*?)\s*?:本层地震剪力不满足')

# 最大层间位移角
WDISPPATTERN = ('X 方向地震作用下的楼层最大位移.*?X向最大层间位移角：(.*?)\(.*?' +
                'Y 方向地震作用下的楼层最大位移.*?Y向最大层间位移角：(.*?)\(.*?' +
                'X 方向风荷载作用下的楼层最大位移.*?X向最大层间位移角：(.*?)\(.*?' +
                'Y 方向风荷载作用下的楼层最大位移.*?Y向最大层间位移角：(.*?)\(.*?')
# 最大位移比
WDISPBI1PATTERN = ('X方向最大位移与层平均位移的比值：(.*?)\(.*?')
WDISPBI2PATTERN = ('X方向最大层间位移与平均层间位移的比值：(.*?)\(.*?')
WDISPBI3PATTERN = ('Y方向最大位移与层平均位移的比值：(.*?)\(.*?')
WDISPBI4PATTERN = ('Y方向最大层间位移与平均层间位移的比值：(.*?)\(.*?')
# 薄弱层
BORUOCENG = ('薄弱层地震剪力放大系数=(.*?)\n')
BORUOCENGHAO = ('本工程如下楼层为薄弱层：(.*?)\n{2}')
BORUOCENGHAO2 = ('\n\s*(\d+)\s*(\d+)\s')
# 嵌固端
QIANGUDUAN = (
    'Ratx.*?Floor No.(.*?)Tower No.(.*?)\n.*?Ratx =(.*?)Raty =(.*?)\n')
QIANGUDUAN2 = ('底部嵌固楼层刚度比执行《高规》3.5.2-2:\s*(.*?)\n')
# 质量比
ZHILIANGBI = ('\n\s*(\d+)\s*(\d+)\s.*?质量比>1.5 不满足《高规》3.5.6')