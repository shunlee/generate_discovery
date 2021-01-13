#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Time    : 2020.01.13
# @Author  : lishun02@baidu.com
# @FileName: generate_discovery.py

import json

import xlrd
import sys
import argparse
import os

G_VERSION = "1.0"
G_TYPE_NAME_INDEX = 0
G_RESOURCE_NAME_INDEX = 1
g_resource_id_index = 2
g_resource_is_vip_index = 3

parser = argparse.ArgumentParser(description="传入需要转成json的运营文件")
parser.add_argument("-f", "--file", help="运营文件的路径")
parser.add_argument("--version", action="store_true", help="查看版本号")

args = parser.parse_args()
filePath = ''

if args.version:
    print "version：" + G_VERSION
    exit(1)

if args.file:
    if not (os.path.exists(args.file)):
        print "文件不存在，patch -->" + args.file
        exit(1)
    else:
        filePath = args.file
else:
    print "传入的文件路径为空！请使用-f或--file传入文件路径，例如：-f ./x.xlsx 或者--file ./x.xlsx"
    exit(1)

reload(sys)
sys.setdefaultencoding('utf8')


class RosourceListItemData:
    def __init__(self):
        self.resourceName = ''
        self.resourceUrl = ''
        self.imageUrl = ''
        self.isVip = False
        self.isAggregation = False


class resourceData:
    def __init__(self):
        self.name = ''
        self.subtitle = ''
        self.displayItemCount = -1
        self.resourceType = ''
        self.resourceList = []


class RootData:
    def __init__(self):
        self.status = 0
        self.msg = "ok"
        self.data = []


my_excel = xlrd.open_workbook(filePath)
my_sheet = my_excel.sheets()[0]
max_row = my_sheet.nrows
max_clos = my_sheet.ncols
data = []
rootData = RootData()

line_one_data = my_sheet.col_values(0)

print max_clos

for c in range(max_clos):
    if str(my_sheet.cell_value(0, c)) == "专辑ID":
        g_resource_id_index = c
    elif str(my_sheet.cell_value(0, c)) == "是否会员":
        g_resource_is_vip_index = c
for r in range(max_row):
    if len(my_sheet.cell_value(r, 0)) != 0:
        print "标题"
        rootData.data.append(resourceData())
        rootData.data[len(rootData.data) - 1].name = my_sheet.cell_value(r, G_RESOURCE_NAME_INDEX)
        if str(my_sheet.cell_value(r, G_TYPE_NAME_INDEX)) == '正方形':
            rootData.data[len(rootData.data) - 1].resourceType = 'resource_1_2'
        try:
            if len(my_sheet.cell_value(r, 4) != 0):
                rootData.data[len(rootData.data) - 1].displayItemCount = my_sheet.cell_value(r, 4)
                rootData.data[len(rootData.data) - 1].subtitle = '换一换'
        except Exception:
            rootData.data[len(rootData.data) - 1].displayItemCount = -1
            rootData.data[len(rootData.data) - 1].subtitle = ''
    elif len(my_sheet.cell_value(r, 1)) == 0:
        print "空行"
    else:
        print "真实数据"
        rosourceListItemData = RosourceListItemData()
        rosourceListItemData.resourceName = my_sheet.cell_value(r, G_RESOURCE_NAME_INDEX)
        if my_sheet.cell_value(r, g_resource_is_vip_index) == '是':
            rosourceListItemData.isVip = True
        if my_sheet.cell_value(r, g_resource_id_index) > 0:
            rosourceListItemData.resourceUrl = 'dueros://audio_unicast_story/albumplay?album_id=' + str(
                long(my_sheet.cell_value(r, g_resource_id_index)))
        rosourceListItemData.imageUrl = 'https://iot-paas-static.cdn.bcebos.com/XTC/imgs/index/' + rosourceListItemData.resourceName + ".png"
        rootData.data[len(rootData.data) - 1].resourceList.append(rosourceListItemData)

json_str = json.dumps(rootData, default=lambda o: o.__dict__, sort_keys=False, indent=4).decode("unicode-escape")
print json_str
