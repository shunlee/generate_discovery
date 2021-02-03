#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Time    : 2020.01.13
# @Author  : lishun02@baidu.com
# @FileName: generate_special_aggregation.py

import json

import xlrd
import sys
import argparse
import os

G_VERSION = "2.0"
G_TYPE_NAME_INDEX = 0
G_RESOURCE_NAME_INDEX = 1
G_OUTPUT_PATH_DIR = "./output/special/"
G_RESOURCE_TYPE_EXCEL_SQUARE = '正方形'
G_RESOURCE_TYPE_EXCEL_CAROUSEL = '大卡轮播'
G_RESOURCE_TYPE_SQUARE = 'resource_1_2'
G_RESOURCE_TYPE_BANNER = 'resource_banner'

g_aggregation_id_index = 2
g_resource_id_index = 3
g_resource_is_vip_index = 5
g_current_type = ''
g_koudai_aggregation_id = 20000

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
        self.name = ''
        self.resourceName = ''
        self.resourceUrl = ''
        self.imageUrl = ''
        self.isVip = False
        self.isAggregation = False
        self.aggregationId = -1


class resourceData:
    def __init__(self):
        self.name = ''
        self.subTitle = ''
        self.displayItemCount = -1
        self.resourceType = ''
        self.resourceList = []


class RootData:
    def __init__(self):
        self.status = 0
        self.msg = "ok"
        self.data = []


def mkdir(path):
    path = path.strip()
    path = path.rstrip("\\")
    if not os.path.exists(path):
        os.makedirs(path)
        print path + ' 创建成功'
        return True
    else:
        print path + ' 目录已存在'
        return False


my_excel = xlrd.open_workbook(filePath)
# my_sheet = my_excel.sheets()[0]
my_sheet = my_excel.sheet_by_name("专辑页")
max_row = my_sheet.nrows
# max_clos = my_sheet.ncols
max_clos = 7
data = []
rootData = RootData()

line_one_data = my_sheet.col_values(0)

mkdir(G_OUTPUT_PATH_DIR)

for c in range(max_clos):
    if str(my_sheet.cell_value(0, c)) == "专辑ID":
        g_resource_id_index = c
    elif str(my_sheet.cell_value(0, c)) == "是否会员":
        g_resource_is_vip_index = c
    elif str(my_sheet.cell_value(0, c)) == "聚合专辑":
        g_aggregation_id_index = c
for r in range(max_row):
    if len(my_sheet.cell_value(r, 0)) != 0:
        print "str(my_sheet.cell_value(r, G_TYPE_NAME_INDEX)) = " + str(my_sheet.cell_value(r, G_TYPE_NAME_INDEX))
        if str(my_sheet.cell_value(r, G_TYPE_NAME_INDEX)) == G_RESOURCE_TYPE_EXCEL_SQUARE:
            rootData.data.append(resourceData())
            g_current_type = G_RESOURCE_TYPE_EXCEL_SQUARE
            rootData.data[len(rootData.data) - 1].name = my_sheet.cell_value(r, G_RESOURCE_NAME_INDEX)
            rootData.data[len(rootData.data) - 1].resourceType = G_RESOURCE_TYPE_SQUARE
            try:
                print '换一换 ---> ' + 'r = ' + str(r) + 'data = ' + str(my_sheet.cell_value(r, max_clos - 1))
                if my_sheet.cell_value(r, max_clos - 1) > 0:
                    rootData.data[len(rootData.data) - 1].displayItemCount = long(my_sheet.cell_value(r, max_clos - 1))
                    rootData.data[len(rootData.data) - 1].subTitle = '换一换'
            except Exception:
                rootData.data[len(rootData.data) - 1].displayItemCount = -1
                rootData.data[len(rootData.data) - 1].subTitle = ''
        elif str(my_sheet.cell_value(r, G_TYPE_NAME_INDEX)) == G_RESOURCE_TYPE_EXCEL_CAROUSEL:
            rootData.data.append(resourceData())
            rootData.data[len(rootData.data) - 1].displayItemCount = 0
            g_current_type = G_RESOURCE_TYPE_EXCEL_CAROUSEL
            rootData.data[len(rootData.data) - 1].name = my_sheet.cell_value(r, G_RESOURCE_NAME_INDEX)
            rootData.data[len(rootData.data) - 1].resourceType = G_RESOURCE_TYPE_BANNER
            rootData.data[len(rootData.data) - 1].name = G_RESOURCE_TYPE_EXCEL_CAROUSEL
    elif len(my_sheet.cell_value(r, 1)) == 0:
        print "空行"
    else:
        rosourceListItemData = RosourceListItemData()
        rosourceListItemData.resourceName = my_sheet.cell_value(r, G_RESOURCE_NAME_INDEX)
        rosourceListItemData.name = my_sheet.cell_value(r, G_RESOURCE_NAME_INDEX)
        if my_sheet.cell_value(r, g_resource_is_vip_index) == '是':
            rosourceListItemData.isVip = True
        # 如果聚合专辑id是数字，认为是聚合，不然都是非集合
        try:
            rosourceListItemData.aggregationId = g_koudai_aggregation_id + long(
                my_sheet.cell_value(r, g_aggregation_id_index))
            rosourceListItemData.isAggregation = True
        except Exception:
            rosourceListItemData.resourceUrl = 'dueros://audio_unicast_story/albumplay?album_id=' + str(
                long(my_sheet.cell_value(r, g_resource_id_index)))
        rosourceListItemData.imageUrl = 'https://iot-paas-static.cdn.bcebos.com/XTC/imgs/index/' + rosourceListItemData.resourceName + ".png"
        rootData.data[len(rootData.data) - 1].resourceList.append(rosourceListItemData)
        if g_current_type == G_RESOURCE_TYPE_EXCEL_CAROUSEL:
            rosourceListItemData.imageUrl = 'https://iot-paas-static.cdn.bcebos.com/XTC/imgs/index/' + rosourceListItemData.resourceName + "_banner.png"

json_str = json.dumps(rootData, default=lambda o: o.__dict__, sort_keys=False, indent=4).decode("unicode-escape")
with open(G_OUTPUT_PATH_DIR + "special.json", "w") as fp:
    fp.write(json_str)
