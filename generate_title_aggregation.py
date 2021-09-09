#!/usr/bin/python
# -*- coding: utf-8 -*-
# @Time    : 2020.01.29
# @Author  : lishun02@baidu.com
# @FileName: generate_title_aggregation.py

import json

import xlrd
import sys
import argparse
import os

G_VERSION = "1.0"
G_TYPE_NAME_INDEX = 0
g_aggregation_id_index = 0
g_resource_name_index = 1
g_resource_id_index = 2
g_resource_is_vip_index = 4
g_resource_image_index = -1
g_current_aggregation_id = -1
g_output_path_dir = "./output/aggregation/"
g_output_file_name_prefix = "aggregation-"
g_output_file_name_suffix = ".json"
g_koudai_aggregation_base_id = 20000  # 口袋故事
g_qingting_aggregation_base_id = 30000  # 蜻蜓
g_dushulang_aggregation_base_id = 40000  # 读书郎
g_sheet_name = ''

g_current_aggregation_base_id = g_qingting_aggregation_base_id

parser = argparse.ArgumentParser(description="传入需要转成json的运营文件")
parser.add_argument("-f", "--file", help="运营文件的路径")
parser.add_argument("--version", action="store_true", help="查看版本号")
parser.add_argument("-s", "--sheet", help="excel的表单名")
parser.add_argument("-r", "--resource", help="资源方")
args = parser.parse_args()
filePath = ''

if args.version:
    print "version：" + G_VERSION
    exit(1)

if args.resource:
    g_current_aggregation_base_id = float(args.resource)
else:
    print "未选择资源放，-r 20000(koudai),30000(qingting),40000(dushulang)"
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

if args.sheet:
    g_sheet_name = args.sheet

reload(sys)
sys.setdefaultencoding('utf8')


class RosourceListItemData:
    def __init__(self):
        self.name = ''
        self.resourceUrl = ''
        self.imageUrl = ''
        self.isVip = False
        self.isAggregation = False
        self.aggregationId = 0


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
if g_sheet_name != '':
    my_sheet = my_excel.sheet_by_name(g_sheet_name)
else:
    my_sheet = my_excel.sheets()[0]

max_row = my_sheet.nrows
max_clos = my_sheet.ncols
data = []
rootData = RootData()

line_one_data = my_sheet.col_values(0)

mkdir(g_output_path_dir)

for c in range(max_clos):
    if str(my_sheet.cell_value(0, c)) == "专辑ID":
        g_resource_id_index = c
    elif str(my_sheet.cell_value(0, c)) == "是否为会员":
        g_resource_is_vip_index = c
    elif str(my_sheet.cell_value(0, c)) == "专辑名":
        g_resource_name_index = c
    elif str(my_sheet.cell_value(0, c)) == "聚合ID":
        g_aggregation_id_index = c
    elif str(my_sheet.cell_value(0, c)) == "节目封面":
        g_resource_image_index = c

for r in range(max_row):
    if r == 0:
        continue
    if len(str(my_sheet.cell_value(r, 0))) != 0:
        if g_current_aggregation_id < 0:
            g_current_aggregation_id = my_sheet.cell_value(r, g_aggregation_id_index)
        elif my_sheet.cell_value(r, g_aggregation_id_index) != g_current_aggregation_id:
            print "新的聚合文件"
            json_str = json.dumps(rootData, default=lambda o: o.__dict__, sort_keys=False, indent=4).decode(
                "unicode-escape")
            print r
            print g_current_aggregation_base_id
            print str(long(g_current_aggregation_base_id + g_current_aggregation_id))
            with open(g_output_path_dir + g_output_file_name_prefix
                      + str(long(g_current_aggregation_base_id + g_current_aggregation_id))
                      + g_output_file_name_suffix, "w") as fp:
                fp.write(json_str)
            g_current_aggregation_id = my_sheet.cell_value(r, g_aggregation_id_index)
            rootData = RootData()
        rosourceListItemData = RosourceListItemData()
        rosourceListItemData.name = my_sheet.cell_value(r, g_resource_name_index)
        rosourceListItemData.resourceUrl = 'dueros://audio_unicast_story/albumplay?album_id=' + str(
            long(my_sheet.cell_value(r, g_resource_id_index)))
        if g_resource_image_index > 0:
            rosourceListItemData.imageUrl = my_sheet.cell_value(r, g_resource_image_index)
        else:
            rosourceListItemData.imageUrl = 'https://iot-paas-static.cdn.bcebos.com/XTC/imgs/index/' + rosourceListItemData.name + ".png"
        if my_sheet.cell_value(r, g_resource_is_vip_index) == '是':
            rosourceListItemData.isVip = True
        rootData.data.append(rosourceListItemData)
        if r == max_row - 1:
            json_str = json.dumps(rootData, default=lambda o: o.__dict__, sort_keys=False, indent=4).decode(
                "unicode-escape")
            with open(g_output_path_dir + g_output_file_name_prefix
                      + str(long(g_current_aggregation_base_id + g_current_aggregation_id))
                      + g_output_file_name_suffix, "w") as fp:
                fp.write(json_str)
    elif len(my_sheet.cell_value(r, 1)) == 0:
        print "空行"
