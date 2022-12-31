#! /usr/bin/env python
# -*- coding: utf-8 -*-
#
# =============================================================================
# Copyright (c) 2019, Fujian BFDX Tech Co., Ltd.
# All rights reserved.
#
# Change Logs:
# Date			Author		Action
# 2022/12/29	cx
# =============================================================================
#

import xlrd2
import xml.dom.minidom as minidom

def writexml(xlsx_path, xml_path):
    # 打开表格
    workbook = xlrd2.open_workbook(xlsx_path)
    # 定位工作表
    worksheet = workbook.sheet_by_name("Output")

    # 创建DOM树对象
    dom = minidom.Document()
    # 创建根节点。每次都要用DOM对象来创建任何节点。
    root_node = dom.createElement('strings')
    # 用DOM对象添加根节点
    dom.appendChild(root_node)

    # 表内词条起始行数和列数
    startnrowidx = 1
    startncolsidx = 3
    print("词条总数：", worksheet.nrows - startnrowidx, "语言种数：", worksheet.ncols - startncolsidx - 1)
    # 获取第一行标题
    tabletitlerow_list = worksheet.row_values(0, startncolsidx, None)

    # tablecol_list = worksheet.col_values(startncolsidx, startnrowidx, None)
    # print(tablecol_list)

    for rownum in range(startnrowidx, worksheet.nrows):
        tablerow_list = worksheet.row_values(rownum, startncolsidx, None)
        print(tablerow_list)

        string = dom.createElement('string')
        # 设置该节点的属性
        string.setAttribute('name', tablerow_list[0])
        root_node.appendChild(string)

        for colnum in range(1, worksheet.ncols - startncolsidx):
            # 这里判断对应ID语言有没有词条 如果是数值类型会报错
            if len(tablerow_list[colnum]):
                # 用DOM对象创建language元素
                language_node = dom.createElement('language')
                # 添加在根元素下
                string.appendChild(language_node)
                # 设置language元素的属性
                language_node.setAttribute('name', tabletitlerow_list[colnum])
                text = dom.createTextNode(tablerow_list[colnum])
                language_node.appendChild(text)

    with open(xml_path, 'w', encoding='UTF-8') as fh:
         dom.writexml(fh, indent='', addindent=' ', newl='\n', encoding='UTF-8')

def deletelines(filename, headnum, tailnum):
    readfile = open(filename, 'r', encoding='utf-8')
    lines = readfile.readlines()
    writefile = open(filename, 'w', encoding='utf-8')
    lines = ''.join(lines[headnum:-tailnum])
    writefile.write(lines)

if __name__ == "__main__":
    src_path = "./language.xlsx"
    dst_path = "./strings.xml"
    writexml(src_path, dst_path)
    # 去除XML标签和根节点，直接删除行
    deletelines(dst_path, 2, 1)
