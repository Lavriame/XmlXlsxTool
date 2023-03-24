#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import codecs
import xlsxwriter
from xml.dom.minidom import parse

def get_files(root_path): 
	file_list = []
	for i in os.listdir(root_path):
		path = root_path + r'/' + i
		if os.path.isfile(path):
			file_list.append(path)
		elif os.path.isdir(path):
			files = get_files(path)
			for f in files:
				file_list.append(f)
	return file_list


def word_in_files(root_path, word):
	file_list = get_files(root_path)
	result = []
	for path in file_list:
		if word in codecs.open(path, 'r', encoding='utf-8').read():  
			result.append(path)
	return result

def XML2XLSX(xml_path, xlsx_path):
	# domtree = parse(r"C:\Users\Administrator\PycharmProjects\pythonProject2\strings.xml")
	domtree = parse(xml_path)
	# 文档根元素
	rootNode = domtree.documentElement
	print("根元素:", rootNode.nodeName)

	# 所有词条
	strings = rootNode.getElementsByTagName("string")

	# 创建excel
	workbook = xlsxwriter.Workbook(xlsx_path)
	worksheet0 = workbook.add_worksheet('Output')

	# XML内相关节点属性
	language_list = ['ID', 'en_US', 'zh_CN']
	xlsxstartidx = 3

	# 从D列开始写
	worksheet0.write_row('D1', language_list)
	num = 1
	used = 0
	useless = 0

	print("********所有词条********")
	for string in strings:
		if string.hasAttribute("name"):
			ID = string.getAttribute("name")
			print("name:", ID)
			worksheet0.write(num, xlsxstartidx, ID)

			# 检索是否是有用ID
			if word_in_files(r".\src", '"' + ID + '"'):
				used = used + 1
				# 获取是否有language标签
				if string.getElementsByTagName("language"):
					languages = string.getElementsByTagName("language")
					# print("languages:", string.getElementsByTagName("language"))

					worksheet0.write(num, xlsxstartidx + len(language_list) + 1, '√')

					# 获取是language标签下的语言标签
					for language in languages:
						name = language.getAttribute("name")
						print("language name:", name)
						print(language.childNodes[0].data)

						for languageidx in range(1, len(language_list)):
							if name == language_list[languageidx]:
								worksheet0.write(num, xlsxstartidx + languageidx, language.childNodes[0].data)

						# if name == language_list[1]:
						# 	worksheet0.write(num, 4, language.childNodes[0].data)
						# elif name == language_list[2]:
						# 	worksheet0.write(num, 5, language.childNodes[0].data)
						# else:
						# 	worksheet0.write(num, 6, language.childNodes[0].data)

			else:
				useless = useless + 1
				worksheet0.write(num, xlsxstartidx + len(language_list) + 1, '×')

			num = num + 1

	print("********共计", strings.length, "个词条********")
	print("********有效", used, "无效", useless, "********")
	workbook.close()

def addlines(filename):
	readfile = codecs.open(filename, 'r', encoding='utf-8')
	lines = readfile.readlines()
	if lines[0] != '<strings>\n':
		writefile = codecs.open(filename, 'w', encoding='utf-8')
		print(lines[0])
		lines[0] = '<strings>\n' + lines[0]
		lines[-1] = lines[-1] + '\n</strings>'
		lines = ''.join(lines[0:])
		writefile.write(lines)

if __name__ == '__main__':
	dst_path = "./language.xlsx"
	src_path = "./strings.xml"
	addlines(src_path)
	XML2XLSX(src_path, dst_path)
