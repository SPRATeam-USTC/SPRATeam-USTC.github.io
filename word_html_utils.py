#!/usr/bin/env python
# -*- coding:utf-8 -*-
###
# File: /Users/majiefeng/Desktop/其他项目/doc2hyperlink.py
# Project: /Users/majiefeng/Desktop/其他项目
# Created Date: 2023-09-13 15:49:30
# Author: JeffreyMa
# -----
# Last Modified: 2023-09-13 17:48:40
# Modified By: JeffreyMa
# -----
# Copyright (c) 2023 USTC
# -----
# HISTORY:
# Date      	By	Comments
# ----------	---	----------------------------------------------------------
###
import docx
import os

def extract_hyperlinks(docx_file):
    # 打开Word文档
    doc = docx.Document(docx_file)
    
    # 创建一个空列表来存储超链接内容，由于可能会有重复的text，所以不能用dict
    hyperlinks = []
    
    # 遍历文档中的所有超链接
    for paragraph in doc.paragraphs:
        for link in paragraph._element.xpath(".//w:hyperlink"):
            inner_run = link.xpath("w:r", namespaces=link.nsmap)[0]
            text = inner_run.text
            rId = link.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            link = doc._part.rels[rId]._target
            print(text, link)
            hyperlinks.append((text, link))
    
    return hyperlinks

def judge_contain(text, substrs):
    contain_ = [str_ in text for str_ in substrs]
    return True in contain_

def assign_plain_texts(plain_texts, text2links):
    cur_ptr = 0
    for line in plain_texts:
        if not judge_contain(line, [x[0] for x in text2links[cur_ptr:]]):
            print(line.strip())
            continue
        while cur_ptr < len(text2links):
            text, link = text2links[cur_ptr]
            if text in line:
                line = line.replace(text, "<a href=\"{}\">{}</a>".format(link, text))
                cur_ptr += 1
                break
            cur_ptr += 1
        print(line.strip())

if __name__ == "__main__":
    # 指定要提取超链接的Word文档文件名
    docx_file = "/Users/majiefeng/Desktop/CV-2023-editable.docx"
    plain_texts = open("temp_text.txt").readlines()
    
    # 提取超链接内容
    text2links = extract_hyperlinks(docx_file)
    # assign_plain_texts(plain_texts, text2links)