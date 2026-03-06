#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
目录标题提取模块 - 负责从文档中提取目录页的标题信息
"""

import re
from typing import List, Dict
from config import TITLE_PREFIX_PATTERN, PAGE_NUMBER_PATTERN

def getOutlineLevel(inputXml):
    """
    功能 从xml字段中提取出<w:outlineLvl w:val="number"/>中的数字number
    参数 inputXml
    返回 number
    """
    start_index = inputXml.find('<w:outlineLvl')
    end_index = inputXml.find('>', start_index)
    number = inputXml[start_index:end_index+1]
    number = re.search(r"\d+", number).group()  # Fixed: Added raw string prefix
    return number

def isTitle(paragraph):
    """
    功能 判断该段落是否设置了大纲等级
    参数 paragraph:段落
    返回 None:普通正文，没有大纲级别 0:一级标题 1:二级标题 2:三级标题
    """
    # 如果是空行，直接返回None
    if paragraph.text.strip() == '':
        return None
        
    # 如果该段落是直接在段落里设置大纲级别的，根据xml判断大纲级别
    paragraphXml = paragraph._p.xml
    if paragraphXml.find('<w:outlineLvl') >= 0:
        #print(paragraph.text)
        return getOutlineLevel(paragraphXml)
    # 如果该段落是通过样式设置大纲级别的，逐级检索样式及其父样式，判断大纲级别
    targetStyle = paragraph.style
    while targetStyle is not None:
        # 如果在该级style中找到了大纲级别，返回
        if targetStyle.element.xml.find('<w:outlineLvl') >= 0:
            #print(paragraph.text)
            return getOutlineLevel(targetStyle.element.xml)
        else:
            targetStyle = targetStyle.base_style
    # 如果在段落、样式里都没有找到大纲级别，返回None
    return None

def extract_toc_titles(paragraphs: List) -> List[Dict[str, str]]:
    """从段落列表中提取目录页的有效标题（跨多页），返回原始与清洗后标题"""
    toc_entries = []
    for i, para in enumerate(paragraphs):
    #for para in paragraphs:
        raw_text = para.text  # 原始文本（含制表符、换行符等）
        # 跳过空行
        if not raw_text.strip():
            continue
        # 检查大纲级别
        outline_level = isTitle(para)
        if outline_level is None:
            continue  # 不是标题，跳过
        # 检查是否符合标题前缀模式
        if TITLE_PREFIX_PATTERN.match(raw_text):
            # 清洗：1) 移除页码；2) 替换 \t\r\n 为单空格；3) strip 首尾空格
            cleaned = PAGE_NUMBER_PATTERN.sub('', raw_text)
            cleaned = re.sub(r'[\t\r\n]+', ' ', cleaned).strip()
            toc_entries.append({
                "raw": raw_text,        # 用于严格匹配
                "cleaned": cleaned,      # 用于输出到"目录标题"列
                "idx_num": i              # 记录段落索引，后续关联正文标题和表格
            })
    
    return toc_entries

def extract_full_toc(paragraphs: List) -> List[Dict[str, str]]:
    """
    从段落列表中提取完整的目录信息（包含所有类型的标题）
    返回原始与清洗后的完整标题列表，用于章节名匹配
    """
    full_toc_entries = []
    
    for para in paragraphs:
        raw_text = para.text  # 原始文本（含制表符、换行符等）
        if not raw_text.strip():
            continue
        # 检查大纲级别
        outline_level = isTitle(para)
        if outline_level is None:
            continue  # 不是标题，跳过
             
        # 清洗标题：移除页码，替换特殊字符为空格，strip首尾空格
        cleaned = PAGE_NUMBER_PATTERN.sub('', raw_text)
        cleaned = re.sub(r'[\t\r\n]+', ' ', cleaned).strip()
        
        full_toc_entries.append({
            "raw": raw_text,        # 原始文本
            "cleaned": cleaned      # 清洗后文本（用于匹配）
        })
    
    return full_toc_entries