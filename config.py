#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置文件 - 包含全局常量和正则表达式定义
"""

import re

# 全局正则表达式定义
# 匹配目录中的有效标题行（表X. / 图X. / 列表X. 开头）
TOC_LINE_PATTERN = re.compile(r'^([表图列表]\d+\.\s*.+)(?:\s+[\.\u3000]*\s+\d+)?$')

# 标题格式检测（用于判断是否为下一个标题）
TITLE_PATTERN = re.compile(r'^[表图列表]\d+\.')

# 表格编号模式（用于脚注终止条件）- 匹配"表 XXX"、"图 XXX"、"列表 XXX"格式
TABLE_NUMBER_PATTERN = re.compile(r'^(表|图|列表)\s+\d+')

# 标题前缀模式
TITLE_PREFIX_PATTERN = re.compile(r'^(表|图|列表)\s*[\d\.\-\_]+\.')

# 页码模式：行尾数字（可能带点线或空格）
PAGE_NUMBER_PATTERN = re.compile(r'\s*[\.\u3000\u2026\-]*\s*\d+$')

# Output Name前缀映射
PREFIX_MAP = {"表": "t", "图": "f", "列表": "l"}

# 默认输入文件路径
DEFAULT_INPUT_FILE = r"d:\OneDrive\Python\Extract Title Footnote\Dummy Shell.docx"

# 自定义footnote关键词配置
# 默认包含programming和programmer，用户可添加更多
DEFAULT_CUSTOM_KEYWORDS = ["programming", "programmer"]