#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Output Name生成模块 - 负责根据目录标题生成规范化的输出名称
"""

import re
from config import PREFIX_MAP

def generate_output_name(directory_title: str) -> str:
    """
    根据目录标题生成Output Name
    规则：
    - 表→t，图→f，列表→l
    - 按.分割编号部分
    - 纯数字部分补零到2位（1→01，14→14）
    - 单独字母前补0（a→0a）
    - 数字+字母组合保持原样（1a→1a，14a→14a）
    """
    # 提取前缀和编号部分
    # 修改正则以正确提取编号部分（不包含后续文本）
    prefix_pattern = re.compile(r'^(表|图|列表)\s*([\d\.\-\_a-zA-Z]+)')
    match = prefix_pattern.match(directory_title)
    
    if not match:
        return ""  # 如果不匹配格式，返回空字符串
    
    prefix = match.group(1)
    number_part = match.group(2)  # 只获取编号部分
    
    mapped_prefix = PREFIX_MAP.get(prefix, "")
    
    if not number_part:
        return mapped_prefix
    
    # 先将横线和下划线统一替换为点号，然后再按点号分割
    normalized_number = number_part.replace('-', '.').replace('_', '.')
    
    # 按点号分割编号部分
    parts = normalized_number.split('.')
    
    processed_parts = []
    for i, part in enumerate(parts):
        part = part.strip()
        if not part:
            continue
            
        # 处理纯数字部分
        if part.isdigit():
            processed_parts.append(part.zfill(2))
        # 处理字母部分
        elif part.isalpha() and len(part) == 1:
            processed_parts.append("0" + part)
        # 处理数字+字母组合 - 保持原样
        elif len(part) > 1 and part[:-1].isdigit() and part[-1].isalpha():
            # 直接保持原样，不进行任何处理
            processed_parts.append(part)
        else:
            # 其他情况保持原样
            processed_parts.append(part)
    
    # 拼接结果
    return mapped_prefix + "".join(processed_parts)