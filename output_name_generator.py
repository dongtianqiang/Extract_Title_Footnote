#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Output Name生成模块 - 负责根据目录标题生成规范化的输出名称
"""

import re
from config import PREFIX_MAP

def generate_output_name(directory_title: str) -> str:
    """
    根据目录标题生成 Output Name
    规则：
    - 表→t，图→f，列表→l
    - 按。分割编号部分
    - 纯数字部分补零到 2 位（1→01，14→14）
    - 包含数字和字母的组合：数字部分补零，字母部分保持原样（1a→01a，14a→14a）
    - 纯字母部分保持原样（a→a，ab→ab）
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
                
        # Check if part contains any digits
        has_digit = any(char.isdigit() for char in part)
        # Check if part contains any alphabetic character
        has_alpha = any(char.isalpha() for char in part)
            
        if has_digit:
            # Extract leading digits and trailing letters
            digit_end_index = 0
            for j, char in enumerate(part):
                if char.isdigit():
                    digit_end_index = j + 1
                else:
                    break
                
            # Pad the digit portion to 2 digits
            digit_part = part[:digit_end_index].zfill(2)
            # Keep the letter portion as-is
            alpha_part = part[digit_end_index:]
                
            processed_parts.append(digit_part + alpha_part)
        elif has_alpha:
            # Pure letters - keep as-is
            processed_parts.append(part)
        else:
            # Other cases - keep as-is
            processed_parts.append(part)
    
    # 拼接结果
    return mapped_prefix + "".join(processed_parts)