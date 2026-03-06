#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具函数模块 - 包含各种辅助函数
"""

from docx.text.run import Run

def is_fully_italic(paragraph) -> bool:
    """判断段落是否全为斜体（programming note）"""
    runs_with_text = [run for run in paragraph.runs if run.text.strip()]
    if not runs_with_text:
        return False
    return all(run.italic for run in runs_with_text)

# 为了向后兼容，保留旧函数名但重定向到新函数
def contains_programming_keyword(paragraph) -> bool:
    """判断段落是否包含"programming"或"programmer"关键词（不区分大小写）"""
    return contains_custom_keyword(paragraph)

def contains_custom_keyword(paragraph, custom_keywords=None) -> bool:
    """判断段落是否包含自定义关键词（不区分大小写）
    
    Args:
        paragraph: 段落对象
        custom_keywords: 自定义关键词列表，将与默认的programming/programmer一起使用
    
    Returns:
        bool: 是否包含任意一个关键词
    """
    # 始终包含默认关键词
    from config import DEFAULT_CUSTOM_KEYWORDS
    all_keywords = DEFAULT_CUSTOM_KEYWORDS.copy()
    
    # 如果提供了自定义关键词，添加到默认关键词列表中
    if custom_keywords:
        all_keywords.extend(custom_keywords)
    
    text = paragraph.text.strip()
    if not text:
        return False
    # 转为小写进行比较
    lower_text = text.lower()
    return any(keyword.lower() in lower_text for keyword in all_keywords)

# 为了向后兼容，保留旧函数名但重定向到新函数
def extract_content_before_keyword(paragraph) -> str:
    """提取段落中"programming"/"programmer"关键词之前的内容"""
    return extract_content_before_custom_keyword(paragraph)

def extract_content_before_custom_keyword(paragraph, custom_keywords=None) -> str:
    """提取段落中自定义关键词之前的内容
    
    Args:
        paragraph: 段落对象
        custom_keywords: 自定义关键词列表，将与默认的programming/programmer一起使用
    
    Returns:
        str: 关键词之前的内容
    """
    text = paragraph.text
    return extract_content_before_custom_keyword_from_text(text, custom_keywords)

def extract_content_before_custom_keyword_from_text(text: str, custom_keywords=None) -> str:
    """从文本中提取自定义关键词之前的内容
    
    Args:
        text: 输入文本
        custom_keywords: 自定义关键词列表，将与默认的programming/programmer一起使用
    
    Returns:
        str: 关键词之前的内容
    """
    # 始终包含默认关键词
    from config import DEFAULT_CUSTOM_KEYWORDS
    all_keywords = DEFAULT_CUSTOM_KEYWORDS.copy()
    
    # 如果提供了自定义关键词，添加到默认关键词列表中
    if custom_keywords:
        all_keywords.extend(custom_keywords)
    
    lower_text = text.lower()
    
    # 查找所有关键词的位置
    keyword_positions = []
    for keyword in all_keywords:
        pos = lower_text.find(keyword.lower())
        if pos != -1:
            keyword_positions.append(pos)
    
    # 找到最早出现的关键词位置
    if keyword_positions:
        earliest_pos = min(keyword_positions)
        return text[:earliest_pos].strip()
    else:
        return text.strip()

def process_superscript_subscript_text(paragraph) -> str:
    """
    处理段落中的上标和下标文本
    将上标内容替换为(*ESC*){super 上标内容}
    将下标内容替换为(*ESC*){sub 下标内容}
    """
    processed_text = ""
    
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
            
        # 检查上标格式
        if getattr(run.font, 'superscript', None):
            # 上标：替换为(*ESC*){super 内容}
            processed_text += f"(*ESC*){{super {text}}}"
        # 检查下标格式
        elif getattr(run.font, 'subscript', None):
            # 下标：替换为(*ESC*){sub 内容}
            processed_text += f"(*ESC*){{sub {text}}}"
        else:
            # 普通文本：保持原样
            processed_text += text
    
    return processed_text