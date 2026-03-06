#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标题匹配模块 - 负责在文档正文中查找与目录标题匹配的段落
"""

from typing import List, Optional
from toc_extractor import getOutlineLevel, isTitle

def find_title_paragraph_index(paragraphs: List, target_raw_title: str) -> Optional[int]:
    """在段落中查找严格匹配的标题段落索引（使用原始文本）"""
    for i, para in enumerate(paragraphs):
        if para.text == target_raw_title and isTitle(para) is not None:  # 完全相等且必须是标题
            return i
    return None