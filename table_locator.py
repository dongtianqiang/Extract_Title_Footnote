#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格关联模块 - 负责查找标题段落后的关联表格
"""

from typing import List, Optional

def find_next_table_after_index(tables, para_idx: int, paragraphs: List) -> Optional[int]:
    """找到在指定段落索引之后的第一个表格"""
    # 获取所有表格在文档中的相对位置（通过XML元素顺序）
    # 方法：遍历 doc.element.body 找到每个 table 的 parent index
    body = paragraphs[0]._element.getparent()  # <w:body>
    table_elements = body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')
    
    # 构建段落与表格的全局顺序映射
    all_elements = []
    for elem in body:
        if elem.tag.endswith('p'):
            all_elements.append(('paragraph', elem))
        elif elem.tag.endswith('tbl'):
            all_elements.append(('table', elem))
    
    # 找到目标段落的全局位置
    target_pos = None
    for idx, (elem_type, elem) in enumerate(all_elements):
        if elem_type == 'paragraph' and elem is paragraphs[para_idx]._element:
            target_pos = idx
            break
    if target_pos is None:
        return None
    
    # 找下一个 table
    for idx in range(target_pos + 1, len(all_elements)):
        if all_elements[idx][0] == 'table':
            # 返回 tables 列表中的索引（doc.tables 顺序与 XML 顺序一致）
            try:
                table_elem = all_elements[idx][1]
                for i, t in enumerate(tables):
                    if t._element is table_elem:
                        return i
            except Exception:
                pass
    return None