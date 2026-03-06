#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据处理模块 - 负责处理标题和脚注的拆分及特殊字符转义
"""

import re

def process_title_split(results: list) -> tuple:
    """
    处理正文标题拆分，返回处理后的结果列表和最大行数
    """
    # 首先找出正文标题中换行符的最大数量，确定需要多少个title列
    max_title_lines = 0
    for result in results:
        # 使用处理后的标题进行拆分计算
        title_text = result.get("processed_title", result["正文标题"])
        if title_text:
            line_count = title_text.count('\n') + 1
            #if line_count > 1:
                #print(f"标题拆分: 标题 '{result['目录标题']}' 包含 {line_count} 行")
            max_title_lines = max(max_title_lines, line_count)
    
    # 为每个结果添加title列（基于处理后的标题）
    for result in results:
        # 优先使用处理后的标题，如果没有则使用原始标题
        main_title = result.get("processed_title", result["正文标题"])
        if main_title:
            # 按换行符拆分
            lines = main_title.split('\n')
            # 填充title1, title2, ...列
            for i, line in enumerate(lines, 1):
                result[f"title{i}"] = line
            # 如果拆分的行数少于最大行数，用空字符串填充剩余的title列
            for i in range(len(lines) + 1, max_title_lines + 1):
                result[f"title{i}"] = ""
        else:
            # 如果正文标题为空，所有title列都设为空
            for i in range(1, max_title_lines + 1):
                result[f"title{i}"] = ""
    
    # 处理title列中的%符号替换
    # 对所有title列中的内容，将%替换为(*ESC*){unicode 0025}
    for result in results:
        for i in range(1, max_title_lines + 1):
            title_col = f"title{i}"
            if title_col in result and result[title_col]:
                # 将%替换为(*ESC*){unicode 0025}
                result[title_col] = result[title_col].replace('%', '(*ESC*){unicode 0025}')
    
    return results, max_title_lines

def check_xxx_patterns(results: list) -> list:
    """
    检查processed_footnote中是否包含XXX模式（不限大小写，数量≥2）
    如果发现则在备注中补充说明
    """
    # XXX模式的正则表达式（不区分大小写，至少2个X）
    xxx_pattern = re.compile(r'[xX]{2,}')
    
    for result in results:
        processed_footnote = result.get("processed_footnote", "")
        if not processed_footnote:
            continue
            
        # 检查是否存在XXX模式
        if xxx_pattern.search(processed_footnote):
            # 构造提醒信息
            warning_msg = "注意：footnote存在XX，请更新"
            
            # 更新备注字段
            if result.get("备注"):
                result["备注"] += "；" + warning_msg
            else:
                result["备注"] = warning_msg
    
    return results

def process_footnote_split(results: list, max_cols: int = 7) -> tuple:
    """
    处理脚注拆分，返回处理后的结果列表和实际使用的最大行数
    基于processed_footnote列进行拆分，确保编码版本替换后的结果能正确反映在拆分列中
    支持最大列数限制：超出的列将合并到最后一列，使用(*ESC*){newline}分隔
    """
    # 找出脚注中换行符的最大数量，确定需要多少个footnote列
    max_footnote_lines = 0
    for result in results:
        # 使用处理后的脚注进行拆分计算
        footnote_text = result.get("processed_footnote", result["脚注"])
        if footnote_text:
            line_count = footnote_text.count('\n') + 1
            max_footnote_lines = max(max_footnote_lines, line_count)
    
    # 确定实际使用的列数（不超过最大限制）
    actual_max_cols = min(max_footnote_lines, max_cols)
    
    # 为每个结果添加footnote列（基于处理后的脚注）
    for result in results:
        # 优先使用处理后的脚注，如果没有则使用原始脚注
        footnote_text = result.get("processed_footnote", result["脚注"])
        if footnote_text:
            # 按换行符拆分
            lines = footnote_text.split('\n')
            
            # 处理列数限制
            if len(lines) <= max_cols:
                # 拆分行数不超过限制，正常拆分
                for i, line in enumerate(lines, 1):
                    result[f"footnote{i}"] = line
                # 填充剩余的空列
                for i in range(len(lines) + 1, actual_max_cols + 1):
                    result[f"footnote{i}"] = ""
            else:
                # 拆分行数超过限制，前max_cols-1列正常拆分，最后一列合并剩余内容
                # 前max_cols-1列正常处理
                for i in range(1, max_cols):
                    result[f"footnote{i}"] = lines[i-1] if i <= len(lines) else ""
                
                # 最后一列合并剩余的所有行，使用(*ESC*){newline}分隔
                remaining_lines = lines[max_cols-1:]  # 从第max_cols行开始的所有剩余行
                merged_content = "(*ESC*){newline}".join(remaining_lines)
                result[f"footnote{max_cols}"] = merged_content
                
                # 如果实际需要的列数小于max_cols，则填充空列
                for i in range(max_cols + 1, actual_max_cols + 1):
                    result[f"footnote{i}"] = ""
        else:
            # 如果脚注为空，所有footnote列都设为空
            for i in range(1, actual_max_cols + 1):
                result[f"footnote{i}"] = ""
    
    # 处理footnote列中的%符号替换
    # 对所有footnote列中的内容，将%替换为(*ESC*){unicode 0025}
    for result in results:
        for i in range(1, actual_max_cols + 1):
            footnote_col = f"footnote{i}"
            if footnote_col in result and result[footnote_col]:
                # 将%替换为(*ESC*){unicode 0025}
                result[footnote_col] = result[footnote_col].replace('%', '(*ESC*){unicode 0025}')
    
    return results, actual_max_cols

def process_footnote_encoding_version(results: list) -> list:
    """
    处理拆分后的footnote编码版本替换
    
    规则：
    - 匹配格式："编码版本：XXXXX。" (X不限制数量和大小写)
    - 根据标题关键词决定替换值：
      * MedDRA组：'非药物治疗'、'手术'、'病史'、'不良事件'、'AE'、'系统器官分类'、'SOC'、"PT"、"首选语"、'首选术语'
      * WHO Drug组：'既往用药'、'合并用药'、'药物治疗'（不含'非药物治疗'）、'ATC'、'按治疗分类/化学分类'、'按治疗分类/化学物质'
    - MedDRA优先级高于WHO Drug
    - 更新备注列说明替换情况
    """
    # 定义关键词组
    meddra_keywords = ['非药物治疗', '手术', '病史', '不良事件', 'AE', '系统器官分类', 'SOC', 'PT', '首选语', '首选术语']
    whodrug_keywords = ['既往用药', '合并用药', '药物治疗', 'ATC', '按治疗分类/化学分类', '按治疗分类/化学物质']
    
    # 排除词：药物治疗不能包含非药物治疗
    exclude_keywords = ['非药物治疗']
    
    # 编码版本匹配正则表达式
    pattern = r'编码版本：[^。]*。'
    
    # 处理每个结果
    for result in results:
        # 获取原始标题用于关键词匹配
        title = result.get("目录标题", "")
        if not title:
            continue
            
        # 获取处理后的脚注文本
        processed_footnote = result.get("processed_footnote", "")
        if not processed_footnote:
            continue
            
        # 检查是否存在编码版本格式
        if re.search(pattern, processed_footnote):
            # 检查标题中包含的关键词
            title_contains_meddra = any(keyword in title for keyword in meddra_keywords)
            title_contains_whodrug = any(keyword in title for keyword in whodrug_keywords if keyword not in exclude_keywords)
            
            # 确定替换值（MedDRA优先级更高）
            replacement = None
            if title_contains_meddra:
                replacement = "&meddra."
            elif title_contains_whodrug:
                replacement = "&whodrug."
            
            # 如果有替换值，则执行替换
            if replacement:
                # 执行替换
                old_footnote = processed_footnote
                result["processed_footnote"] = re.sub(pattern, f'编码版本：{replacement}。', processed_footnote)
                
                # 更新备注列
                if result["备注"]:
                    result["备注"] += "; "
                result["备注"] += f"更新编码版本为{replacement}"

    return results

def reorder_columns(df, max_title_lines: int, max_footnote_lines: int):
    """
    重新排列DataFrame的列顺序
    """
    # 调整列顺序：将新增列放在Output Name列之前
    base_columns = ["序号", "标题", "脚注", "脚注匹配状态", "备注", 
                   "Batch", "Type", "TLF", "Prgm Name", "Output Name"]
    title_columns = [f"title{i}" for i in range(1, max_title_lines + 1)]
    footnote_columns = [f"footnote{i}" for i in range(1, max_footnote_lines + 1)]
    all_columns = base_columns + title_columns + footnote_columns
    
    # 重新排列列的顺序
    return df[all_columns]