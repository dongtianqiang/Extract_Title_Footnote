#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX标题脚注提取工具包
"""

__version__ = "1.0.0"
__author__ = "Lingma"

# 导出主要功能
from .main import main, process_document
from .config import *
from .toc_extractor import extract_toc_titles
from .title_matcher import find_title_paragraph_index
from .table_locator import find_next_table_after_index
from .output_name_generator import generate_output_name
from .data_processor import process_title_split, process_footnote_split, reorder_columns
from .utils import (
    is_fully_italic, 
    contains_programming_keyword, 
    extract_content_before_keyword,
    process_superscript_subscript_text
)

__all__ = [
    'main',
    'process_document',
    'extract_toc_titles',
    'find_title_paragraph_index',
    'find_next_table_after_index',
    'extract_footnotes',
    'extract_footnote_from_paragraphs',
    'generate_output_name',
    'process_title_split',
    'process_footnote_split',
    'reorder_columns',
    'is_fully_italic',
    'contains_programming_keyword',
    'extract_content_before_keyword',
    'process_superscript_subscript_text'
]