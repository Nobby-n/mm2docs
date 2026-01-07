# -*- coding: utf-8 -*-
"""mmclass.py
"""
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Any, Optional
from openpyxl.styles import PatternFill, Font

@dataclass
class USDMConfig:
    """USDM出力用設定
    """
    excel_file_ext: str = '*.xlsx'
    usdm_file_pfx: str = '_USDM.xlsx'
    usdm_sheet: str = 'USDM'
    tmpl_file: str = './templates/USDM_Template.xlsx'
    tmpl_sheet: str = 'Template'
    lv1_title_cell: str = 'B1'
    row_start: int = 4
    col_req: int = 2
    col_sub: int = 3
    col_grp: int = 4
    col_reason: int = 5
    col_remark: int = 7
    col_spec: int = 5
    col_module_start: int = 9
    freeze_panes: str = 'F4'
    prohibit_char: str = r"[ '*/:?[\]`’＊／：？［＼］￥]+"
    sub_txt: str = '_'
    cell_style: Dict[str, List[List[Any]]] = field(default_factory=dict)

@dataclass
class WordConfig:
    """Word出力用設定
    """
    tmpl_name: str = './templates/SpecTemplate.docx'
    file_name: str = '機能仕様書_r0.docx'
    max_head_lv: int = 4
    head_lv_offset: int = 1
    txt_remark: str = 'Normal' # 実際はUSDMcfgで定義されているが値が使われていない箇所がある
    txt_reason: str = 'Normal'

@dataclass
class MdConfig:
    """Markdown出力用設定
    """
    header: str = '<link rel="stylesheet" href="style.css" />\n\n'
    section_remark: str = '<section class="remark">\n'
    section_reason: str = '<section class="reason">\n'
    sec_close: str = '</section>\n\n'
    escape_strs: List[str] = field(default_factory=lambda: ['\\', '*', '_', '#', '+', '-', '.', '!', '{', '}', '[', ']', '(', ')'])

@dataclass
class MMNode:
    """Mindmapのノードデータ
    """
    text: str
    level: int
    note: str = ""

