# -*- coding: utf-8 -*-
"""mmclass.py
"""
from dataclasses import dataclass
from typing import List, Tuple, Dict, Any, Optional

@dataclass
class USDMConfig:
    """USDM出力用設定
    """
    excel_file_ext: str
    usdm_file_pfx: str
    usdm_sheet: str
    tmpl_file: str
    tmpl_sheet: str
    lv1_title_cell: str
    row_start: int
    col_req: int
    col_sub: int
    col_grp: int
    col_reason: int
    col_remark: int
    col_spec: int
    col_module_start: int
    freeze_panes: str
    prohibit_char: str
    sub_txt: str
    cell_style: Dict[str, List[List[Any]]]

@dataclass
class WordConfig:
    """Word出力用設定
    """
    tmpl_name: str
    file_name: str
    max_head_lv: int
    head_lv_offset: int
    txt_remark: str
    txt_reason: str

@dataclass
class MdConfig:
    """Markdown出力用設定
    """
    header: str
    section_remark: str
    section_reason: str
    sec_close: str
    escape_strs: List[str]

@dataclass
class MMNode:
    """Mindmapのノードデータ
    """
    text: str
    level: int
    clean_text: str = ""
    note: str = ""

    def __post_init__(self):
        if not self.clean_text:
            self.clean_text = self.text

