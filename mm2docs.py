# -*- coding: utf-8 -*-
"""mm2docs.py:

## 基本機能
Mindmap (Xmind, FreePlane) から USDM (Excel), Word, Markdown 形式のドキュメントを生成する。

## 利用方法
README.md 参照
"""
import os
import sys
import re
import argparse
import copy
import xml.etree.ElementTree as ET
from typing import List, Tuple, Optional

import openpyxl
from openpyxl.utils import get_column_letter
from xmind import load as load_xmind_sdk
from docx import Document

import mmconfig as cfg
from mmclass import MMNode

class MMConverter:
    """Mindmapデータを読み取り、各種ドキュメントに変換するクラス
    """
    def __init__(self, mm_file: str, template: str, output: str):
        """
        Parameters
        ----------
        mm_file : str
            Source Mindmap file path (.xmind or .mm)
        template : str
            Template file path (.xlsx or .docx)
        output : str
            Output file path (.xlsx, .docx, or .md)
        """
        self.mm_file = mm_file
        self.template = template
        self.output = output
        self.nodes: List[MMNode] = []
        self.doc_data: List[Tuple[str, int]] = [] # Word/MD用 [text, level]

    def load_mm_data(self) -> bool:
        """Mindmapファイルからデータを読み込む
        """
        ext = os.path.splitext(self.mm_file)[1].lower()
        if ext == '.xmind':
            return self._load_xmind()
        elif ext == '.mm':
            return self._load_freeplane()
        else:
            print(f"サポートされていない拡張子です: {ext}")
            return False

    def _load_xmind(self) -> bool:
        """Xmindファイルを読み込む
        """
        try:
            wb = load_xmind_sdk(self.mm_file)
            ws = wb.getPrimarySheet()
            rt = ws.getRootTopic()
            
            raw_nodes = []
            self._get_all_xmind_nodes(rt.getSubTopics(), raw_nodes)
            self._parse_raw_nodes(raw_nodes)
            return True
        except Exception as e:
            print(f"Xmindファイルの読み込みに失敗しました: {e}")
            return False

    def _get_all_xmind_nodes(self, topics, raw_nodes):
        """Xmindのトピックを再帰的に取得してリストに追加する

        Parameters
        ----------
        topics : list
            XMindのトピックオブジェクトのリスト
        raw_nodes : list
            取得したノードのテキストを格納するリスト
        """
        for topic in topics:
            raw_nodes.append(topic.getTitle())
            if topic.getSubTopics():
                self._get_all_xmind_nodes(topic.getSubTopics(), raw_nodes)

    def _load_freeplane(self) -> bool:
        """FreePlane (.mm) ファイルを読み込む
        """
        try:
            tree = ET.parse(self.mm_file)
            root = tree.getroot()
            
            raw_nodes = []
            # FreePlaneのルートノードの下のノードから取得開始
            for node in root.findall('.//node'):
                text = node.get('TEXT')
                if text:
                    raw_nodes.append(text)
            
            self._parse_raw_nodes(raw_nodes)
            return True
        except Exception as e:
            print(f"FreePlaneファイルの読み込みに失敗しました: {e}")
            return False

    def _parse_raw_nodes(self, raw_nodes: List[str]):
        """生のノード文字列を解析してレベルを決定し、MMNodeリストに格納する

        Parameters
        ----------
        raw_nodes : List[str]
            解析対象のノードテキストのリスト
        """
        stcomment = False
        for node_text in raw_nodes:
            if not node_text:
                continue
                
            for pattern, level in cfg.PTNLV:
                match = re.match(pattern, node_text)
                if match:
                    # match/case を使用してリファクタリング
                    match level:
                        case cfg.LVNOTEE:
                            if stcomment:
                                stcomment = False
                            continue
                        case _ if stcomment:
                            continue
                        case cfg.LVNOTE:
                            continue
                        case cfg.LVNOTES:
                            if not stcomment:
                                stcomment = True
                            continue
                        case l if l <= cfg.LV4:
                            # 見出しレベル1無しで下位が現れたら補完
                            if not self.nodes and l > cfg.LV1:
                                self.nodes.append(MMNode("xxxx xxxx", cfg.LV1))
                            
                            text = match.group(2) if len(match.groups()) >= 2 else node_text
                            self.nodes.append(MMNode(text, l))
                        case cfg.LVREASON:
                            if self.nodes:
                                text = match.group(2) if len(match.groups()) >= 2 else node_text
                                self.nodes.append(MMNode(text.replace('\r\n', '\n'), cfg.LVREASON))
                        case cfg.LVREMARK:
                            if self.nodes:
                                text = match.group(2) if len(match.groups()) >= 2 else node_text
                                # 連続する備考は連結
                                if self.nodes[-1].level == cfg.LVREMARK:
                                    self.nodes[-1].text += f"\n{text}"
                                else:
                                    self.nodes.append(MMNode(text.replace('\r\n', '\n'), cfg.LVREMARK))
                        case cfg.LVSPEC:
                            if self.nodes:
                                self.nodes.append(MMNode(node_text.replace('\r\n', '\n'), cfg.LVSPEC))
                    break

    def save(self) -> bool:
        """指定された拡張子に応じて保存処理を行う
        """
        ext = os.path.splitext(self.output)[1].lower()
        match ext:
            case '.xlsx':
                return self.save_xlsx()
            case '.docx':
                return self.save_docx()
            case '.md':
                return self.save_md()
            case _:
                print(f"サポートされていない出力形式です: {ext}")
                return False

    def save_xlsx(self) -> bool:
        """USDM Excel形式で保存
        """
        print(f"Excel (USDM) 出力中: {self.output}")
        if not os.path.splitext(self.template)[1].lower() == '.xlsx':
            print("エラー: テンプレートファイルは .xlsx 形式である必要があります。")
            return False

        try:
            wb = openpyxl.load_workbook(self.template)
            tmpl_sheet = wb[cfg.usdm_cfg.tmpl_sheet]
            
            # 書式情報の取得
            default_styles = []
            for col in range(1, tmpl_sheet.max_column + 1):
                default_styles.append(tmpl_sheet.cell(row=cfg.usdm_cfg.row_start, column=col)._style)

            scnt = 1
            sn_list = []
            
            # データの転記
            current_sheet = None
            rownum = cfg.usdm_cfg.row_start
            id_cnt = [0, 0, 0]
            sn = ""

            for node in self.nodes:
                match node.level:
                    case cfg.LV1:
                        # 新しいシートの作成
                        current_sheet = wb.copy_worksheet(tmpl_sheet)
                        
                        # シート名設定
                        kw, lv1_txt = self._get_lv1_info(node.text)
                        sn = kw
                        if sn in sn_list:
                            for i in range(1, 101):
                                if f"{sn}{i}" not in sn_list:
                                    sn = f"{sn}{i}"
                                    break
                        sn_list.append(sn)
                        
                        current_sheet.title = f"{cfg.usdm_cfg.usdm_sheet}_{sn}"
                        wb.move_sheet(current_sheet, offset=-wb.index(current_sheet) + scnt)
                        scnt += 1
                        
                        # 初期化
                        current_sheet.delete_rows(cfg.usdm_cfg.row_start, current_sheet.max_row + 1)
                        current_sheet[cfg.usdm_cfg.lv1_title_cell].value = lv1_txt
                        id_cnt = [0, 0, 0]
                        rownum = cfg.usdm_cfg.row_start - 1
                        
                        # ウィンドウ枠の固定
                        current_sheet.freeze_panes = cfg.usdm_cfg.freeze_panes

                    case cfg.LVREMARK:
                        if current_sheet:
                            # 直前の行に備考をセット
                            self._set_cell(current_sheet, rownum, cfg.usdm_cfg.col_remark, node.text, cfg.TXTREMARK, default_styles)
                            rownum -= 1 # 備考は行を増やさない

                    case cfg.LV2:
                        id_cnt = [id_cnt[0] + 1, 0, 0]
                        data = [""] * 7
                        data[cfg.usdm_cfg.col_req - 1] = cfg.TXTREQ2
                        data[cfg.usdm_cfg.col_req] = f"{sn}-{str(id_cnt[0]).zfill(2)}"
                        data[cfg.usdm_cfg.col_req + 1] = node.text
                        self._set_row(current_sheet, rownum, data, cfg.TXTREQ2, default_styles)

                    case cfg.LV3:
                        id_cnt = [id_cnt[0], id_cnt[1] + 1, 0]
                        data = [""] * 7
                        data[cfg.usdm_cfg.col_sub - 1] = cfg.TXTREQ3
                        data[cfg.usdm_cfg.col_sub] = f"{sn}-{str(id_cnt[0]).zfill(2)}-{str(id_cnt[1]).zfill(2)}"
                        data[cfg.usdm_cfg.col_sub + 1] = node.text
                        self._set_row(current_sheet, rownum, data, cfg.TXTREQ3, default_styles)

                    case cfg.LV4:
                        data = [""] * 7
                        data[cfg.usdm_cfg.col_grp - 1] = f"＜{node.text}＞"
                        self._set_row(current_sheet, rownum, data, cfg.TXTREQ4, default_styles)

                    case cfg.LVREASON:
                        data = [""] * 7
                        data[cfg.usdm_cfg.col_reason - 2] = cfg.TXTREASON
                        data[cfg.usdm_cfg.col_reason - 1] = node.text
                        self._set_row(current_sheet, rownum, data, cfg.TXTREASON, default_styles)

                    case cfg.LVSPEC:
                        id_cnt = [id_cnt[0], id_cnt[1], id_cnt[2] + 1]
                        data = [""] * 7
                        data[cfg.usdm_cfg.col_spec - 2] = f"{sn}-{str(id_cnt[0]).zfill(2)}-{str(id_cnt[1]).zfill(2)}-{str(id_cnt[2]*10).zfill(3)}"
                        data[cfg.usdm_cfg.col_spec - 1] = node.text
                        self._set_row(current_sheet, rownum, data, cfg.TXTSPEC, default_styles)

                if current_sheet:
                    # 最初のデータ行（4行目）の高さ自動調整修正
                    if rownum == 4:
                        current_sheet.row_dimensions[rownum].height = None
                    rownum += 1

            wb.save(self.output)
            return True
        except Exception as e:
            print(f"Excel出力中にエラーが発生しました: {e}")
            return False

    def _get_lv1_info(self, txt: str) -> Tuple[str, str]:
        """大要求レベルのノード情報をキーワードとテキストに分割する

        Parameters
        ----------
        txt : str
            大要求ノードのテキスト

        Returns
        -------
        kw : str
            キーワード（シート名に使用、禁止文字置換済み）
        lv1_txt : str
            大要求の内容テキスト
        """
        try:
            parts = txt.split(None, 1)
            kw_raw = parts[0]
            lv1_txt = parts[1] if len(parts) > 1 else parts[0]
        except:
            kw_raw = lv1_txt = txt
        
        kw = re.sub(cfg.usdm_cfg.prohibit_char, cfg.usdm_cfg.sub_txt, kw_raw)
        return kw, lv1_txt

    def _set_row(self, sheet, rnum, rdata, lv_name, default_styles):
        """Excelシートに1行分のデータを書き込む

        Parameters
        ----------
        sheet : openpyxl.worksheet.worksheet.Worksheet
            書き込み対象のシート
        rnum : int
            行番号
        rdata : list
            1行分のデータリスト
        lv_name : str
            レベル名（書式決定用）
        default_styles : list
            テンプレートから取得したデフォルト書式のリスト
        """
        for cn, txt in enumerate(rdata):
            self._set_cell(sheet, rnum, cn + 1, txt, lv_name, default_styles)
        # 関連モジュール列の書式セット
        for i, style in enumerate(default_styles[cfg.usdm_cfg.col_module_start-1:]):
            cnum = cfg.usdm_cfg.col_module_start + i
            sheet.cell(row=rnum, column=cnum)._style = copy.deepcopy(style)

    def _set_cell(self, sheet, rnum, cnum, txt, lv_name, default_styles):
        """Excelシートの特定のセルにデータと書式をセットする

        Parameters
        ----------
        sheet : openpyxl.worksheet.worksheet.Worksheet
            書き込み対象のシート
        rnum : int
            行番号
        cnum : int
            列番号
        txt : str
            セットするテキスト
        lv_name : str
            レベル名（書式決定用）
        default_styles : list
            テンプレートから取得したデフォルト書式のリスト
        """
        cell = sheet.cell(row=rnum, column=cnum)
        cell.value = txt
        cell._style = copy.deepcopy(default_styles[cnum - 1])
        if lv_name in cfg.usdm_cfg.cell_style:
            styles = cfg.usdm_cfg.cell_style[lv_name]
            if cnum - 1 < len(styles):
                cell.fill = styles[cnum - 1][0]
                cell.font = styles[cnum - 1][1]

    def save_docx(self) -> bool:
        """Word形式で保存
        """
        print(f"Word出力中: {self.output}")
        if not os.path.splitext(self.template)[1].lower() == '.docx':
            print("エラー: テンプレートファイルは .docx 形式である必要があります。")
            return False

        try:
            doc = Document(self.template)
            for node in self.nodes:
                match node.level:
                    case l if l <= cfg.word_cfg.max_head_lv:
                        doc.add_paragraph('') # 見出しの前に改行
                        doc.add_heading(node.text, level=l + cfg.word_cfg.head_lv_offset)
                    case cfg.LVREMARK:
                        doc.add_paragraph(node.text, style=cfg.word_cfg.txt_remark)
                    case cfg.LVREASON:
                        doc.add_paragraph(node.text, style=cfg.word_cfg.txt_reason)
                    case _:
                        doc.add_paragraph(node.text)
            
            doc.save(self.output)
            return True
        except Exception as e:
            print(f"Word出力中にエラーが発生しました: {e}")
            return False

    def save_md(self) -> bool:
        """Markdown形式で保存
        """
        print(f"Markdown出力中: {self.output}")
        try:
            md_text = cfg.md_cfg.header
            for node in self.nodes:
                text = self._escape_md(node.text)
                # ノード内の改行を <br /> に変換
                text = text.replace('\n', '<br />')
                
                match node.level:
                    case l if l <= cfg.word_cfg.max_head_lv:
                        md_text += f'{"#" * l} {text}\n\n'
                    case cfg.LVREMARK:
                        md_text += f'{cfg.md_cfg.section_remark}{text}\n{cfg.md_cfg.sec_close}'
                    case cfg.LVREASON:
                        md_text += f'{cfg.md_cfg.section_reason}{text}\n{cfg.md_cfg.sec_close}'
                    case _:
                        md_text += f'{text}\n\n'
            
            with open(self.output, 'w', encoding='utf-8') as f:
                f.write(md_text)
            return True
        except Exception as e:
            print(f"Markdown出力中にエラーが発生しました: {e}")
            return False

    def _escape_md(self, txt: str) -> str:
        """Markdownの特殊文字をエスケープする
        ただし、バックトークンで囲まれた部分は除外する

        Parameters
        ----------
        txt : str
            エスケープ対象のテキスト

        Returns
        -------
        str
            エスケープ後のテキスト
        """
        # バックトークンで囲まれた部分はエスケープしない
        # それ以外の部分をエスケープ対象文字列で置換する
        parts = re.split(r'(`[^`]*`)', txt)
        for i in range(len(parts)):
            # 偶数番目の要素（バックトークンの外側）のみエスケープ処理を行う
            if i % 2 == 0:
                for s in cfg.md_cfg.escape_strs:
                    parts[i] = parts[i].replace(s, '\\' + s)
        return "".join(parts)

def main():
    parser = argparse.ArgumentParser(description='Mindmap to Documents Converter')
    parser.add_argument('mm', help='Source Mindmap file (.xmind or .mm)')
    parser.add_argument('-t', '--template', help='Template file (.xlsx or .docx)', default='')
    parser.add_argument('-o', '--output', help='Output file (.xlsx, .docx, or .md)', default='')

    args = parser.parse_args()

    # パス解決
    mm_file = args.mm
    if not os.path.exists(mm_file):
        print(f"エラー: 指定されたファイルが見つかりません: {mm_file}")
        sys.exit(1)

    output = args.output
    if not output:
        # デフォルト出力先の設定
        base = os.path.splitext(mm_file)[0]
        output = base + "_USDM.xlsx" # デフォルトはExcel

    template = args.template
    if not template:
        # 出力形式に応じたデフォルトテンプレート
        ext = os.path.splitext(output)[1].lower()
        if ext == '.xlsx':
            template = cfg.usdm_cfg.tmpl_file
        elif ext == '.docx':
            template = cfg.word_cfg.tmpl_name

    converter = MMConverter(mm_file, template, output)
    if converter.load_mm_data():
        if converter.save():
            print("変換完了！")
        else:
            print("変換失敗。")
            sys.exit(1)
    else:
        print("Mindmapの読み込みに失敗しました。")
        sys.exit(1)

if __name__ == '__main__':
    main()

