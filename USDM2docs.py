# -*- coding: utf-8 -*-
"""USDM2docs.py:

## 基本機能
USDM (Excel) から Word, Markdown 形式のドキュメントを生成する。

## 利用方法
python USDM2docs.py [USDMファイルパス] -o [出力ファイルパス]
"""
import os
import sys
import io
import argparse
import zipfile
import xml.etree.ElementTree as ET

import openpyxl
from PIL import Image as PILImage

import mmconfig as cfg
from mmclass import MMNode
from mm2docs import MMConverter

_XLSX_NS_SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_XLSX_NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_XLSX_NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


class USDMConverter(MMConverter):
    """USDM Excelデータを読み取り、各種ドキュメントに変換するクラス
    """
    def load_mm_data(self) -> bool:
        """USDMファイルからデータを読み込む (Override)
        """
        print(f"Loading USDM: {self.mm_file}")
        if not os.path.exists(self.mm_file):
             print(f"エラー: 指定されたファイルが見つかりません: {self.mm_file}")
             return False
             
        try:
            wb = openpyxl.load_workbook(self.mm_file, data_only=True)
        except Exception as e:
            print(f"Excelファイルの読み込みに失敗しました: {e}")
            return False

        self.nodes = []
        usdm_prefix = cfg.usdm_cfg.usdm_sheet + "_"

        # シートごとの処理
        for sheet in wb.worksheets:
            if not sheet.title.startswith(usdm_prefix):
                continue

            # B1セル（タイトル）のチェック
            lv1_text = sheet[cfg.usdm_cfg.lv1_title_cell].value
            if not lv1_text:
                continue

            kw = sheet.title[len(usdm_prefix):]
            self.nodes.append(
                MMNode(f"{kw} {lv1_text}", cfg.LV1, clean_text=str(lv1_text)))

            for row in sheet.iter_rows(min_row=cfg.usdm_cfg.row_start, values_only=True):
                val_req = row[cfg.usdm_cfg.col_req - 1]
                val_req_txt = row[cfg.usdm_cfg.col_req + 1]
                val_sub = row[cfg.usdm_cfg.col_sub - 1]
                val_sub_txt = row[cfg.usdm_cfg.col_sub + 1]
                val_grp = row[cfg.usdm_cfg.col_grp - 1]
                val_reason_lbl = row[cfg.usdm_cfg.col_reason - 2]
                val_reason_txt = row[cfg.usdm_cfg.col_reason - 1]
                val_spec_txt = row[cfg.usdm_cfg.col_spec - 1]
                val_remark = row[cfg.usdm_cfg.col_remark - 1]

                match (val_req, val_sub, val_grp, val_reason_lbl):
                    case (cfg.TXTREQ2, _, _, _) if val_req_txt:
                        self.nodes.append(MMNode(str(val_req_txt), cfg.LV2))
                    case (_, cfg.TXTREQ3, _, _) if val_sub_txt:
                        self.nodes.append(MMNode(str(val_sub_txt), cfg.LV3))
                    case (_, _, grp, _) if grp and str(grp).startswith('＜') and str(grp).endswith('＞'):
                        self.nodes.append(MMNode(str(grp)[1:-1], cfg.LV4))
                    case (_, _, _, cfg.TXTREASON) if val_reason_txt:
                        self.nodes.append(MMNode(str(val_reason_txt), cfg.LVREASON))
                    case _:
                        if val_spec_txt:
                            # 他の要素でない場合、仕様とみなす
                            self.nodes.append(MMNode(str(val_spec_txt), cfg.LVSPEC))

                # 備考の追加 (行の最後に処理)
                if val_remark:
                    self.nodes.append(MMNode(str(val_remark), cfg.LVREMARK))

        self._load_usdm_images()
        return True

    def _load_usdm_images(self):
        """mm_imagesフォルダの画像をノードに関連付ける。
        フォルダが無い、または画像ファイルが無い場合は
        Excelの画像シートから抽出を試みる。
        """
        usdm_dir = os.path.dirname(os.path.abspath(self.mm_file))
        image_dir = os.path.join(usdm_dir, cfg.IMAGE_DIR)

        if not os.path.isdir(image_dir) or not any(
                f.lower().endswith('.png') for f in os.listdir(image_dir)):
            self._extract_images_from_usdm_xlsx()

        if not os.path.isdir(image_dir):
            return

        image_map: dict[str, dict[str, str]] = {}
        for fname in os.listdir(image_dir):
            if not fname.lower().endswith('.png'):
                continue
            base = os.path.splitext(fname)[0]
            parts = base.split('_', 2)
            if len(parts) >= 3:
                image_map.setdefault(parts[0], {})[parts[2]] = \
                    os.path.join(image_dir, fname)

        if not image_map:
            return

        current_kw = "img"
        for node in self.nodes:
            if node.level == cfg.LV1:
                current_kw, _ = self._get_lv1_info(node.text)
                continue

            kw_images = image_map.get(current_kw)
            if not kw_images:
                continue

            sanitized = self._sanitize_filename(node.text, max_len=50)
            if sanitized in kw_images:
                node.image_path = kw_images[sanitized]

    def _extract_images_from_usdm_xlsx(self):
        """USDM Excelの画像シートから画像を抽出してmm_imagesに保存する。

        xlsxをZIPとして開き、シート→描画→メディアの関連を辿って
        画像データを取り出す。
        """
        usdm_dir = os.path.dirname(os.path.abspath(self.mm_file))
        image_dir = os.path.join(usdm_dir, cfg.IMAGE_DIR)

        try:
            with zipfile.ZipFile(self.mm_file, 'r') as zf:
                name_to_rid = {}
                wb_xml = ET.fromstring(zf.read('xl/workbook.xml'))
                for s in wb_xml.iter(f'{{{_XLSX_NS_SS}}}sheet'):
                    name_to_rid[s.get('name')] = \
                        s.get(f'{{{_XLSX_NS_REL}}}id')

                wb_rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
                rid_to_target = {}
                for rel in wb_rels.iter(f'{{{_XLSX_NS_PKG}}}Relationship'):
                    rid_to_target[rel.get('Id')] = rel.get('Target')

                usdm_prefix = cfg.usdm_cfg.usdm_sheet + '_'
                extracted = False

                for sheet_name, rid in name_to_rid.items():
                    if sheet_name.startswith(usdm_prefix):
                        continue
                    if sheet_name == cfg.usdm_cfg.tmpl_sheet:
                        continue

                    target = rid_to_target.get(rid, '')
                    sheet_base = os.path.basename(target)
                    rels_path = f'xl/worksheets/_rels/{sheet_base}.rels'
                    if rels_path not in zf.namelist():
                        continue

                    drawing_ref = self._find_rel_target(
                        zf, rels_path, 'drawing')
                    if not drawing_ref:
                        continue

                    dr_rels_path = \
                        f'xl/drawings/_rels/{os.path.basename(drawing_ref)}.rels'
                    media_ref = self._find_rel_target(
                        zf, dr_rels_path, 'image')
                    if not media_ref:
                        continue

                    media_path = 'xl/media/' + os.path.basename(media_ref)
                    if media_path not in zf.namelist():
                        continue

                    if not extracted:
                        os.makedirs(image_dir, exist_ok=True)
                        extracted = True

                    img_data = zf.read(media_path)
                    out_path = os.path.join(image_dir, sheet_name + '.png')
                    PILImage.open(io.BytesIO(img_data)).save(out_path, 'PNG')
                    print(f"  Excel画像抽出: {sheet_name}.png")
        except Exception as e:
            print(f"  Excel画像シートの抽出でエラー: {e}")

    @staticmethod
    def _find_rel_target(zf, rels_path: str, type_keyword: str):
        """xlsx内のリレーションシップXMLから指定タイプのTargetを返す"""
        if rels_path not in zf.namelist():
            return None
        rels_xml = ET.fromstring(zf.read(rels_path))
        for rel in rels_xml.iter(f'{{{_XLSX_NS_PKG}}}Relationship'):
            if type_keyword in rel.get('Type', ''):
                return rel.get('Target')
        return None


def main():
    parser = argparse.ArgumentParser(description='USDM to Documents Converter')
    parser.add_argument('usdm', help='Source USDM file (.xlsx)')
    parser.add_argument('-t', '--template', help='Template file (.docx)', default='')
    parser.add_argument('-o', '--output', help='Output file (.docx or .md)', default='')

    args = parser.parse_args()

    usdm_file = args.usdm
    output = args.output
    if not output:
        base = os.path.splitext(usdm_file)[0]
        output = base + ".docx" # Default to Word

    template = args.template
    if not template:
        ext = os.path.splitext(output)[1].lower()
        if ext == '.docx':
            template = cfg.word_cfg.tmpl_name

    converter = USDMConverter(usdm_file, template, output)
    if converter.load_mm_data():
        if converter.save():
            print("変換完了！")
        else:
            print("変換失敗。")
            sys.exit(1)
    else:
        print("USDMの読み込みに失敗しました。")
        sys.exit(1)

if __name__ == '__main__':
    main()
