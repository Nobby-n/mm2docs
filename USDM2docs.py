# -*- coding: utf-8 -*-
"""USDM2docs.py:

## 基本機能
USDM (Excel) から Word, Markdown 形式のドキュメントを生成する。

## 利用方法
python USDM2docs.py [USDMファイルパス] -o [出力ファイルパス]
"""
import os
import sys
import argparse
import openpyxl
import mmconfig as cfg
from mmclass import MMNode
from mm2docs import MMConverter

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

        # シートごとの処理
        for sheet in wb.worksheets:
            # USDM_ で始まるシートのみ処理対象とする
            if not sheet.title.startswith(cfg.usdm_cfg.usdm_sheet + "_"):
                continue

            # B1セル（タイトル）のチェック
            lv1_text = sheet[cfg.usdm_cfg.lv1_title_cell].value
            if not lv1_text:
                continue
            
            # Level 1 ノード追加
            self.nodes.append(MMNode(str(lv1_text), cfg.LV1))

            # 行ごとの処理
            for row in sheet.iter_rows(min_row=cfg.usdm_cfg.row_start, values_only=True):
                # 各列の値を取得 (indexは0始まり)
                # mmconfigのcol_xxxは1始まりなので -1 する
                
                val_req = row[cfg.usdm_cfg.col_req - 1]      # 要求ラベル列
                val_req_txt = row[cfg.usdm_cfg.col_req + 1]  # 要求テキスト列
                
                val_sub = row[cfg.usdm_cfg.col_sub - 1]      # サブ要求ラベル列
                val_sub_txt = row[cfg.usdm_cfg.col_sub + 1]  # サブ要求テキスト列
                
                val_grp = row[cfg.usdm_cfg.col_grp - 1]      # 仕様グループ列
                
                val_reason_lbl = row[cfg.usdm_cfg.col_reason - 2] # 理由ラベル列
                val_reason_txt = row[cfg.usdm_cfg.col_reason - 1] # 理由テキスト列
                
                val_spec_txt = row[cfg.usdm_cfg.col_spec - 1]     # 仕様テキスト列
                
                val_remark = row[cfg.usdm_cfg.col_remark - 1]     # 備考列

                # ノード判定と追加
                # 優先順位: 要求 > サブ要求 > グループ > 理由 > 仕様
                
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

        return True

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
