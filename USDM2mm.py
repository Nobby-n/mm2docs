# -*- coding: utf-8 -*-
"""USDM2mm.py:

## 基本機能
USDM (Excel) から Freeplane (.mm) 形式のドキュメントを生成する。

## 利用方法
python USDM2mm.py [USDMファイルパス] -o [出力ファイルパス]
"""
import os
import sys
import shutil
import argparse
import xml.etree.ElementTree as ET

import mmconfig as cfg
from mmclass import MMNode
from USDM2docs import USDMConverter


class USDM2MMConverter(USDMConverter):
    """USDM Excelデータを読み取り、Freeplane形式に変換するクラス
    """
    def save(self) -> bool:
        """Freeplane (.mm) 形式で保存
        """
        print(f"Freeplane (.mm) 出力中: {self.output}")
        
        # Freeplane形式の基本構造を作成
        map_root = ET.Element('map', version='freeplane 1.11.1')
        
        # ルートノード（USDMファイル名などを想定）
        root_title = os.path.splitext(os.path.basename(self.mm_file))[0]
        root_node = ET.SubElement(map_root, 'node', TEXT=root_title, ID=f'ID_{root_title}')

        # 階層管理用のスタック
        # [node_object, level]
        stack = [[root_node, 0]]

        output_dir = os.path.dirname(os.path.abspath(self.output))
        images_dir = os.path.join(output_dir, 'images')
        has_images = any(
            n.image_path and os.path.exists(n.image_path) for n in self.nodes)
        if has_images:
            os.makedirs(images_dir, exist_ok=True)

        for node in self.nodes:
            # タグの付与
            tag = ""
            match node.level:
                case cfg.LV1:
                    tag = "# "
                    target_level = 1
                case cfg.LV2:
                    tag = "## "
                    target_level = 2
                case cfg.LV3:
                    tag = "### "
                    target_level = 3
                case cfg.LV4:
                    tag = "#### "
                    target_level = 4
                case cfg.LVREASON:
                    tag = "? "
                    target_level = stack[-1][1] + 1 # 直近のノードの子にする
                case cfg.LVREMARK:
                    tag = "// "
                    target_level = stack[-1][1] + 1 # 直近のノードの子にする
                case cfg.LVSPEC:
                    tag = ""
                    target_level = stack[-1][1] + 1 # 直近のノードの子にする
                case _:
                    continue

            # 適切な親ノードを探す
            while len(stack) > 1 and stack[-1][1] >= target_level:
                stack.pop()
            
            parent = stack[-1][0]
            text_for_node = node.clean_text if node.level == cfg.LV1 else node.text
            new_node = ET.SubElement(parent, 'node', TEXT=tag + text_for_node)
            
            # 階層構造を持つレベル（LV1-LV4）のみスタックに積む
            if node.level <= cfg.LV4:
                stack.append([new_node, target_level])

            if node.image_path and os.path.exists(node.image_path):
                img_filename = os.path.basename(node.image_path)
                shutil.copy2(node.image_path,
                             os.path.join(images_dir, img_filename))
                img_child = ET.SubElement(new_node, 'node')
                rc = ET.SubElement(img_child, 'richcontent', TYPE='NODE')
                html_el = ET.SubElement(rc, 'html')
                ET.SubElement(html_el, 'head')
                body_el = ET.SubElement(html_el, 'body')
                ET.SubElement(body_el, 'img', src=f'images/{img_filename}')
                print(f"  画像出力: images/{img_filename}")

        try:
            tree = ET.ElementTree(map_root)
            # インデントを整える（Python 3.9+）
            ET.indent(tree, space="  ", level=0)
            tree.write(self.output, encoding='utf-8', xml_declaration=True)
            return True
        except Exception as e:
            print(f"Freeplane出力中にエラーが発生しました: {e}")
            return False


def main():
    parser = argparse.ArgumentParser(description='USDM to Freeplane Converter')
    parser.add_argument('usdm', help='Source USDM file (.xlsx)')
    parser.add_argument('-o', '--output', help='Output file (.mm)', default='')

    args = parser.parse_args()

    usdm_file = args.usdm
    output = args.output
    if not output:
        base = os.path.splitext(usdm_file)[0]
        output = base + "_reverse.mm"

    # USDMConverterを継承したUSDM2MMConverterを使用
    # テンプレートは不要なので空文字を渡す
    converter = USDM2MMConverter(usdm_file, "", output)
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
