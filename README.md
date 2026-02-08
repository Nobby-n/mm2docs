# mm2docs: Mindmap to Documents Converter

## 概要

Mindmap (Xmind8, FreePlane) から、USDM形式のExcel、仕様書形式のWord、およびMarkDown形式のドキュメントを自動生成するツールです。

## 特徴

- **複数のMindmap形式に対応**: Xmind8 (.xmind) および FreePlane (.mm) からデータを抽出可能。
- **3種類の出力形式**: 出力ファイルの拡張子 (.xlsx, .docx, .md) に応じて適切なフォーマットで出力。
- **USDM出力**: 要件定義手法であるUSDM形式のExcelファイルを生成。
- **Word出力**: テンプレートに基づいた仕様書の雛形を生成。
- **Markdown出力**: AIプロンプトやドキュメント管理に便利なMarkdown形式を生成。

## インストール

Python 3.12以上が必要です。必要なライブラリをインストールしてください。

```bash
pip install openpyxl xmind-sdk python-docx
```

## 検証環境

以下の環境で動作を確認しています。

- **OS**: Windows 11 (24H2)
- **Python**: 3.12.3

## 利用方法

コマンドプロンプトから以下のように実行します。

### 基本的な使い方

```bash
python mm2docs.py [Mindmapファイルパス] -o [出力ファイルパス]
```

### 出力形式の指定

出力形式は `-o` オプションで指定する拡張子によって自動判別されます。

- `.xlsx`: USDM Excel形式
- `.docx`: Word形式
- `.md`: Markdown形式

### テンプレートの指定

ExcelとWordの出力にはテンプレートファイルが必要です。

```bash
python mm2docs.py sample.xmind -t templates/USDM_Template.xlsx -o sample_USDM.xlsx
python mm2docs.py sample.xmind -t templates/SpecTemplate.docx -o sample_spec.docx
```

### バッチファイルでの利用

ドラッグ＆ドロップで変換可能なバッチファイルも同梱しています。

実行環境に合わせて編集して利用してください。

- `convert_mm2USDM.bat`: USDM Excel形式へ変換
- `convert_mm2doc.bat`: Word形式へ変換
- `convert_mm2Md.bat`: Markdown形式へ変換

## 設定

`mmconfig.py` および `mmclass.py` で、抽出ルール（正規表現）やExcelの書式、Wordのスタイルなどをカスタマイズできます。

## ドキュメント

- **[表記法・リファレンス](Notation_RequirementMindmap.md)**: Mindmapのタグ付けルールと、各フォーマットへの変換イメージ。
- **[要件定義のプロセス](Concept_RequirementDefinition.md)**: Mindmapを使った分析からUSDM作成までの流れと、要件定義の考え方。

## ライセンス

MIT License
