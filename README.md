# mm2docs: Mindmap to Documents Converter

## 概要

Mindmap (Xmind8, FreePlane) から、USDM形式のExcel、機能仕様書形式のWord、およびMarkDown形式のドキュメントを自動生成するツールです。

Mindmapを起点に仕様書を作成する `mm2docs.py` を主軸とし、既存のUSDM資産を活用するための補助ツールとして以下を同梱しています。

- **`USDM2docs.py`**: USDM形式のExcelファイルから、WordやMarkdown形式の仕様書を生成します。
- **`USDM2mm.py`**: USDM形式のExcelファイルから、Mindmap（FreePlane/FreeMind形式）を生成（逆変換）します。

## 特徴

- **複数のMindmap形式に対応**: Xmind8 (.xmind) および FreePlane (.mm) からデータを抽出可能。
- **3種類の出力形式**: 出力ファイルの拡張子 (.xlsx, .docx, .md) に応じて適切なフォーマットで出力。
- **USDM出力**: 要件定義手法であるUSDM形式のExcelファイルを生成。
- **Word/Markdown出力**: テンプレートに基づいた仕様書や、AIプロンプトにも活用できるMarkdownを生成。
- **USDMからの再変換**: `USDM2docs.py`を使い、手修正したUSDM ExcelからでもWordやMarkdownを再生成可能。
- **Mindmapへの逆変換**: `USDM2mm.py`を使い、既存のUSDMをMindmapで可視化し、レビューや再編集を容易に。

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

### 1. `mm2docs`: Mindmapからドキュメントを生成

#### 基本的な使い方

```bash
python mm2docs.py [Mindmapファイルパス] -o [出力ファイルパス]
```

#### 出力形式の指定

出力形式は `-o` オプションで指定する拡張子によって自動判別されます。

- `.xlsx`: USDM Excel形式
- `.docx`: Word形式
- `.md`: Markdown形式

#### テンプレートの指定

ExcelとWordの出力にはテンプレートファイルが必要です。

```bash
python mm2docs.py sample.xmind -t templates/USDM_Template.xlsx -o sample_USDM.xlsx
python mm2docs.py sample.xmind -t templates/SpecTemplate.docx -o sample_spec.docx
```

#### バッチファイルでの利用

ドラッグ＆ドロップで変換可能なバッチファイルも同梱しています。

- `convert_mm2USDM.bat`: USDM Excel形式へ変換
- `convert_mm2doc.bat`: Word形式へ変換
- `convert_mm2Md.bat`: Markdown形式へ変換

### 2. `USDM2docs`:USDMファイルからWord、Markdownへの変換

コマンドプロンプトから以下のように実行します。

#### 基本的な使い方

```bash
python USDM2docs.py [USDMファイルパス] -o [出力ファイルパス]
```

#### 出力形式の指定

出力形式は `-o` オプションで指定する拡張子によって自動判別されます。

- `.docx`: Word形式
- `.md`: Markdown形式

#### テンプレートの指定

Wordの出力にはテンプレートファイルが必要です。

```bash
python USDM2docs.py sample_USDM.xlsx -t templates/SpecTemplate.docx -o sample_spec.docx
```

#### バッチファイルでの利用

ドラッグ＆ドロップで変換可能なバッチファイルも同梱しています。

- `convert_USDM2doc.bat`: Word形式へ変換
- `convert_USDM2Md.bat`: Markdown形式へ変換

### 3. `USDM2mm`: USDMファイルからMindmap（Freemind形式）への逆変換

コマンドプロンプトから以下のように実行します。

#### 基本的な使い方

```bash
python USDM2mm.py [USDMファイルパス] -o [出力ファイルパス]
```

#### バッチファイルでの利用

ドラッグ＆ドロップで変換可能なバッチファイルも同梱しています。

- `convert_USDM2mm.bat`: USDMからFreeplane形式へ変換

## 設定

`mmconfig.py` および `mmclass.py` で、抽出ルール（正規表現）やExcelの書式、Wordのスタイルなどをカスタマイズできます。

## Mindmapの表記法

[HowToUse_mm2docs.md](HowToUse_mm2docs.md)参照

## 要件定義の考え方

[Concept_RequirementDefinition.md](Concept_RequirementDefinition.md)参照

## サンプル

`samples/` フォルダに、サンプルのMindmapと、それから出力したMarkdownファイルを格納しています。

- [Xmind2USDM_Refactoring_SRS.xmind](samples/Xmind2USDM_Refactoring_SRS.xmind) (Xmind8形式)
- [Xmind2USDM_Refactoring_SRS.mm](samples/Xmind2USDM_Refactoring_SRS.mm) (FreePlane/FreeMind形式)
- [Xmind2USDM_Refactoring_SRS.md](samples/Xmind2USDM_Refactoring_SRS.md) (出力例)

## ライセンス

MIT License
