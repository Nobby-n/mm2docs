<link rel="stylesheet" href="style.css" />

# overall 概要

## 目的・背景

### リファクタリングの背景

基本機能を実現したツールは`Xmind2USDM.py`として開発済みであるが、WordやMarkDown形式で出力するのにUSDMを出力してから`USDM2Word.py`を使わないとできないのを1つのスクリプト`mm2docs.py`で処理できるようにしたい。<br/>また、データ自体や処理に応じて適切なClass化をして見通しを良くしたい。

### ツール自体の目的

要件をわかりやすく整理し、次工程に伝える手法を確立したい

要件を漏れなくダブり無くまとめるのにMindmapを使いたい。

作成したMindmapを自動変換して決まったフォーマットで出力できるようにしたい。<br/>下記の3種類の出力に対応できること。<br/>\- USDMのExcel形式<br/>\- 仕様書の雛形としてのWord形式<br/>\- AI指示用のプロンプトや汎用ドキュメントとしてのMarkDown形式

## 前提・制約

### 使用技術

Python 3\.12以降で動作するPythonスクリプトで実現する。<br/>Python及び必要なライブラリがインストールしてあるWindows PC上で実行するものとする。<br/>`python.exe`にはpathが通っているものとする。<br/>あるいは、`py.exe`による実行とする。

MindmapはXmind8で作成した\.xmind形式か、FreePlaneで作成した\.mm形式のいずれかとする。

xmind形式のファイルはライブラリ`xmind-sdk`を使用してノードのデータを取得する。

FreePlaneのmm形式のファイルはXMLファイルとしてノードのデータを取得する

出力するExcel、Wordのファイル形式はOffice2019以降の\.xlsx形式、\.docx形式とする。

### コーディング規約

pythonのpep8をベースとして、既存のスクリプトのコーディングスタイルを踏襲すること

既存のスクリプト同様にnumpyスタイルのdocstringを付与すること。

Sphinxを用いてドキュメント出力できることを前提とする

# function 機能仕様

## 概要

リファクタリング前のツールで実現できていることを機能仕様として示す。<br/>リファクタリング後も外部仕様は変わらないこと。<br/>`README.md`を参照のこと。

Windows上のコマンドプロンプト上で実行できること<br/>Pythonスクリプトの存在するフォルダ上での例： `python Xmind2USDM.py （オプション指定）`

コマンドプロンプト上で実行しなくても良いようにバッチファイルを用意し、バッチファイル上に処理対象のファイルをDrag&Dropするだけで変換可能になっている。<br/>\- `convert_Xmind2USDM.bat`<br/>\- `convert_USDM2doc.bat`<br/>\- `convert_USDM2Md.bat`<br/>\- `convert_Xmind2Md.bat`<br/><br/>これらのバッチファイルは実行環境に合わせて編集する必要がある。<br/>\- python\.exe のpath<br/>\- 実行スクリプトのpath<br/>\- テンプレートフォルダのpath

実行環境は`USDMcfg.py`で定義し、`Xmind2USDM.py`を変更しなくて良いようになっている。

USDMのテンプレート、Wordのテンプレートを`templates`フォルダに保存しておき変換時に利用する。

MarkDown形式では理由や備考について`templates`フォルダ内のスタイルシート`style.css`で装飾している。

## 利用方法

`README_Xmind2USDM.md`、`README_USDM2word.md`参照

# refact リファクタリング仕様

## `Xmind2USDM.py`と`USDM2Word.py`の統合

現状はMindmapからUSDMへの変換とUSDMからWordファイル、MarkDownへの変換が別ツールになっている。<br/>これを、1つのスクリプト`mm2docs.py`の中で出力ファイル指定の拡張子で判別し、1つのスクリプトで3種類の出力を得られるようにする。<br/>出力ファイルの指定が<br/>  `*.xlsx`であれば、Excel形式<br/>  `*.docx`であれば、Word形式<br/>  `*.md`であれば、MarkDown形式

## 処理フロー

下記の処理フローとする<br/>1\. 定義ファイル`config.py`からデータ取得や出力のための情報を取得する<br/>2\. Mindmapからノードデータの取得<br/>3\. 出力先の指定に応じた出力処理

### 定義ファイルの処理

定義ファイルのファイル名を`USDMcfg.py`から`config.py`に変更する。

Mindmapの読み取り用、USDM出力用、USDM読み取り用、Word出力用、MarkDown出力用のデータがフラットに定義されているが、これらを用途別にdataclass化してわかりやすく整理したい

dataclassの定義と各定義データのインスタンス化は別のファイルにする。<br/>データのインスタンス化（実データの設定）を`mmconfig.py`で、class定義を`mmclass.py`で行う

### Mindmapのノードデータの取得

`mm2docs.py`実行時に指定されたMindmapファイルからノードデータを取得する。<br/>`*.xmind`形式のファイルであれば、`xmind-sdk`を使ってノードデータを取得する。<br/>`*.mm`形式のファイルであれば、適切なXML用ライブラリを使ってノードデータを取得する。

ノードデータは現状listに保持しているが、`mmclass.py`内でdataclassを定義してそこに保持する

ノードデータの仕様レベルの判別に`if`、`elif`を用いているが、case文を使って可読性、変更の容易性を向上させたい

### 指定された形式でファイル出力

取得したノードデータを指定された出力形式でファイルに出力する

`*.xlsx`であれば、Excel形式（USDMフォーマット）<br/>テンプレートは指定された`*.xlsx`ファイルを使用する<br/>指定のテンプレートファイルの拡張子が`xlsx`で無ければエラーを表示して終了する。

不具合箇所として、<br/>USDM出力後、4行目（最初のデータ出力行）の行高さが自動設定されずに1行目だけしか表示されないことがある。<br/>他の行は正しく行の高さが自動設定されている。<br/>4行目も正しく高さ設定できるようにしたい。

また、`F4`の位置で`ウィンドウ枠の固定`をしたい。

`*.docx`であれば、Word形式<br/>テンプレートは指定された`*.docx`ファイルを使用する<br/>指定のテンプレートファイルの拡張子が`docx`で無ければエラーを表示して終了する。

`*.md`であれば、MarkDown形式<br/>Md形式の場合はテンプレートの指定は不要

Md形式のエスケープ文字列のうち、"`"はエスケープ対象外とし、さらに"`"で囲まれたテキストについては、エスケープ処理をしないものとする。

## テスト仕様

従来の`Xmind2USDM.py`で出力したファイルとリファクタリング後の`mm2docs.py`で出力したファイルの内容が同じであることを確認する。

差異が生じた場合は、コードを修正して差異が出ないようにする。

### USDM出力

python\.exe mm2docs\.py Xmind2USDM\_リファクタリング要件定義\.xmind \-t \.templates/USDM\_Template\.xlsx \-o Xmind2USDM\_リファクタリング要件定義\_USDM\_tesr\.xlsx

`Xmind2USDM_リファクタリング要件定義_USDM_test.xmind`と`Xmind2USDM_リファクタリング要件定義_USDM.xlsx`が一致すること

### Word出力

python\.exe mm2docs\.py Xmind2USDM\_リファクタリング要件定義\.xmind \-t \.templates/SpecTemplate\.docx \-o Xmind2USDM\_リファクタリング要件定義\_USDM\_r0\_test\.docx

`Xmind2USDM_リファクタリング要件定義_USDM_r0_test.docx`と`Xmind2USDM\_リファクタリング要件定義\_USDM\_r0\.docxが一致すること

### Markdown出力

python\.exe mm2docs\.py Xmind2USDM\_リファクタリング要件定義\.xmind \-t \.templates/SpecTemplate\.docx \-o Xmind2USDM\_リファクタリング要件定義\_USDM\_r0\_test\.docx

`Xmind2USDM_リファクタリング要件定義_USDM_r0_test.docx`と`Xmind2USDM\_リファクタリング要件定義\_USDM\_r0\.docxが一致すること

# docs ドキュメント生成

## README、CHANGELOG

リファクタリング後のスクリプトの処理を反映した`README.md`ファイルを生成する。<br/>内容は既存の`README.md`に準ずる。<br/>リファクタリング前の`README.md`は`README_old.md`にrenameして残す。

リファクタリング後のスクリプトを初版として`CHANGELOG.md`ファイルを生成する。<br/>過去の更新履歴は不要。<br/>リファクタリング前の`CHANGELOG.md`は`CHANGELOG_old.md`にrenameして残す。

## Sphinxによるドキュメント生成

sphinxを用いて`mm2docs.py`のドキュメントを生成する。<br/>出力先はサブフォルダ`docs`とする。

