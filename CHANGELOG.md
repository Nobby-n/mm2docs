# CHANGELOG

# [1.1.0] - 2026-03-30

### Added

- Mindmap内の画像をドキュメントに転記する機能を追加。XMind (.xmind) および FreePlane (.mm) の両形式に対応。
  - **USDM Excel**: 画像ごとに専用シートを作成し、タイトルと画像を配置。
  - **Word**: 見出し直後または該当ノード位置に画像とキャプションを挿入。
  - **Markdown**: `<img>` タグで画像を挿入（`width` 属性を `mmconfig.py` で変更可能）。
- `USDM2docs.py` に画像対応を追加。USDM Excelから Word/Markdown 変換時、`mm_images/` フォルダの画像をノードに自動関連付け。フォルダが無い場合は Excel の画像シートから自動抽出するフォールバック機能付き。
- `USDM2mm.py` に画像対応を追加。USDM Excelから FreePlane (.mm) 変換時、画像ノードを `richcontent` として出力し、`images/` フォルダに画像ファイルをコピー。
- `mmclass.py` の `MMNode` に `image_path` フィールドを追加。
- `mmconfig.py` に画像関連の設定を追加（`IMAGE_DIR`, `IMAGE_FILENAME_PROHIBIT`, `IMAGE_FILENAME_SUB`）。
- `mmconfig.py` の `USDMConfig` に画像シートの設定を追加（`img_title_cell`, `img_insert_cell`, `img_title_font`）。
- `mmconfig.py` の `MdConfig` に Markdown 画像の幅設定を追加（`img_width`、デフォルト `auto`）。

### Changed

- `mm2docs.py` の `_parse_raw_nodes` を拡張し、LV1 以外の全ノードレベル（LV2-LV4, LVREASON, LVREMARK, LVSPEC）で画像を関連付けられるように変更。
- `USDM2docs.py` の `load_mm_data` にて、LV1 ノードのテキストにシート名由来のキーワードを含めるよう変更（画像マッチングに必要）。

### Fixed

- `mm2docs.py` にて、備考（`//`）テキストをUSDM Excel に転記する際、G列（備考列）の背景色が要求レベルの薄緑色になる不具合を修正。備考のスタイル指定に `TXTREMARK`（`cell_style` に存在しないキー）を使用していたため書式が適用されず、テンプレートのデフォルトスタイルがそのまま残っていた。正しいキー `TXTREASON` に修正し、元の `Xmind2USDM.py` と同じ動作に復元。
- 同箇所にて、備考の書き込み行位置が1行ずれていた問題を修正（`rownum` のデクリメントをセル書き込みの前に移動）。

## [1.0.0] - 2026-02-09

### Added

- USDM形式のExcelからWord/Markdownを生成する `USDM2docs.py` を追加。
- USDM形式のExcelからMindmap(.mm)を生成（逆変換）する `USDM2mm.py` を追加。
- 上記スクリプトに対応するバッチファイル (`convert_USDM2doc.bat`, `convert_USDM2Md.bat`, `convert_USDM2mm.bat`) を追加。

### Changed

- `README.md` を大幅に更新し、ツール群全体の構成と各スクリプトの利用方法を明確化。
- ドキュメントの構成を見直し、説明の分かりやすさを向上 (`Concept_RequirementDefinition.md`, `HowToUse_mm2docs.md`)。

## [0.9.2] - 2026-01-08

### Added

- `README.md` に検証環境（OS, Pythonバージョン）の記載を追加。

### Changed

- 設定値の管理方法を変更。`mmclass.py` はデータ構造の定義（dataclass）のみとし、すべてのデフォルト値を `mmconfig.py` で管理するように集約。

## [0.9.1] - 2026-01-06

### Added

- メソッドに numpy 形式の docstring を追加。
- Markdown出力時、ノード内の改行を `<br />` に変換する機能を追加。
- Markdown出力時、バックトークン（` ` `）で囲まれたテキストをエスケープ対象外とする処理を追加。
- 新しいバッチファイル名体系（`convert_mm2*.bat`）への移行。

### Fixed

- `mm2docs.py` でのマッチ結果判定時の型エラーを修正。
- USDM Excel出力時、一部のレベルで列位置がズレていた問題を修正。

## [0.9.0] - 2026-01-06

### Added

- `mm2docs.py` を新規作成。
- Xmind2USDM.py と USDM2word.py を統合し、単一のスクリプトで Excel, Word, Markdown 出力を可能に。
- FreePlane (.mm) 形式の Mindmap 読み込みに対応。
- `mmclass.py`, `mmconfig.py` による設定情報の整理（dataclass化）。
- Python 3.10+ の `match/case` 文を用いたリファクタリング。
- Excel出力時のウィンドウ枠固定 (F4) と4行目の行高さ自動調整機能を追加。
- `README.md` および `CHANGELOG.md` を刷新。
