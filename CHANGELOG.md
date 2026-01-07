# CHANGELOG

## [1.1.0] - 2026-01-06

### Added

- メソッドに numpy 形式の docstring を追加。
- Markdown出力時、ノード内の改行を `<br />` に変換する機能を追加。
- Markdown出力時、バックトークン（` ` `）で囲まれたテキストをエスケープ対象外とする処理を追加。
- 新しいバッチファイル名体系（`convert_mm2*.bat`）への移行。

### Fixed

- `mm2docs.py` でのマッチ結果判定時の型エラーを修正。
- USDM Excel出力時、一部のレベルで列位置がズレていた問題を修正。

## [1.0.0] - 2026-01-06

### Added

- `mm2docs.py` を新規作成。
- Xmind2USDM.py と USDM2word.py を統合し、単一のスクリプトで Excel, Word, Markdown 出力を可能に。
- FreePlane (.mm) 形式の Mindmap 読み込みに対応。
- `mmclass.py`, `mmconfig.py` による設定情報の整理（dataclass化）。
- Python 3.10+ の `match/case` 文を用いたリファクタリング。
- Excel出力時のウィンドウ枠固定 (F4) と4行目の行高さ自動調整機能を追加。
- `README.md` および `CHANGELOG.md` を刷新。
