# Office 3 Apps with Claude

**PowerPoint・Word・Excel のテンプレートを、AI エージェント（Claude Code）に直接作らせる**ための手順書とスクリプトです。

Office ファイルは ZIP の中に XML が入っているだけなので、ライブラリが対応していない設定——
配色パレット、スライドマスターのガイド、フォント埋め込み、表スタイル、Word のスタイル体系——も、
XML を直接書けば全部設定できます。ここにあるのは、その進め方（PLAYBOOK）と、実際に動くスクリプト、
そして一式を作り切った完成見本です。

- 使っているライブラリ：`fonttools`（フォント埋め込み）、`pillow`、`ezdxf` / `shapely`（ロゴ変換）だけ。
  **python-pptx / python-docx は使っていません**（テーマ・マスター・ガイド・埋め込みフォントを触れないため）。
- デザインの正（SSOT）は Figma。Figma から値とレイアウトを書き出して、スクリプトが POTX / DOTX を組み立てます。

## できること

| | 内容 |
|---|---|
| PowerPoint | テーマ（配色12枠・フォント・効果）、スライドマスター、レイアウト7種、ガイド、ページ番号、フォント埋め込み、表スタイル |
| Word | スタイル体系（表題・見出し1〜3・本文・キャプション・表）、見出しの自動採番、Web レイアウト表示、ヘッダー・フッター、フォント埋め込み |
| Excel / グラフ | ネイティブグラフ（グラフ XML＋埋め込みブック）。1つの定義から Excel・PowerPoint・Word の3か所へ。横に置いた Excel とのずれ検出つき |
| 図解（Organism） | Figma で作図した図解を、図形の種類ごとの汎用変換で PowerPoint の図形に |

## 完成見本

`bigtree/` に、note「だいきの試作室」（Bigtree Lab）向けに作った一式が入っています。そのまま開いて確認できます。

| 場所 | 中身 |
|---|---|
| `bigtree/templates/` | `bigtree_lab.potx` / `bigtree_lab.dotx`（フォント埋め込み済み） |
| `bigtree/samples/` | レイアウト見本、Organism、表、ネイティブグラフ入りの PPTX / DOCX |
| `bigtree/design/` | トークン（色・文字サイズ）、Figma からの書き出し、Word スタイル表、SSOT の地図 |
| `bigtree/assets/` | ロゴ（Fusion のスケッチ → SVG）、キービジュアル |

## 使い方

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt

# 完成見本をそのまま作り直す
.venv/bin/python scripts/build_potx_figma.py     # POTX とサンプル PPTX
.venv/bin/python scripts/build_dotx.py           # DOTX とサンプル DOCX
.venv/bin/python scripts/build_charts.py         # 横の Excel ＋ ネイティブグラフ
.venv/bin/python scripts/build_organisms.py      # 図解（Organism）のスライド
.venv/bin/python scripts/build_tables.py         # 表のスライド

# 自分のブランドで作る（値・ロゴ・レイアウトの書き出しを置いたフォルダを指定）
OFFICE3_BRAND=~/mybrand .venv/bin/python scripts/build_potx_figma.py
```

- 出力先は `build/`（Git には入れません）。`OFFICE3_OUT` で変えられます。
- フォントは同梱していません。**Noto Sans JP の静的フォント**（Regular / Bold）を Google Fonts から入手して
  `~/Library/Fonts` に置くか、`OFFICE3_FONT_DIR` で場所を指定してください（可変フォントは埋め込みに使えません）。

## はじめての人はここから

1. **[PLAYBOOK.md](./PLAYBOOK.md)** — 何をどの順に決め、何を作るかの手順書。**AI エージェントにこれを読ませて進めます。**
2. `notes/` — 実際にハマった罠と対処（文字位置・表の罫線・フォント埋め込み・OOXML 全般）
3. `bigtree/design/SSOT_MAP.md` — 「どの層の正をどこに置くか」の考え方

## 環境

macOS + Microsoft 365（PowerPoint / Word / Excel）で確認しています。Windows 版・Google スライド／ドキュメントでの
挙動は未確認です。Python は 3.11 で動作確認。

## ライセンス

MIT（[LICENSE](./LICENSE)）。`bigtree/` の中のロゴ・写真・文章は Bigtree Lab のものですが、**作り方を理解するための
見本として同じ MIT で公開**しています。自分のブランドで作り直す前提でご利用ください。
Noto Sans JP は SIL Open Font License です（同梱していません）。
