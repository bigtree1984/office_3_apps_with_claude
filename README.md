# Office 3 Apps with Claude

**PowerPoint・Word・Excel のテンプレートを、AI エージェント（Claude Code）に直接作らせる**ための手順書とスクリプトです。

Office ファイルは ZIP の中に XML が入っているだけなので、ライブラリが対応していない設定——
配色パレット、スライドマスターのガイド、フォント埋め込み、表スタイル、Word のスタイル体系——も、
XML を直接書けば全部設定できます。ここにあるのは、その進め方（PLAYBOOK）と、実際に動くスクリプト、
そして一式を作り切った完成見本です。

- 使っているライブラリ：`fonttools`（フォント埋め込み）、`pillow` だけ。`ezdxf` / `shapely` は `dxf_to_svg.py`（CAD から SVG を作る**任意**の道具）用で、ロゴが SVG なら要りません。
  **python-pptx / python-docx は使っていません**（テーマ・マスター・ガイド・埋め込みフォントを触れないため）。
- デザインの正（SSOT）は Figma。Figma から値とレイアウトを書き出して、スクリプトが POTX / DOTX を組み立てます。

## どこまで守るか（ここを先に読んでください）

**この一式は、全部が仕様ではありません。** 誰がやっても効く考え方と、作者の好みで決めただけのものが
混ざっています。見分けがつかないまま踏襲すると、**自分のブランドに合わないものを作り続けることになります。**

文書には印が付いています。**印のない記述は「既定」で、そのまま使ってよいものです。**

| 印 | 意味 |
|---|---|
| **【変えない】** | 守らないと壊れる／効果が消える。変えるなら理由を確かめてから |
| **【先に聞く】** | 作り始める前に決めること（明地か暗地か、スライドの大きさ、Word の表示、フォント、埋め込みの要否、資料を誰が組むか） |
| **【たたき台】** | 使ってよいが**作り替える前提**。Organism やページの装飾がこれ |
| （印なし） | 既定。要望が出たら変える |

詳しくは `PLAYBOOK.md` の冒頭にあります。**エージェントに使わせるときは、まずそこを読ませてください。**

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
| `bigtree/templates/` | `bigtree_lab.potx`（206KB）/ `bigtree_lab.dotx`（8KB）。**フォント埋め込みなし** |
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

# 検討用のプレビュー（Figma が無くても、ブラウザで見てフィードバックできる）
.venv/bin/python scripts/preview.py --open

# 自分のブランドで作る（値・ロゴ・レイアウトの書き出しを置いたフォルダを指定）
cp -r bigtree ~/mybrand      # 雛形だけ欲しいときは templates/tokens.template.json
OFFICE3_BRAND=~/mybrand .venv/bin/python scripts/build_potx_figma.py
```

- 出力先は `build/`（Git には入れません）。`OFFICE3_OUT` で変えられます。
### 公開ファイルの作り直し

```
.venv/bin/python scripts/publish_samples.py
```

`bigtree/templates/` と `bigtree/samples/` を作り直します。**フォントは埋め込みません**（再配布しないため、
かつ埋め込むと 1 ファイル 6.6MB＝32倍になり、作り直すたび履歴に積み上がるため）。
クライアント名などが混ざっていないかも中身を検査し、見つかればコピーせずに止まります。
自分のフォントを埋め込みたいときは、提出直前に `scripts/embed_fonts.py` を通してください。

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
