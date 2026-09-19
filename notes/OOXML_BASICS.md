# ノウハウ：Office ファイルを XML で直接作るときの基本

## 1. Office ファイル ＝ ZIP ＋ XML

拡張子を `.zip` にして展開すれば中身が見える。主な置き場所：

| 設定 | PowerPoint | Word |
|---|---|---|
| 配色・フォント・効果（テーマ） | `ppt/theme/theme1.xml` | `word/theme/theme1.xml`（**同じ形式。使い回せる**） |
| レイアウト | `ppt/slideMasters/`・`ppt/slideLayouts/` | — |
| スタイル | マスターの `<p:txStyles>` | `word/styles.xml` |
| ガイド | マスター／レイアウトの `<p15:sldGuideLst>`（単位は 1/8pt） | — |
| ページ設定 | `ppt/presentation.xml` の `<p:sldSz>` | `document.xml` の `<w:sectPr>` |
| 表スタイル | `ppt/tableStyles.xml` | `styles.xml` 内の table スタイル |
| フォント埋め込み | `ppt/fonts/*.fntdata`（**EOT 形式**） | `word/fonts/*.odttf`（**GUID で XOR 難読化**） |
| グラフ | `ppt/charts/chart*.xml` ＋ `ppt/embeddings/*.xlsx` | `word/charts/…`（同じ） |

各パーツは `[Content_Types].xml` に型を、`_rels/*.rels` に参照関係を書く。ここを書き忘れると
「コンテンツに問題が見つかりました」になる。

## 2. 単位

| 単位 | 意味 |
|---|---|
| EMU | 1pt = 12,700 EMU / 1inch = 914,400 EMU。図形の位置・大きさ |
| 1/100 pt | 文字サイズ（`sz="1400"` = 14pt）、行間（`spcPts`） |
| 1/1000 % | 不透明度（`alpha val="12000"` = 12%）、比率 |
| twip | Word の距離（1pt = 20 twip） |
| 1/8 pt | PowerPoint のガイド位置 |

Figma のフレームを 1440×810px にして 10 インチのスライドに対応させると **2px = 1pt** で暗算できる。

## 3. 色は必ずテーマ参照にする

`<a:srgbClr val="2F6B55"/>` と直接書くと、配色を変えても追随しない。
`<a:schemeClr val="accent1"/>`（Word は `w:themeColor="accent1"`）を使う。

**色の割り当て（clrMap）は標準のままにする。** `bg2` と `tx2` を入れ替えたテンプレートを見たことがあるが、
AI も人も読み違える。

## 4. フォント埋め込み

- PowerPoint：**EOT**。ヘッダー＋フォント本体。PowerPoint 自身は MTX 圧縮をかけるが、
  **圧縮なしでも受け入れられる**（実測）。ヘッダーは fontTools で OS/2・head・name から組み立てられる。
- Word：**ODTTF**。GUID の 16 バイトを逆順にして、先頭 32 バイトと XOR する。
- どちらも**静的フォント**（Regular / Bold など太さごとのファイル）を使う。可変フォントは不可。
- Black などの特太は「別ファミリ」として扱う必要がある。ファイルが重くなるので、Regular と Bold だけに絞るのが現実的。

## 5. 作ったファイルの確かめ方

1. XML として妥当か（`xml.dom.minidom` で読めるか）——ここまでは自動で確認できる
2. **Office が受け入れるか**——開くまで分からない。1 が通っても 2 で落ちることがある
3. 往復テスト——人が保存し直した後も設定が残るか

「XML の妥当性 ≠ アプリの受理」。最終判定はアプリに開かせるしかない。
