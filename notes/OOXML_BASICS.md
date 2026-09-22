# ノウハウ：Office ファイルを XML で直接作るときの基本

> **`notes/` の内容は【変えない】です。** ここに書いてあるのは作者の好みではなく、
> **Office がそう作られているために起きること**と、その対処です。好みで外すと壊れます。

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

## 6. 同じ値を2か所に置かない

文字サイズを `tokens.json` と `figma_layouts.json` の両方に持っていたため、**トークンを変えてもグラフの文字が追随しなかった**（2026-09-20 に発見）。
グラフの文字サイズにいたっては、スクリプトに数値を直接書いていた（たまたま一致していたので気づけなかった）。

- 値の正は**必ず1か所**。他のファイルは「そこを見る」と書くだけにする。
- **確かめ方**：トークンを一時的に別の値に変えて生成し、出力が追随するかを見る。目視では気づけない。
- 同じ性質の事故：「指定しない＝Office の既定が出る」。**書き漏らしと二重持ちが、デザインシステムの二大事故**。

## chartEx ―― ウォーターフォールなど「新しいグラフ」は別形式

Excel 2016 以降のウォーターフォール・ツリーマップ・サンバースト・じょうごは、
従来のグラフ（`c:chartSpace`）ではなく **chartEx**（`cx:chartSpace`）という別の形式で保存される。

```
ppt/charts/chart7.xml     cx:chartSpace           application/vnd.ms-office.chartex+xml
ppt/charts/colors7.xml    cs:colorStyle           application/vnd.ms-office.chartcolorstyle+xml
ppt/charts/style7.xml     cs:chartStyle           application/vnd.ms-office.chartstyle+xml
```

スライド側の `a:graphicData` の uri は `http://schemas.microsoft.com/office/drawing/2014/chartex`、
グラフへの関係は `http://schemas.microsoft.com/office/2014/relationships/chartEx`。

ウォーターフォールの形は `<cx:series layoutId="waterfall">`。
**連結線（棒と棒を結ぶ横線）も、増加／減少／合計の色分けも、この形式の機能として入っている。**
合計として扱う柱は `<cx:layoutPr><cx:subtotals><cx:idx val="0"/></cx:subtotals></cx:layoutPr>` で指す。
値は増減、合計の行だけ到達点そのものを入れる。

### 詰まった2点（どちらも修復ダイアログになる）

1. **要素の順番**。`cx:series` の中は `tx → dataLabels → dataId → layoutPr`。
   `dataLabels` を後ろに置いただけで PowerPoint が「コンテンツに問題が見つかりました」を出す。
2. **色とスタイルを別部品で持つ**。`colors{n}.xml` / `style{n}.xml` が無いと開けない。
   しかも**ウォーターフォールの増加・減少・合計の色は、色部品の先頭3つ**から順に取られる
   （系列側で色を指定するのではない）。
   さらに `style{n}.xml` の `cs:dataPoint` を `<cs:fillRef idx="0"/>` にすると、
   **棒の塗りが消えて線だけになる**。`<cs:fillRef idx="1"><cs:styleClr val="auto"/></cs:fillRef>` が要る。

### 調べ方

仕様書を読むより、**その形式を実際に書き出しているソフトのソースを読むほうが速い**。
今回は R のライブラリ（encharter）の生成コードと突き合わせて、順番と必要な部品を確認した。
推測で直すと、外すたびに修復ダイアログを閉じてもらうことになる。

### Word に入れるときも同じ（見落としやすい）

chartEx は PowerPoint・Excel・Word で**それぞれ別の場所に貼り付けコードがある**。
PowerPoint だけ直して満足していると、Word 側が古い `c:chart` のまま残って壊れる（実際に壊した）。
貼り付け先は3か所：`inject_pptx` / `xlsx` / `inject_docx`。uri・関係の種類・コンテンツタイプの
3点セットを、どこでも同じように差し替える。

Word の PDF 書き出しは AppleScript の `save as document 1 file name … file format format PDF`。
ただし **Word のサンドボックスは /tmp に書けない**。パスを渡しても
`~/Library/Containers/com.microsoft.Word/Data/Documents/` に出る（PowerPoint は /tmp に書ける）。
