# デザインシステムの SSOT 地図（2026-09-19 決定）

> 原則（Bigtree Lab ブログから継承）：**SSOT は情報の性質で決める。置き場所の好みでは決めない。**

| 層 | 中身 | SSOT（正） | 見た目を確かめる場所 | 状態 |
|---|---|---|---|---|
| ① 共通トークン | 色、パレットの並び順、フォント、ウェイト、影 | **Figma**（変数） | Figma | 決定 |
| ② PowerPoint | レイアウト、ガイド、Organism、ボディゾーン Layout（9種） | **Figma** | Figma | 決定 |
| ③ グラフ | 系列の色順、強調の色と枠線、淡色の濃さ、軸の文字サイズ、目盛線 | **`CHART_RULES.md` ＋ `scripts/build_charts.py`**（色は①を参照） | Figma（見本帳：Office と HTML の出力を並べる） | 決定 |
| ④ Word | スタイル表（サイズ、行間、段落間隔、インデント、採番） | **md（スタイル表）** | Figma（見本帳：見本 DOCX を PDF→画像で並べる） | 決定 |

## ③ グラフの描き手は2つ

| 描き手 | 用途 | 実装 |
|---|---|---|
| A. Office ネイティブグラフ | PowerPoint / Word の運用報告 | `scripts/build_charts.py`（Python でグラフ XML を書く。描画は Office） |
| B. 画像のグラフ | ブログの挿絵など、編集させないもの | HTML テンプレート → PNG（この一式には入っていない。描画は Chrome） |

どちらも同じ作法（`CHART_RULES.md`）と同じトークン（`tokens.json`）を読む。1か所直せば両方変わる。

## 更新の向き

- Figma のトークンを変える → Claude が書き出す（tokens） → 3つのスクリプトが読む
- グラフの作法・Word のスタイル表を変える → ファイルを直す → Claude が見本を描き直して Figma に貼る
