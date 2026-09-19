# ノウハウ：日本語フォントの埋め込み

## なぜ必要か

編集できる .pptx / .docx で渡すと、相手の環境にフォントが無ければ別の書体で表示される。
Noto Sans JP は SIL Open Font License なので埋め込み可。

## 形式（アプリで違う）

| | PowerPoint | Word |
|---|---|---|
| 置き場所 | `ppt/fonts/fontN.fntdata` | `word/fonts/fontN.odttf` |
| 形式 | EOT（Embedded OpenType） | ODTTF（TTF の先頭32バイトを GUID で XOR） |
| 宣言 | `presentation.xml` の `<p:embeddedFontLst>` ＋ `embedTrueTypeFonts="1"` | `fontTable.xml` の `<w:embedRegular>` ＋ `settings.xml` の `<w:embedTrueTypeFonts/>` |

## 実装のポイント

- **静的フォントを使う**（可変フォント不可）。ウェイトは Regular と Bold の2つに絞るとファイルが軽い。
- PowerPoint の EOT は**非圧縮で通る**。先頭4バイト＝ファイル全体のサイズ、ヘッダーに PANOSE・ウェイト・
  Unicode 範囲・チェックサム・フォント名（UTF-16LE）が並ぶ。
- 「使用文字だけ埋め込む」は編集可能で渡す前提と矛盾するので使わない（`saveSubsetFonts="0"`）。

## 運用で起きること

| 現象 | 事実 |
|---|---|
| 自動操作（AppleScript）で保存すると埋め込みが消える | **人が画面から保存すれば残る**。自動保存で確認してはいけない |
| 人が保存するとファイルが約2倍になる | PowerPoint はフォントを ZIP 圧縮せずに格納する（中身は同一） |
| Word で保存すると「使用文字だけ」に削られる | 日本語フォントは Word が強制的にサブセット化する。**1回目の編集では全文字が使える**ので、提出直前に全文字を入れ直す運用にする |
| Google スライド／ドキュメント・Keynote | 埋め込みフォントは無視される |
