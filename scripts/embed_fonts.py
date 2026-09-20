"""提出直前に、フォントを「全文字ぶん」入れ直す後処理（PPTX / DOCX）。

Word は保存のたびに日本語フォントを「使用文字だけ」に削る。PowerPoint も、フォントを持っていない
相手に渡すなら全文字入っているほうが安全。人が編集し終えたファイルをこれに通してから提出する。

使い方:
  .venv/bin/python scripts/embed_fonts.py 提出用.pptx [出力先]
  .venv/bin/python scripts/embed_fonts.py 提出用.docx --check    # 状態を見るだけ

- 埋め込むフォントは design/tokens.json の font.family（Regular / Bold）。
- フォント本体は OFFICE3_FONT_DIR（既定 ~/Library/Fonts）から探す。
"""
import os
import re
import shutil
import struct
import sys
import uuid
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402  (make_eot, FONT_DIR, tokens)
import build_dotx as bd  # noqa: E402  (obfuscate)

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument"
FAMILY = bp.FONT
WEIGHTS = [("regular", "Regular"), ("bold", "Bold")]


def font_file(style):
    name = FAMILY.replace(" ", "") + f"-{style}.ttf"
    p = bp.FONT_DIR / name
    if not p.exists():
        raise SystemExit(f"フォントが見つかりません: {p}\n（Google Fonts の静的フォントを置くか OFFICE3_FONT_DIR を設定してください）")
    return p


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def write_zip(path, files):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k in ["[Content_Types].xml"] + [k for k in files if k != "[Content_Types].xml"]:
            z.writestr(k, files[k])


def add_default(ct, ext, ctype):
    if f'Extension="{ext}"' in ct:
        return ct
    return ct.replace("<Override", f'<Default Extension="{ext}" ContentType="{ctype}"/><Override', 1)


def rels_without_fonts(xml):
    """既存のフォント参照を消して、次に使える rId 番号を返す。"""
    xml = re.sub(r'<Relationship[^>]*/font"[^>]*/>', "", xml)
    xml = re.sub(r'<Relationship[^>]*Type="[^"]*/font"[^>]*/>', "", xml)
    used = [int(m) for m in re.findall(r'Id="rId(\d+)"', xml)]
    return xml, (max(used) + 1 if used else 1)


def check(files):
    """現状の埋め込み状態を表示する。"""
    fonts = [n for n in files if n.startswith(("ppt/fonts/", "word/fonts/"))]
    print(f"  埋め込みフォント: {len(fonts)} 個")
    for n in sorted(fonts):
        print(f"    {n}  {len(files[n]) / 1e6:.1f} MB")
    ft = files.get("word/fontTable.xml", b"").decode("utf-8", "ignore")
    for m in re.finditer(r'<w:font w:name="([^"]+)">(.*?)</w:font>', ft, re.S):
        subset = ' w:subsetted="1"' in m.group(2)
        if "embed" in m.group(2):
            print(f"    Word: {m.group(1)} … {'使用文字だけ（サブセット）' if subset else '全文字'}")


def embed_pptx(files):
    pres = files["ppt/presentation.xml"].decode()
    rels = files["ppt/_rels/presentation.xml.rels"].decode()
    for n in [k for k in files if k.startswith("ppt/fonts/")]:
        del files[n]
    rels, rid = rels_without_fonts(rels)
    entries = ""
    for i, (slot, style) in enumerate(WEIGHTS, 1):
        files[f"ppt/fonts/font{i}.fntdata"] = bp.make_eot(font_file(style))
        rels = rels.replace("</Relationships>", f'<Relationship Id="rId{rid}" Type="{REL}/font" Target="fonts/font{i}.fntdata"/></Relationships>')
        entries += f'<p:{slot} r:id="rId{rid}"/>'
        rid += 1
    lst = f'<p:embeddedFontLst><p:embeddedFont><p:font typeface="{FAMILY}" charset="-128"/>{entries}</p:embeddedFont></p:embeddedFontLst>'
    pres = re.sub(r"<p:embeddedFontLst>.*?</p:embeddedFontLst>", "", pres, flags=re.S)
    pres = pres.replace("<p:defaultTextStyle>", lst + "<p:defaultTextStyle>")
    if "embedTrueTypeFonts" not in pres:
        pres = pres.replace("<p:presentation ", "<p:presentation embedTrueTypeFonts=\"1\" ", 1)
    pres = pres.replace('saveSubsetFonts="1"', 'saveSubsetFonts="0"')
    files["ppt/presentation.xml"] = pres.encode()
    files["ppt/_rels/presentation.xml.rels"] = rels.encode()
    files["[Content_Types].xml"] = add_default(files["[Content_Types].xml"].decode(), "fntdata", "application/x-fontdata").encode()
    return len(WEIGHTS)


def embed_docx(files):
    ft = files.get("word/fontTable.xml", b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>').decode()
    for n in [k for k in files if k.startswith("word/fonts/")]:
        del files[n]
    refs, rels_items = "", []
    for i, (slot, style) in enumerate(WEIGHTS, 1):
        guid = "{" + str(uuid.uuid4()).upper() + "}"
        files[f"word/fonts/font{i}.odttf"] = bd.obfuscate(font_file(style).read_bytes(), guid)
        rels_items.append((f"rId{i}", f"fonts/font{i}.odttf"))
        refs += f'<w:embed{"Regular" if slot == "regular" else "Bold"} r:id="rId{i}" w:fontKey="{guid}"/>'
    # 他のフォントの埋め込み参照も消す（本体を削ったので、残すと参照切れになる）
    ft = re.sub(r'<w:embed\w+ [^>]*/>', "", ft)
    entry = f'<w:font w:name="{FAMILY}"><w:charset w:val="80"/><w:family w:val="swiss"/><w:pitch w:val="variable"/>{refs}</w:font>'
    ft = re.sub(rf'<w:font w:name="{re.escape(FAMILY)}">.*?</w:font>', "", ft, flags=re.S)
    ft = ft.replace("</w:fonts>", entry + "</w:fonts>") if "</w:fonts>" in ft else ft.replace("/>", f">{entry}</w:fonts>", 1)
    files["word/fontTable.xml"] = ft.encode()
    files["word/_rels/fontTable.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="{i}" Type="{REL}/font" Target="{t}"/>' for i, t in rels_items)
        + "</Relationships>").encode()
    st = files["word/settings.xml"].decode()
    if "<w:embedTrueTypeFonts/>" not in st:
        st = st.replace("<w:defaultTabStop", "<w:embedTrueTypeFonts/><w:defaultTabStop", 1)
    st = st.replace("<w:saveSubsetFonts/>", "")
    files["word/settings.xml"] = st.encode()
    ct = add_default(files["[Content_Types].xml"].decode(), "odttf", f"{CT}.obfuscatedFont")
    if "/word/fontTable.xml" not in ct:
        ct = ct.replace("</Types>", f'<Override PartName="/word/fontTable.xml" ContentType="{CT}.wordprocessingml.fontTable+xml"/></Types>')
    files["[Content_Types].xml"] = ct.encode()
    # fontTable の関係が無いときだけ足す（Target は "fontTable.xml"。重複させると Word が新規文書として開く）
    if "/fontTable" not in files.get("word/_rels/document.xml.rels", b"").decode():
        r = files["word/_rels/document.xml.rels"].decode()
        rid = f"rIdFontTable"
        r = r.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{REL}/fontTable" Target="fontTable.xml"/></Relationships>')
        files["word/_rels/document.xml.rels"] = r.encode()
    return len(WEIGHTS)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    src = Path(args[0])
    files = read_zip(src)
    print(f"{src.name}（{src.stat().st_size / 1e6:.1f} MB）")
    print("  ── 処理前 ──")
    check(files)
    if "--check" in sys.argv:
        return
    n = embed_pptx(files) if src.suffix in (".pptx", ".potx") else embed_docx(files)
    dst = Path(args[1]) if len(args) > 1 else src.with_name(src.stem + "_embedded" + src.suffix)
    write_zip(dst, files)
    print(f"  ── 処理後 ── {FAMILY} を {n} ウェイト、全文字で埋め込み")
    check(read_zip(dst))
    print(f"  → {dst}（{dst.stat().st_size / 1e6:.1f} MB）")


if __name__ == "__main__":
    main()
