"""Build a test DOTX / sample DOCX from raw OOXML (no python-docx).

Capability test only: values (sizes, spacing) are placeholders.
Outputs (build/):
  bigtree_lab.dotx               template
  bigtree_lab_sample.docx        sample document, grid OFF (recommended)
  bigtree_lab_grid_on.docx       same content, Japanese line grid ON (B1 demo)
Usage: .venv/bin/python scripts/build_dotx.py
"""
import struct
import os
import sys
import uuid
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx  # noqa: E402  (theme is shared with PowerPoint)

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND", ROOT / "bigtree")).resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT", ROOT / "build")).resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
FONT_DIR = Path(os.environ.get("OFFICE3_FONT_DIR", Path.home() / "Library/Fonts"))
FIGURE = Path("/Users/shibanodaiki/0011_zenitha/tech_lab_blog/out/F-01_table.png")

XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
DRAW_NS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument"

# A4, twips (1pt = 20tw)
PAGE_W, PAGE_H, MARGIN = 11906, 16838, 1134
TEXT_W = PAGE_W - 2 * MARGIN
DOCX = build_potx._TOKENS["type_docx"]["roles"]
COLORS = build_potx.COLORS
BODY_SZ = int(DOCX["body"]["pt"] * 2)  # half-points
E_IND = 420   # 1 字ぶんの字下げ（twips, 本文 10.5pt）


def sz(role):
    return int(DOCX[role]["pt"] * 2)


def line(role, after=0, before=0):
    """Exact line spacing in twips (see reference/TEXT_LAYOUT_NOTES.md: Word's 'auto' multiple is
    relative to the font's own leading, like PowerPoint's percent)."""
    r = DOCX[role]
    tw = int(round(r["pt"] * r["line"] / 100 * 20))
    b = f' w:before="{before}"' if before else ""
    a = f' w:after="{after}"' if after else ' w:after="0"'
    return f'<w:spacing{b}{a} w:line="{tw}" w:lineRule="exact"/>'


def role_rpr(role, theme_color, major=False):
    r = DOCX[role]
    bold = "<w:b/><w:bCs/>" if r["weight"] == "Bold" else ""
    font = THEME_FONT_MAJOR if major else THEME_FONT_MINOR
    return f'{font}{bold}{color(theme_color, COLORS[{"text1": "dk1", "text2": "dk2", "accent1": "accent1", "accent3": "accent3", "accent6": "accent6", "hyperlink": "hlink"}[theme_color]])}<w:sz w:val="{sz(role)}"/><w:szCs w:val="{sz(role)}"/>'


# ---------------------------------------------------------------- theme
def theme():
    t = build_potx.theme()
    # Word picks the Japanese font from the script-specific entry when the UI/lang is ja-JP
    return t.replace('<a:cs typeface=""/>', '<a:cs typeface="Noto Sans JP"/><a:font script="Jpan" typeface="Noto Sans JP"/>')


# ---------------------------------------------------------------- styles
THEME_FONT_MINOR = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>'
THEME_FONT_MAJOR = '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>'


def color(theme_color, hexval):
    return f'<w:color w:val="{hexval}" w:themeColor="{theme_color}"/>'


def pstyle(sid, name, ppr="", rpr="", based="Normal", nxt="Normal", builtin=True, extra=""):
    b = f'<w:basedOn w:val="{based}"/>' if based else ""
    n = f'<w:next w:val="{nxt}"/>' if nxt else ""
    return (f'<w:style w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/>{b}{n}{extra}'
            f'<w:uiPriority w:val="1"/><w:qFormat/><w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')


def cstyle(sid, name, rpr):
    return (f'<w:style w:type="character" w:customStyle="1" w:styleId="{sid}"><w:name w:val="{name}"/>'
            f'<w:uiPriority w:val="2"/><w:qFormat/><w:rPr>{rpr}</w:rPr></w:style>')


def styles(grid_on):
    snap = "" if grid_on else '<w:snapToGrid w:val="0"/>'
    heading = lambda lvl, role, before, c: pstyle(
        f"Heading{lvl}", f"heading {lvl}",
        f'<w:keepNext/>{snap}<w:numPr><w:ilvl w:val="{lvl - 1}"/><w:numId w:val="1"/></w:numPr>'
        f'{line(role, after=120, before=before)}<w:outlineLvl w:val="{lvl - 1}"/>',
        role_rpr(role, c, major=True))
    items = [
        # Normal: the base of everything
        (f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>'
         f'<w:pPr>{snap}{line("body", after=160)}<w:jc w:val="both"/></w:pPr>'
         f'<w:rPr>{role_rpr("body", "text1")}</w:rPr></w:style>'),
        pstyle("Title", "Title", f'{snap}{line("title", after=120)}', role_rpr("title", "text1", major=True) + '<w:kern w:val="28"/>'),
        pstyle("Subtitle", "Subtitle", f'{snap}{line("subtitle", after=480)}', role_rpr("subtitle", "text2")),
        heading(1, "h1", 480, "accent1"),
        heading(2, "h2", 360, "text1"),
        heading(3, "h3", 240, "text1"),
        pstyle("ListBullet", "List Bullet", f'<w:numPr><w:numId w:val="2"/></w:numPr>{line("body", after=60)}'),
        pstyle("ListNumber", "List Number", f'<w:numPr><w:numId w:val="3"/></w:numPr>{line("body", after=60)}'),
        pstyle("Caption", "caption", f'{snap}{line("caption", after=240, before=60)}<w:jc w:val="center"/>', role_rpr("caption", "text2")),
        pstyle("Footer", "footer", f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs>{line("caption")}',
               role_rpr("caption", "text2"), nxt=None),
        # 図: line spacing must stay "auto" — an exact line height clips inline images to that height
        pstyle("TOCHeading", "TOC Heading", f'{snap}{line("h1", after=120, before=240)}', role_rpr("h1", "accent1", major=True), nxt="Normal"),
        pstyle("TOC1", "toc 1", f'{snap}{line("body", after=60)}<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/></w:tabs>'),
        pstyle("TOC2", "toc 2", f'{snap}{line("body", after=40)}<w:ind w:left="{E_IND}"/>'),
        pstyle("TOC3", "toc 3", f'{snap}{line("body", after=40)}<w:ind w:left="{E_IND * 2}"/>'),
        pstyle("FigureBlock", "図", f'{snap}<w:keepNext/><w:spacing w:before="120" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>', nxt="Caption"),
        pstyle("TableText", "表内テキスト", f'{snap}{line("table")}<w:jc w:val="left"/>', role_rpr("table", "text1"), nxt="TableText"),
        # character styles: meaning, not appearance
        cstyle("Emph", "強調", f'<w:b/><w:bCs/>{color("accent1", COLORS["accent1"])}'),
        cstyle("Caution", "注意", f'<w:b/><w:bCs/>{color("accent6", COLORS["accent6"])}'),
        cstyle("Figure", "数値", f'<w:b/><w:bCs/><w:sz w:val="{sz("h3")}"/><w:szCs w:val="{sz("h3")}"/>'),
        ('<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/>'
         f'<w:unhideWhenUsed/><w:rPr>{color("hyperlink", COLORS["hlink"])}<w:u w:val="single"/></w:rPr></w:style>'),
        ('<w:style w:type="character" w:styleId="PlaceholderText"><w:name w:val="Placeholder Text"/>'
         f'<w:semiHidden/><w:rPr>{color("text2", COLORS["dk2"])}</w:rPr></w:style>'),
        ('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>'
         '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'),
        ('<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/>'
         '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>'
         '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'),
        # report table: same rules as the PowerPoint table style (header band, banded rows, horizontal rules only)
        ('<w:style w:type="table" w:customStyle="1" w:styleId="ReportTable"><w:name w:val="レポート表"/><w:basedOn w:val="TableNormal"/>'
         '<w:uiPriority w:val="3"/><w:qFormat/>'
         f'<w:pPr>{line("table")}<w:jc w:val="left"/>{snap}</w:pPr>'
         f'<w:rPr><w:sz w:val="{sz("table")}"/><w:szCs w:val="{sz("table")}"/></w:rPr>'
         # borders not mentioned fall back to Word's defaults, so "none" must be explicit (same trap as PowerPoint)
         '<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>'
         '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
         '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
         f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{COLORS["accent1"]}" w:themeColor="accent1"/>'
         f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{COLORS["accent5"]}" w:themeColor="accent5"/>'
         '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>'
         '<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/>'
         '<w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr>'
         f'<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/>{color("background1", COLORS["lt1"])}</w:rPr>'
         f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{COLORS["accent1"]}" w:themeFill="accent1"/></w:tcPr></w:tblStylePr>'
         f'<w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{COLORS["lt2"]}" w:themeFill="background2"/></w:tcPr></w:tblStylePr>'
         '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>'
         '</w:style>'),
    ]
    defaults = (f'<w:docDefaults><w:rPrDefault><w:rPr>{THEME_FONT_MINOR}<w:kern w:val="2"/>'
                f'<w:sz w:val="{BODY_SZ}"/><w:szCs w:val="{BODY_SZ}"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
                '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults>')
    return XML + f'<w:styles {W_NS}>{defaults}{"".join(items)}</w:styles>'


# ---------------------------------------------------------------- numbering
def numbering():
    def lvl(i, fmt, text, left, hanging, pstyle=None, font=None):
        ps = f'<w:pStyle w:val="{pstyle}"/>' if pstyle else ""
        rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:hint="default"/></w:rPr>' if font else ""
        return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>{ps}<w:lvlText w:val="{text}"/>'
                f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr>{rpr}</w:lvl>')
    headings = "".join(lvl(i, "decimal", ".".join(f"%{k + 1}" for k in range(i + 1)) + ("." if i == 0 else ""),
                           [425, 567, 709][i], [425, 567, 709][i], f"Heading{i + 1}") for i in range(3))
    bullets = "".join(lvl(i, "bullet", "•" if i % 2 == 0 else "◦", 420 * (i + 1), 210, font="Noto Sans JP") for i in range(3))
    numbers = "".join(lvl(i, ["decimal", "decimalEnclosedCircle", "aiueoFullWidth"][i], f"%{i + 1}" + ("." if i == 0 else ""),
                          420 * (i + 1), 420) for i in range(3))
    return (XML + f'<w:numbering {W_NS}>'
            f'<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>{headings}</w:abstractNum>'
            f'<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>{bullets}</w:abstractNum>'
            f'<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>{numbers}</w:abstractNum>'
            '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
            '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
            '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num></w:numbering>')


# ---------------------------------------------------------------- settings
def settings():
    return (XML + f'<w:settings {W_NS}>'
            '<w:view w:val="web"/><w:zoom w:percent="100"/>'
            '<w:embedTrueTypeFonts/>'
            # <w:updateFields> は入れない：開くたびに「他のファイルを参照するフィールド…」という
            # 分かりにくいダイアログが出るため。目次の中身は Claude が生成時に書き込む。
            '<w:defaultTabStop w:val="420"/>'                       # 2 chars at 10.5pt
            '<w:characterSpacingControl w:val="compressPunctuation"/>'  # 約物の詰め
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>'
            '<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" '
            'w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" '
            'w:followedHyperlink="followedHyperlink"/>'
            '</w:settings>')


# ---------------------------------------------------------------- fonts (ODTTF)
EMBED = [("Noto Sans JP", [("embedRegular", "NotoSansJP-Regular.ttf"), ("embedBold", "NotoSansJP-Bold.ttf")]),
]


def obfuscate(ttf_bytes, guid):
    """ECMA-376 Part 2 font obfuscation: XOR the first 32 bytes with the reversed GUID bytes."""
    key = bytes.fromhex(guid.strip("{}").replace("-", ""))[::-1]
    data = bytearray(ttf_bytes)
    for i in range(32):
        data[i] ^= key[i % 16]
    return bytes(data)


def font_table(embed):
    parts, rels, n = [], [], 0
    for face, slots in EMBED:
        refs = ""
        if embed:
            for slot, ttf in slots:
                n += 1
                guid = "{" + str(uuid.uuid4()).upper() + "}"
                parts.append((f"word/fonts/font{n}.odttf", obfuscate((FONT_DIR / ttf).read_bytes(), guid)))
                rels.append(("font", f"fonts/font{n}.odttf"))
                refs += f'<w:{slot} r:id="rId{n}" w:fontKey="{guid}"/>'
        parts_xml = (f'<w:font w:name="{face}"><w:charset w:val="80"/><w:family w:val="swiss"/>'
                     f'<w:pitch w:val="variable"/>{refs}</w:font>')
        parts.append(("__xml__", parts_xml))
    xml = XML + f'<w:fonts {W_NS}>' + "".join(v for k, v in parts if k == "__xml__") + '</w:fonts>'
    return xml, [(k, v) for k, v in parts if k != "__xml__"], rels


# ---------------------------------------------------------------- document body helpers
def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def run(text, cstyle=None):
    rpr = f'<w:rPr><w:rStyle w:val="{cstyle}"/></w:rPr>' if cstyle else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def para(content):
    if isinstance(content, str) and not content.startswith("<w:r"):
        content = run(content)
    return f'<w:p>{content}</w:p>'


def p_style(style, text, ilvl=None):
    lv = f'<w:numPr><w:ilvl w:val="{ilvl}"/></w:numPr>' if ilvl else ""
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/>{lv}</w:pPr>{text if text.startswith("<w:r") else run(text)}</w:p>'


def field(instr, placeholder):
    return (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{placeholder}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')


def sdt(alias, tag, text, style=None, placeholder=True):
    """入力枠（コンテンツコントロール）。クリックすると案内文が消えて入力できる。"""
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    ph = '<w:showingPlcHdr/>' if placeholder else ""
    rpr = '<w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' if placeholder else ""
    return (f'<w:sdt><w:sdtPr><w:alias w:val="{alias}"/><w:tag w:val="{tag}"/><w:id w:val="{abs(hash(tag)) % 90000000}"/>'
            f'{ph}<w:text/></w:sdtPr><w:sdtContent><w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p></w:sdtContent></w:sdt>')


def toc(levels="1-2", page_numbers=False):
    """目次。見出しスタイルから自動生成される。既定は見出し1〜2まで（3まで拾うと長くなる）。
    Web レイアウト前提なのでページ番号は出さない。クリックで本文へ飛べる（\\h）。"""
    instr = f' TOC \\o "{levels}" \\h \\z \\u' + ("" if page_numbers else " \\n")
    return (p_style("TOCHeading", "目次")
            + '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
            + field(instr, "［ここに目次が入ります。右クリック →「フィールド更新」］")
            + '</w:p>')


def table(rows):
    ncol = len(rows[0])
    colw = TEXT_W // ncol
    grid = "".join(f'<w:gridCol w:w="{colw}"/>' for _ in range(ncol))
    trs = ""
    for ri, row in enumerate(rows):
        trpr = '<w:trPr><w:tblHeader/></w:trPr>' if ri == 0 else ""
        tcs = "".join(f'<w:tc><w:tcPr><w:tcW w:w="{colw}" w:type="dxa"/></w:tcPr>'
                      f'<w:p><w:pPr><w:pStyle w:val="TableText"/><w:jc w:val="{"right" if ci and ri else "left"}"/></w:pPr>{run(c)}</w:p></w:tc>'
                      for ci, c in enumerate(row))
        trs += f'<w:tr>{trpr}{tcs}</w:tr>'
    return (f'<w:tbl><w:tblPr><w:tblStyle w:val="ReportTable"/><w:tblW w:w="{TEXT_W}" w:type="dxa"/>'
            f'<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
            f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{trs}</w:tbl>')


def image(rid, cx, cy, name):
    return (f'<w:p><w:pPr><w:pStyle w:val="FigureBlock"/></w:pPr><w:r><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="{name}"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')


def sect_pr(grid_on):
    grid = '<w:docGrid w:type="lines" w:linePitch="360"/>' if grid_on else '<w:docGrid w:linePitch="360"/>'
    return (f'<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/>'
            f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>'
            f'<w:pgMar w:top="{MARGIN}" w:right="{MARGIN}" w:bottom="{MARGIN}" w:left="{MARGIN}" w:header="567" w:footer="567" w:gutter="0"/>'
            f'<w:cols w:space="425"/><w:titlePg/>{grid}</w:sectPr>')


def footer():
    return (XML + f'<w:ftr {W_NS}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>'
            + run("© 2026 Bigtree Lab") + '<w:r><w:tab/></w:r>' + field("PAGE", "1") + run(" / ") + field("NUMPAGES", "1")
            + '</w:p></w:ftr>')


def cover():
    """表紙。項目は PowerPoint の表紙と同じ（区分・宛先・タイトル・サブタイトル・日付/作成者）。"""
    return (sdt("資料の区分", "classification", "公開用 ｜ ドラフト", "Caption")
            + sdt("宛先", "recipient", "〇〇〇〇 御中", "Normal")
            + sdt("タイトル", "title", "文書タイトル", "Title")
            + sdt("サブタイトル", "subtitle", "サブタイトル", "Subtitle")
            + sdt("日付・作成者", "meta", "2026-00-00 ｜ 所属 ｜ 作成者", "Caption"))


def body_template():
    return (cover() + toc() + p_style("Heading1", "見出し1") + para("本文をここに書きます。"))


def body_sample(img_rid, img_cx, img_cy):
    lorem = ("Web 解析と EC の運用データは、BigQuery から毎週自動で取得している。"
             "本レポートでは、直近 4 週間のセッション数・CVR・広告費用対効果（ROAS）の推移をまとめ、"
             "次の打ち手を提案する。数値はすべて税抜・速報値である。")
    return "".join([
        sdt("資料の区分", "classification", "公開用 ｜ ドラフト", "Caption", placeholder=False),
        sdt("宛先", "recipient", "だいきの試作室　読者のみなさまへ", "Normal", placeholder=False),
        sdt("タイトル", "title", "note トラフィックレポート", "Title", placeholder=False),
        sdt("サブタイトル", "subtitle", "2026年9月の読まれ方と次の一手", "Subtitle", placeholder=False),
        sdt("日付・作成者", "meta", "2026-09-20 ｜ Bigtree Lab ｜ だいき君", "Caption", placeholder=False),
        toc(),
        p_style("Heading1", "背景"),
        para(lorem),
        para(run("この段落には文字スタイルを当てている。") + run("強調したい語句", "Emph") + run("、")
             + run("注意喚起", "Caution") + run("、そして KPI の ") + run("ROAS 412%", "Figure")
             + run(" のような数値。リンクは ")
             + '<w:hyperlink r:id="rIdLink" w:history="1">' + run("note「だいきの試作室」", "Hyperlink") + '</w:hyperlink>'
             + run(" のように組み込みのスタイルを使う。")),
        p_style("Heading2", "箇条書き"),
        p_style("ListBullet", "箇条書きの第1レベル"),
        p_style("ListBullet", "Tab を押すとここ（第2レベル）に下がる", ilvl=1),
        p_style("ListBullet", "第3レベル", ilvl=2),
        p_style("ListBullet", "第1レベルに戻る"),
        p_style("Heading2", "番号付きリスト"),
        p_style("ListNumber", "手順その1"),
        p_style("ListNumber", "手順その2"),
        p_style("ListNumber", "手順その3"),
        p_style("Heading1", "検証結果"),
        p_style("Heading2", "週次の主要指標"),
        para("見出し番号（1. / 1.1）はすべて自動採番で、手では打っていない。"),
        table([["週", "セッション", "CVR", "ROAS"],
               ["W35", "12,480", "2.1%", "388%"],
               ["W36", "13,102", "2.3%", "401%"],
               ["W37", "12,955", "2.4%", "412%"],
               ["W38", "14,230", "2.2%", "395%"]]),
        p_style("Caption", "表 " + "") .replace("</w:p>", field("SEQ 表 \\* ARABIC", "1") + run("：週次の主要指標（ダミー）") + "</w:p>"),
        p_style("Heading2", "挿絵"),
        para("挿絵は行内配置にしている。文章を加筆しても、画像は段落と一緒に流れる。"),
        image(img_rid, img_cx, img_cy, "Figure 1"),
        p_style("Caption", "図 ").replace("</w:p>", field("SEQ 図 \\* ARABIC", "1") + run("：note 用に作った挿絵を Word に流用") + "</w:p>"),
        p_style("Heading3", "行間の確認用（英数字混じり）"),
        *[para(lorem) for _ in range(6)],
        p_style("Heading1", "まとめ"),
        para("フッターの「ページ / 総ページ」は自動で更新される。Web レイアウトでは表示されず、印刷・PDF 化したときに効く。"),
    ])


# ---------------------------------------------------------------- package
def build(path, kind, grid_on=False, embed=True):
    files = {}
    rels = [("styles", "styles.xml"), ("settings", "settings.xml"), ("numbering", "numbering.xml"),
            ("fontTable", "fontTable.xml"), ("theme", "theme/theme1.xml"), ("webSettings", "webSettings.xml")]
    extra_rels = [f'<Relationship Id="rIdFooter" Type="{REL}/footer" Target="footer1.xml"/>']
    overrides = []
    if kind == "template":
        body = body_template()
    else:
        from PIL import Image
        w, h = Image.open(FIGURE).size
        cx = int(TEXT_W * 635 * 0.8)  # 80% of text width, twips→EMU
        cy = int(cx * h / w)
        files["word/media/image1.png"] = FIGURE.read_bytes()
        extra_rels.append(f'<Relationship Id="rIdImg1" Type="{REL}/image" Target="media/image1.png"/>')
        extra_rels.append(f'<Relationship Id="rIdLink" Type="{REL}/hyperlink" Target="https://note.com/bigtree_lab" TargetMode="External"/>')
        body = body_sample("rIdImg1", cx, cy)

    main = "wordprocessingml.template.main+xml" if kind == "template" else "wordprocessingml.document.main+xml"
    files["word/document.xml"] = (XML + f'<w:document {W_NS} {DRAW_NS}><w:body>{body}{sect_pr(grid_on)}</w:body></w:document>')
    files["word/styles.xml"] = styles(grid_on)
    files["word/numbering.xml"] = numbering()
    files["word/settings.xml"] = settings()
    files["word/webSettings.xml"] = XML + f'<w:webSettings {W_NS}/>'
    files["word/theme/theme1.xml"] = theme()
    files["word/footer1.xml"] = footer()
    ft_xml, font_parts, font_rels = font_table(embed)
    files["word/fontTable.xml"] = ft_xml
    for k, v in font_parts:
        files[k] = v
    relxml = lambda items, extra="": (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                      + "".join(f'<Relationship Id="rId{i + 1}" Type="{REL}/{t}" Target="{tg}"/>' for i, (t, tg) in enumerate(items))
                                      + extra + '</Relationships>')
    files["word/_rels/document.xml.rels"] = relxml(rels, "".join(extra_rels))
    if font_rels:
        files["word/_rels/fontTable.xml.rels"] = relxml(font_rels)
    files["docProps/core.xml"] = (XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                                  'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>DOTX capability test</dc:title>'
                                  '<dc:creator>Claude Code</dc:creator></cp:coreProperties>')
    files["docProps/app.xml"] = (XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                                 '<Application>Claude Code</Application></Properties>')
    files["_rels/.rels"] = (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                            f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>'
                            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                            f'<Relationship Id="rId3" Type="{REL}/extended-properties" Target="docProps/app.xml"/></Relationships>')
    wml = f"{CT}.wordprocessingml"
    overrides += [("/word/document.xml", f"{CT}.{main}"), ("/word/styles.xml", f"{wml}.styles+xml"),
                  ("/word/numbering.xml", f"{wml}.numbering+xml"), ("/word/settings.xml", f"{wml}.settings+xml"),
                  ("/word/webSettings.xml", f"{wml}.webSettings+xml"), ("/word/fontTable.xml", f"{wml}.fontTable+xml"),
                  ("/word/footer1.xml", f"{wml}.footer+xml"), ("/word/theme/theme1.xml", f"{CT}.theme+xml"),
                  ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
                  ("/docProps/app.xml", f"{CT}.extended-properties+xml")]
    files["[Content_Types].xml"] = (
        XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        f'<Default Extension="odttf" ContentType="{CT}.obfuscatedFont"/>'
        + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides) + '</Types>')

    order = ["[Content_Types].xml", "_rels/.rels"] + [k for k in files if k not in ("[Content_Types].xml", "_rels/.rels")]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k in order:
            v = files[k]
            z.writestr(k, v if isinstance(v, bytes) else v.encode("utf-8"))
    print(f"wrote {path.relative_to(OUT.parent)} ({path.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    # --no-embed で公開用（フォントを埋め込まない）。scripts/publish_samples.py から使う
    build(OUT / "bigtree_lab.dotx", "template", embed="--no-embed" not in sys.argv)
    build(OUT / "bigtree_lab_sample.docx", "sample", grid_on=False, embed=False)  # samples stay light
    build(OUT / "bigtree_lab_grid_on.docx", "sample", grid_on=True, embed=False)
