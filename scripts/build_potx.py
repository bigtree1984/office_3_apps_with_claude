"""Build a test POTX / sample PPTX from raw OOXML (no python-pptx).

Capability test only: values (colors, sizes) are placeholders.
Usage: .venv/bin/python scripts/build_potx.py [--embed]
"""
import os
import struct
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND") or ROOT / "bigtree").resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT") or ROOT / "build").resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
FONT_DIR = Path(os.environ.get("OFFICE3_FONT_DIR") or Path.home() / "Library/Fonts")

W, H = 9144000, 5143500            # 16:9, 10in x 5.625in (Google Slides size; placeholder, overwritten below from tokens)
EMU_PER_GUIDE = 1587.5            # guide pos unit = 1/8 pt
M = 457200                        # outer margin 0.5in

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
P15 = 'xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"'
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument"

# ---------------------------------------------------------------- theme
# shared tokens: exported from Figma (design/tokens.json is the scripts' copy; Figma is the SSOT)
import json as _json
_TOKENS = _json.loads((BRAND / "design/tokens.json").read_text())
COLORS = _TOKENS["colors"]
FONT = _TOKENS["font"]["family"]
# 見出しだけ別の書体にしたいとき（例：見出しは明朝、本文はゴシック）は font.heading に書く。無ければ本文と同じ。
# テーマの majorFont（見出し）／minorFont（本文）に分けて入れるので、人が PowerPoint で「見出しのフォント」を
# 選んでも同じ書体になる。どの役割が見出し書体かは type_pptx.roles の "font": "heading" で決める
FONT_HEADING = _TOKENS["font"].get("heading") or FONT
# ブランドに関する文字列は brand ブロックが正。theme_name は、テーマ名だけ別にしたいとき用の任意の上書き
BRAND_INFO = {**{"slug": "template", "url": "https://example.com", "logo": "assets/logo/logo.svg",
                 "meta": "0000-00-00 ｜ 所属 ｜ 作成者"},
              **_TOKENS.get("brand", {})}
BRAND_INFO.setdefault("name", _TOKENS.get("theme_name", "Template"))
BRAND_INFO.setdefault("copyright", f'© {BRAND_INFO["name"]}')
THEME_NAME = _TOKENS.get("theme_name") or BRAND_INFO["name"]
SLUG = BRAND_INFO["slug"]
COPYRIGHT = BRAND_INFO["copyright"]
W, H = _TOKENS["slide"]["width_emu"], _TOKENS["slide"]["height_emu"]


def rgb(color):
    """トークン名でも 6桁の hex でも受け取って (r, g, b) にする。"""
    h = COLORS.get(color, color if isinstance(color, str) and len(color) == 6 else "888888")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def blend(fg, bg, alpha=1.0):
    """半透明の塗りを下の色と混ぜる（薄い塗りを不透明として測ると誤判定になる）。"""
    a, b = rgb(fg), rgb(bg)
    return "".join(f"{int(round((a[i] * alpha + b[i] * (1 - alpha)) * 255)):02x}" for i in range(3))


def contrast(fg, bg):
    """2色のコントラスト比（WCAG）。"""
    def lum(color):
        r, g, b = rgb(color)
        f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)
    a, b = sorted((lum(fg), lum(bg)))
    return (b + 0.05) / (a + 0.05)


def text_on(bg):
    """その塗りの上に置く文字の色（"lt1" か "dk1"）。**実際のコントラスト比を比べて、高いほうを選ぶ。**

    **色を差し替えた人を守るための判定。** 図解やグラフに「白文字」と書き込んでおくと、
    明るいブランド色に変えた瞬間に読めなくなる（濃い色を前提にした値が残るため）。
    明るさのしきい値で決める方法も試したが、中間色（例：#598977）で逆の答えを出したので、
    **両方の比を計算して比べる**ことにした。
    """
    return "lt1" if contrast("lt1", bg) >= contrast("dk1", bg) else "dk1"


def role_family(role):
    """文字の役割（名前か roles の dict）が使う書体。"font": "heading" なら見出し書体、それ以外は本文書体。"""
    r = role if isinstance(role, dict) else _TOKENS["type_pptx"]["roles"][role]
    return FONT_HEADING if r.get("font") == "heading" else FONT


def font_path(style="Regular", family=None):
    """フォントの置き場所を返すだけ（存在するとは限らない）。"""
    return FONT_DIR / f'{(family or FONT).replace(" ", "")}-{style}.ttf'


def font_file(style="Regular", family=None):
    """フォントの実ファイル。無ければ、何をすればいいかを書いて止まる（生のトレースを出さない）。"""
    family = family or FONT
    path = font_path(style, family)
    if not path.exists():
        raise SystemExit(
            f"\nフォントが見つかりません: {path}\n"
            f"  {family} の「静的フォント」（Regular / Bold）を用意してください。\n"
            "  ・Google Fonts からダウンロード → static フォルダの .ttf を使う（可変フォントは埋め込みに使えません）\n"
            f"  ・置き場所は ~/Library/Fonts、または OFFICE3_FONT_DIR で指定\n"
            "  ・**フォントを用意せずに一通り動かしたいときは --no-embed** を付けてください\n"
            "    （テンプレートは作れます。文字のはみ出し検査だけ飛ばします）\n")
    return path


def shadow(dist, blur, alpha):
    return (f'<a:effectStyle><a:effectLst><a:outerShdw blurRad="{blur}" dist="{dist}" dir="2700000" '
            f'algn="tl" rotWithShape="0"><a:srgbClr val="{COLORS["dk1"]}"><a:alpha val="{alpha}"/></a:srgbClr>'
            f'</a:outerShdw></a:effectLst></a:effectStyle>')


def theme():
    clr = "".join(f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>' for k, v in COLORS.items())
    fonts = "".join(f'<a:{k}><a:latin typeface="{f}"/><a:ea typeface="{f}"/><a:cs typeface=""/></a:{k}>'
                    for k, f in (("majorFont", FONT_HEADING), ("minorFont", FONT)))
    solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    ln = "".join(f'<a:ln w="{w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
                 for w in (6350, 12700, 19050))
    return (XML + f'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="{THEME_NAME}">'
            f'<a:themeElements><a:clrScheme name="{THEME_NAME}">{clr}</a:clrScheme>'
            f'<a:fontScheme name="{FONT}">{fonts}</a:fontScheme>'
            f'<a:fmtScheme name="{THEME_NAME}"><a:fillStyleLst>{solid * 3}</a:fillStyleLst>'
            f'<a:lnStyleLst>{ln}</a:lnStyleLst>'
            f'<a:effectStyleLst>{shadow(50800, 152400, 18000) * 3}</a:effectStyleLst>'  # 3段とも同じ影。人が「標準スタイル」を選び分けても崩れないため
            f'<a:bgFillStyleLst>{solid * 3}</a:bgFillStyleLst></a:fmtScheme>'
            f'</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')


# ---------------------------------------------------------------- shapes
_id = [1]


def nid():
    _id[0] += 1
    return _id[0]


def xfrm(x, y, cx, cy):
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'


def ph(name, ph_attr, box=None, prompt=None, size=None, color=None, bold=None, anchor=None):
    sppr = f'<p:spPr>{xfrm(*box)}</p:spPr>' if box else '<p:spPr/>'
    rpr_attr = (f' sz="{size}"' if size else "") + (f' b="{bold}"' if bold is not None else "")
    fill = f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>' if color else ""
    lst = f'<a:lstStyle><a:lvl1pPr><a:defRPr{rpr_attr}>{fill}</a:defRPr></a:lvl1pPr></a:lstStyle>' if (rpr_attr or fill) else '<a:lstStyle/>'
    body = '<a:bodyPr' + (f' anchor="{anchor}"' if anchor else "") + '/>'
    if ph_attr.startswith('type="sldNum"'):
        para = '<a:p><a:pPr algn="r"/><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:t>‹#›</a:t></a:fld></a:p>'
    else:
        para = f'<a:p><a:r><a:rPr lang="ja-JP"/><a:t>{prompt}</a:t></a:r></a:p>' if prompt else '<a:p><a:endParaRPr lang="ja-JP"/></a:p>'
    custom = ' hasCustomPrompt="1"' if prompt else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph {ph_attr}{custom}/></p:nvPr></p:nvSpPr>{sppr}'
            f'<p:txBody>{body}{lst}{para}</p:txBody></p:sp>')


def rect(name, box, color):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(*box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP"/></a:p></p:txBody></p:sp>')


def textbox(name, box, text, color, size=900):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(*box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:r><a:rPr lang="en-US" sz="{size}"><a:solidFill><a:schemeClr val="{color}"/></a:solidFill></a:rPr>'
            f'<a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>')


def sptree(shapes):
    return ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) + '</p:spTree>')


HORZ = ' orient="horz"'


def guides(uri, items):
    g = "".join(f'<p15:guide id="{i + 1}"{HORZ if o == "h" else ""} pos="{round(emu / EMU_PER_GUIDE)}" userDrawn="1">'
                f'<p15:clr><a:srgbClr val="{"E46962" if lvl == "master" else "FBAE40"}"/></p15:clr></p15:guide>'
                for i, (o, emu, lvl) in enumerate(items))
    return f'<p:extLst><p:ext uri="{uri}"><p15:sldGuideLst {P15}>{g}</p15:sldGuideLst></p:ext></p:extLst>'


def out_name(suffix):
    """出力ファイル名。ブランドの slug から作る（例：bigtree_lab_sample.pptx）"""
    return f"{SLUG}{suffix}"
FOOT_Y, FOOT_H = H - M + 76200, 228600
SLDNUM_BOX = (W - M - 914400, FOOT_Y, 914400, FOOT_H)
TITLE_BOX = (M, M, W - 2 * M, 609600)
BODY_BOX = (M, 1828800, W - 2 * M, H - 1828800 - M)


# ---------------------------------------------------------------- master
def master(layout_count):
    _id[0] = 1
    shapes = [
        ph("Title", 'type="title"', TITLE_BOX),
        ph("Body", 'type="body" idx="1"', BODY_BOX),
        ph("SlideNumber", 'type="sldNum" sz="quarter" idx="12"', SLDNUM_BOX, size=900, color="tx2"),
        textbox("Copyright", (M, FOOT_Y, 4572000, FOOT_H), COPYRIGHT, "tx2"),
    ]
    lvl = lambda n, sz, bu: (f'<a:lvl{n}pPr marL="{(n - 1) * 342900 + (228600 if bu else 0)}" indent="{-228600 if bu else 0}">'
                             f'<a:lnSpc><a:spcPct val="130000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef>'
                             + ('<a:buFont typeface="Arial"/><a:buChar char="•"/>' if bu else '<a:buNone/>')
                             + f'<a:defRPr sz="{sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                             f'<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl{n}pPr>')
    tx = ('<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2800" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
          '<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
          f'<p:bodyStyle>{lvl(1, 1800, False)}{lvl(2, 1600, True)}{lvl(3, 1400, True)}</p:bodyStyle>'
          '<p:otherStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr></p:otherStyle></p:txStyles>')
    ids = "".join(f'<p:sldLayoutId id="{2147483649 + i}" r:id="rId{i + 1}"/>' for i in range(layout_count))
    g = guides("{27BBF7A9-308A-43DC-89C8-2F10F3537804}",
               [("v", M, "master"), ("v", W - M, "master"), ("h", M, "master"), ("h", H - M, "master")])
    return (XML + f'<p:sldMaster {NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
            f'{sptree(shapes)}</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            f'<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>{tx}{g}</p:sldMaster>')


# ---------------------------------------------------------------- layouts
def layout(name, shapes, guide_items, bg=None, show_master=True):
    _id[0] = 1
    bgxml = f'<p:bg><p:bgPr><a:solidFill><a:schemeClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' if bg else ""
    sm = "" if show_master else ' showMasterSp="0"'
    g = guides("{DCECCB84-F9BA-43D5-87BE-67443E8EF086}", guide_items) if guide_items else ""
    return (XML + f'<p:sldLayout {NS} preserve="1" userDrawn="1"{sm}><p:cSld name="{name}">{bgxml}{sptree(shapes)}</p:cSld>'
            f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{g}</p:sldLayout>')


def layouts():
    sn = lambda color="tx2": ph("SlideNumber", 'type="sldNum" sz="quarter" idx="12"', SLDNUM_BOX, size=900, color=color)
    cx = W // 2
    title = layout("01_表紙", [
        rect("Decor_Band", (0, 0, 228600, H), "accent1"),
        ph("Title", 'type="ctrTitle"', (M * 2, 2286000, W - M * 3, 1066800), "プレゼンテーションタイトル", size=4000, anchor="b"),
        rect("Decor_Line", (M * 2, 3429000, 1828800, 38100), "accent3"),
        ph("Subtitle", 'type="subTitle" idx="1"', (M * 2, 3581400, W - M * 3, 457200), "サブタイトル", size=1800, color="tx2"),
        ph("Meta", 'type="body" sz="quarter" idx="13"', (M * 2, 5029200, W - M * 3, 457200), "日付 ｜ 所属 ｜ 発表者", size=1200, color="tx2"),
        sn(),
    ], [("h", 3429000, "l"), ("v", M * 2, "l")])
    dark = layout("02_構造_Dark", [
        ph("Title", 'type="title"', (M * 2, 2743200, W - M * 4, 1371600), "セクションタイトル", size=3600, color="bg1", anchor="ctr"),
        rect("Decor_Line", (M * 2, 4191000, 914400, 38100), "accent2"),
        textbox("Copyright", (M, FOOT_Y, 4572000, FOOT_H), COPYRIGHT, "bg1"),
        sn("bg1"),
    ], [("h", H // 2, "l"), ("v", cx, "l")], bg="tx1", show_master=False)
    light = layout("03_構造_Light", [
        rect("Decor_Band", (0, 0, W, 228600), "accent1"),
        ph("Title", 'type="title"', (M * 2, 2743200, W - M * 4, 1371600), "セクションタイトル", size=3600, anchor="ctr"),
        rect("Decor_Line", (M * 2, 4191000, 914400, 38100), "accent3"),
        sn(),
    ], [("h", H // 2, "l"), ("v", cx, "l")], bg="bg2")
    body = layout("04_本文", [
        ph("Title", 'type="title"', TITLE_BOX, "スライドタイトル", size=2400),
        rect("Decor_Line", (M, M + 609600 + 76200, W - 2 * M, 12700), "accent1"),
        ph("Lead", 'type="body" sz="quarter" idx="13"', (M, 1219200, W - 2 * M, 457200), "このスライドで言いたいことを1行で", size=1600, bold=1, color="accent1"),
        ph("Content", 'type="body" idx="1"', BODY_BOX, None),
        sn(),
    ], [("h", M + 609600 + 76200, "l"), ("h", 1828800, "l"), ("v", cx, "l")])
    return [title, dark, light, body]


# ---------------------------------------------------------------- sample slides
def slide(shapes):
    _id[0] = 1
    return (XML + f'<p:sld {NS}><p:cSld>{sptree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def s_ph(name, ph_attr, text):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="{name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph {ph_attr}/></p:nvPr></p:nvSpPr><p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/>'
            + "".join(f'<a:p><a:pPr lvl="{lv}"/><a:r><a:rPr lang="ja-JP"/><a:t>{t}</a:t></a:r></a:p>' for lv, t in text)
            + '</p:txBody></p:sp>')


def s_num():
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{nid()}" name="SlideNumber"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr><p:spPr/>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum">'
            '<a:rPr lang="ja-JP"/><a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>')


def slides():
    return [
        (1, slide([s_ph("Title", 'type="ctrTitle"', [(0, "Claude Code 製 POTX テスト")]),
                   s_ph("Subtitle", 'type="subTitle" idx="1"', [(0, "XML を直接書いて組み立てたテンプレート")]),
                   s_ph("Meta", 'type="body" sz="quarter" idx="13"', [(0, BRAND_INFO["meta"])])])),
        (2, slide([s_ph("Title", 'type="title"', [(0, "第1章　背景")]), s_num()])),
        (3, slide([s_ph("Title", 'type="title"', [(0, "第2章　検証")]), s_num()])),
        (4, slide([s_ph("Title", 'type="title"', [(0, "テーマ・レイアウト・ガイドを XML で設定できた")]),
                   s_ph("Lead", 'type="body" sz="quarter" idx="13"', [(0, "Claude Code は POTX の土台を直接書ける")]),
                   s_ph("Content", 'type="body" idx="1"',
                        [(0, "本文 第1レベル：Noto Sans JP / English text"), (1, "第2レベル（箇条書き）"), (2, "第3レベル")]),
                   s_num()])),
    ]


# ---------------------------------------------------------------- font embedding (EOT, uncompressed)
def utf16z(s):
    return s.encode("utf-16le")


def make_eot(ttf_path):
    from fontTools.ttLib import TTFont
    data = Path(ttf_path).read_bytes()
    f = TTFont(ttf_path, lazy=True)
    os2, name = f["OS/2"], f["name"]
    fam = name.getDebugName(1)
    style = name.getDebugName(2)
    ver = name.getDebugName(5)
    full = name.getDebugName(4)
    p = os2.panose
    panose = bytes([p.bFamilyType, p.bSerifStyle, p.bWeight, p.bProportion, p.bContrast,
                    p.bStrokeVariation, p.bArmStyle, p.bLetterForm, p.bMidline, p.bXHeight])
    body = b"".join(struct.pack("<H", 0) + struct.pack("<H", len(utf16z(s))) + utf16z(s)
                    for s in (fam, style, ver, full))[2:]  # first field has no leading padding
    header = (struct.pack("<I", 0x00020001) + struct.pack("<I", 0) + panose
              + struct.pack("<BBIHH", 128, 1 if os2.fsSelection & 1 else 0, os2.usWeightClass, os2.fsType, 0x504C)
              + struct.pack("<IIII", os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3, os2.ulUnicodeRange4)
              + struct.pack("<II", os2.ulCodePageRange1, os2.ulCodePageRange2)
              + struct.pack("<I", f["head"].checkSumAdjustment) + b"\0" * 16
              + struct.pack("<H", 0) + body
              + struct.pack("<HH", 0, 0))  # Padding5, RootStringSize
    size = 8 + len(header) + len(data)
    return struct.pack("<II", size, len(data)) + header + data


# (typeface, [(slot, ttf)])。書体は tokens.json から（見出し書体が別なら2書体×2ウェイト）。
# 以前は "Noto Sans JP" を直書きしていたため、font.family を変えても本文の書体が埋め込まれなかった
# **使っているウェイトだけ**を入れる：見出し書体が Bold しか使わないなら Regular は入れない。
# 理由：和文フォントは1ウェイト 5〜8MB あり、4つ入れると 1ファイル 16MB になった（2書体目の Regular は1文字も使っていなかった）。
# 本文書体は Regular / Bold の両方を必ず入れる（人がスライド上で太字を切り替えるため）。
def _embed_list():
    used = {}
    for r in _TOKENS["type_pptx"]["roles"].values():
        used.setdefault(role_family(r), set()).add("bold" if r["weight"] == "Bold" else "regular")
    used[FONT] = {"regular", "bold"}
    return [(fam, [(slot, f'{fam.replace(" ", "")}-{slot.capitalize()}.ttf') for slot in ("regular", "bold") if slot in used[fam]])
            for fam in dict.fromkeys((FONT, FONT_HEADING))]


EMBED = _embed_list()


# ---------------------------------------------------------------- package
EMB = ' embedTrueTypeFonts="1"'


def build(path, template, embed):
    lays = layouts()
    sl = [] if template else slides()
    files, over = {}, []
    main = "presentationml.template.main+xml" if template else "presentationml.presentation.main+xml"
    over.append(("/ppt/presentation.xml", f"{CT}.{main}"))

    pres_rels = [("slideMaster", "slideMasters/slideMaster1.xml"), ("theme", "theme/theme1.xml"),
                 ("presProps", "presProps.xml"), ("viewProps", "viewProps.xml"), ("tableStyles", "tableStyles.xml")]
    for i, _ in enumerate(sl):
        pres_rels.append(("slide", f"slides/slide{i + 1}.xml"))
    font_xml = ""
    if embed:
        n = 0
        for face, slots in EMBED:
            refs = ""
            for slot, ttf in slots:
                n += 1
                files[f"ppt/fonts/font{n}.fntdata"] = make_eot(font_file(ttf.split("-")[-1].replace(".ttf", ""), face))
                pres_rels.append(("font", f"fonts/font{n}.fntdata"))
                refs += f'<p:{slot} r:id="rId{len(pres_rels)}"/>'
            font_xml += f'<p:embeddedFont><p:font typeface="{face}" charset="-128"/>{refs}</p:embeddedFont>'
        font_xml = f"<p:embeddedFontLst>{font_xml}</p:embeddedFontLst>"

    rels = lambda items: XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(
        f'<Relationship Id="rId{i + 1}" Type="{REL}/{t}" Target="{tg}"/>' for i, (t, tg) in enumerate(items)) + '</Relationships>'

    sld_ids = "".join(f'<p:sldId id="{256 + i}" r:id="rId{6 + i}"/>' for i in range(len(sl)))
    files["ppt/presentation.xml"] = (
        XML + f'<p:presentation {NS} saveSubsetFonts="0"{EMB if embed else ""}>'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        + (f"<p:sldIdLst>{sld_ids}</p:sldIdLst>" if sl else "")
        + f'<p:sldSz cx="{W}" cy="{H}"/><p:notesSz cx="6858000" cy="9144000"/>{font_xml}'
        '<p:defaultTextStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr>'
        '<a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
        '</p:defaultTextStyle></p:presentation>')
    files["ppt/_rels/presentation.xml.rels"] = rels(pres_rels)
    files["ppt/theme/theme1.xml"] = theme()
    files["ppt/slideMasters/slideMaster1.xml"] = master(len(lays))
    files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = rels(
        [("slideLayout", f"../slideLayouts/slideLayout{i + 1}.xml") for i in range(len(lays))] + [("theme", "../theme/theme1.xml")])
    for i, x in enumerate(lays):
        files[f"ppt/slideLayouts/slideLayout{i + 1}.xml"] = x
        files[f"ppt/slideLayouts/_rels/slideLayout{i + 1}.xml.rels"] = rels([("slideMaster", "../slideMasters/slideMaster1.xml")])
        over.append((f"/ppt/slideLayouts/slideLayout{i + 1}.xml", f"{CT}.presentationml.slideLayout+xml"))
    for i, (li, x) in enumerate(sl):
        files[f"ppt/slides/slide{i + 1}.xml"] = x
        files[f"ppt/slides/_rels/slide{i + 1}.xml.rels"] = rels([("slideLayout", f"../slideLayouts/slideLayout{li}.xml")])
        over.append((f"/ppt/slides/slide{i + 1}.xml", f"{CT}.presentationml.slide+xml"))
    files["ppt/presProps.xml"] = XML + f'<p:presentationPr {NS}/>'
    files["ppt/viewProps.xml"] = XML + f'<p:viewPr {NS}><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
    files["ppt/tableStyles.xml"] = XML + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
    files["docProps/core.xml"] = (XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                                  'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>POTX capability test</dc:title>'
                                  '<dc:creator>Claude Code</dc:creator></cp:coreProperties>')
    files["docProps/app.xml"] = (XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                                 '<Application>Claude Code</Application></Properties>')
    files["_rels/.rels"] = rels([("officeDocument", "ppt/presentation.xml")]).replace(
        "</Relationships>",
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        f'<Relationship Id="rId3" Type="{REL}/extended-properties" Target="docProps/app.xml"/></Relationships>')
    over += [("/ppt/slideMasters/slideMaster1.xml", f"{CT}.presentationml.slideMaster+xml"),
             ("/ppt/theme/theme1.xml", f"{CT}.theme+xml"),
             ("/ppt/presProps.xml", f"{CT}.presentationml.presProps+xml"),
             ("/ppt/viewProps.xml", f"{CT}.presentationml.viewProps+xml"),
             ("/ppt/tableStyles.xml", f"{CT}.presentationml.tableStyles+xml"),
             ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
             ("/docProps/app.xml", f"{CT}.extended-properties+xml")]
    files["[Content_Types].xml"] = (
        XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="fntdata" ContentType="application/x-fontdata"/>'
        + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in over) + '</Types>')

    order = ["[Content_Types].xml", "_rels/.rels"] + [k for k in files if k not in ("[Content_Types].xml", "_rels/.rels")]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k in order:
            v = files[k]
            z.writestr(k, v if isinstance(v, bytes) else v.encode("utf-8"))
    print(f"wrote {path.relative_to(OUT.parent)} ({path.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    embed = "--embed" in sys.argv
    OUT.mkdir(exist_ok=True)
    suffix = "_embed" if embed else ""
    build(OUT / f"test{suffix}.potx", template=True, embed=embed)
    build(OUT / f"test{suffix}_sample.pptx", template=False, embed=embed)
