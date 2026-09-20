"""Figma layout export (design/figma_layouts.json) + tokens (design/tokens.json) -> POTX and a sample PPTX.

Flow: Figma (SSOT) --use_figma export--> design/figma_layouts.json --this script--> build/bigtree_lab.potx
Units: Figma px (1440x810) -> EMU at 6350 EMU/px (10in slide), font px/2 = pt.
Usage: .venv/bin/python scripts/build_potx_figma.py [--no-embed]
"""
import json
import re
import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402  (theme, EOT embedding, guides helper)

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND", ROOT / "bigtree")).resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT", ROOT / "build")).resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
SPEC = json.loads((BRAND / "design/figma_layouts.json").read_text())
EMU = bp.W / SPEC["frame"][0]  # 6350
XML, NS, REL, CT = bp.XML, bp.NS, bp.REL, bp.CT
SCHEME = {"dk1": "tx1", "lt1": "bg1", "dk2": "tx2", "lt2": "bg2"}
STYLES = bp._TOKENS["type_pptx"]["roles"]   # 文字スタイルの正は tokens.json


def e(px):
    return int(round(px * EMU))


def clr(var, alpha=None):
    a = f'<a:alpha val="{int(alpha * 100000)}"/>' if alpha is not None else ""
    return f'<a:solidFill><a:schemeClr val="{SCHEME.get(var, var)}">{a}</a:schemeClr></a:solidFill>'


def xfrm(box):
    x, y, w, h = box
    return f'<a:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(h)}"/></a:xfrm>'


def rpr(style, color, tag="a:defRPr"):
    s = STYLES[style]
    b = ' b="1"' if s["weight"] == "Bold" else ' b="0"'
    font = ('<a:latin typeface="Noto Sans JP Black"/><a:ea typeface="Noto Sans JP Black"/>' if s["weight"] == "Black"
            else '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>')
    return f'<{tag} sz="{s["pt"] * 100}"{b}>{clr(color)}{font}</{tag.split()[0]}>'


def lst(style, color, align="l"):
    # line spacing in exact points (Figma px * line% / 2). PowerPoint's percent spacing is relative to the
    # font's own leading (~1.2x) and piles the slack above the line, which pushes text down vs Figma.
    s = STYLES[style]
    pts = round(s["pt"] * s["line"] / 100 * 100)
    return (f'<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="{align}"><a:lnSpc><a:spcPts val="{pts}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:buNone/>{rpr(style, color)}</a:lvl1pPr></a:lstStyle>')


def paras(text):
    return "".join(f'<a:p><a:r><a:rPr lang="ja-JP"/><a:t>{bp_esc(line)}</a:t></a:r></a:p>' for line in text.split("\n"))


def bp_esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


ZERO = 'lIns="0" tIns="0" rIns="0" bIns="0"'


# ---------------------------------------------------------------- logo geometry (from the SVG made by dxf_to_svg.py)
def logo_paths():
    svg = (BRAND / "assets/logo/260920_BT_logo.svg").read_text()
    vw, vh = [float(v) for v in re.search(r'viewBox="0 0 ([\d.]+) ([\d.]+)"', svg).groups()]
    d = re.search(r' d="([^"]+)"', svg).group(1)
    subs = [[tuple(float(v) for v in p.split(",")) for p in re.split(r"\s*L\s*", s.strip())] for s in re.findall(r"M([^Z]+)Z", d)]
    return vw, vh, subs


LOGO_W, LOGO_H, LOGO_SUBS = logo_paths()


def logo(item):
    k = 1000  # path units: 1/1000 of an svg unit
    paths = "".join(
        f'<a:path w="{int(LOGO_W * k)}" h="{int(LOGO_H * k)}"><a:moveTo><a:pt x="{int(s[0][0] * k)}" y="{int(s[0][1] * k)}"/></a:moveTo>'
        + "".join(f'<a:lnTo><a:pt x="{int(x * k)}" y="{int(y * k)}"/></a:lnTo>' for x, y in s[1:]) + '<a:close/></a:path>'
        for s in LOGO_SUBS)
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}" descr="Bigtree Lab ロゴ"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(item["box"])}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst>{paths}</a:pathLst></a:custGeom>{clr(item["color"], item.get("opacity"))}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>')


# ---------------------------------------------------------------- key visual: picture cropped to the Figma curve
def picture(item, rid):
    m = item["mask"]
    ox, oy = m["origin"]
    toks = m["curve"].replace(",", " ").split()
    nums = [float(t) for t in toks if not t.isalpha()]
    pts = [(nums[i] + ox, nums[i + 1] + oy) for i in range(0, len(nums), 2)]  # frame px
    left = min(x for x, _ in pts)
    box = [left, 0, m["right"] - left, SPEC["frame"][1]]
    # path in box-local EMU (may extend above/below the box; the slide edge clips it)
    P = lambda x, y: f'<a:pt x="{e(x - left)}" y="{e(y)}"/>'
    segs = f'<a:moveTo>{P(*pts[0])}</a:moveTo>'
    for i in range(1, len(pts), 3):
        segs += f'<a:cubicBezTo>{P(*pts[i])}{P(*pts[i + 1])}{P(*pts[i + 2])}</a:cubicBezTo>'
    last, first = pts[-1], pts[0]
    segs += (f'<a:lnTo>{P(last[0], m["bottom"])}</a:lnTo><a:lnTo>{P(m["right"], m["bottom"])}</a:lnTo>'
             f'<a:lnTo>{P(m["right"], m["top"])}</a:lnTo><a:lnTo>{P(first[0], m["top"])}</a:lnTo><a:close/>')
    # FILL crop of the image inside Figma's fill box, re-expressed for the picture box
    fx, fy, fw, fh = item["fill_box"]
    iw, ih = item["img_px"]
    sc = max(fw / iw, fh / ih)
    dw, dh = iw * sc, ih * sc
    ix, iy = fx - (dw - fw) / 2, fy - (dh - fh) / 2
    l = (box[0] - ix) / dw
    r = (ix + dw - (box[0] + box[2])) / dw
    t = (box[1] - iy) / dh
    b = (iy + dh - (box[1] + box[3])) / dh
    src = f'<a:srcRect l="{int(l * 100000)}" t="{int(t * 100000)}" r="{int(r * 100000)}" b="{int(b * 100000)}"/>'
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}" descr="メタセコイアを見上げた写真（キービジュアル）"/>'
            '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/>{src}<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{xfrm(box)}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{e(box[2])}" h="{e(box[3])}">{segs}</a:path></a:pathLst></a:custGeom></p:spPr></p:pic>')


# ---------------------------------------------------------------- shapes
def rect(item):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(item["box"])}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{clr(item["color"])}<a:ln><a:noFill/></a:ln></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP"/></a:p></p:txBody></p:sp>')


def text(item):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(item["box"])}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="none" {ZERO} anchor="ctr"><a:noAutofit/></a:bodyPr>{lst(item["style"], item["color"], item.get("align", "l"))}'
            f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{bp_esc(item["text"])}</a:t></a:r></a:p></p:txBody></p:sp>')


def ph_attr(item):
    t = item["ph"]
    typ = "" if t == "obj" else f'type="{t}" '
    idx = f'idx="{item["idx"]}"' if "idx" in item else ""
    sz = ' sz="quarter"' if t in ("sldNum", "body") and item.get("idx", 0) >= 12 else ""
    return (typ + idx + sz).strip()


def placeholder(item, on_slide=False, content=None, master=False):
    attr = ph_attr(item)
    if on_slide:  # a slide inherits position/format from the layout
        body = content if content is not None else paras(item.get("text", ""))
        if item["ph"] == "sldNum":
            body = '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ja-JP"/><a:t>‹#›</a:t></a:fld></a:p>'
        return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
                f'<p:nvPr><p:ph {attr}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody></p:sp>')
    pill = item.get("pill")
    if pill:
        px, py = pill["pad"]
        geom = (f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom><a:noFill/>'
                f'<a:ln w="{e(pill["stroke_px"])}">{clr(pill["stroke"])}</a:ln>')
        bpr = f'<a:bodyPr wrap="none" lIns="{e(px)}" tIns="{e(py)}" rIns="{e(px)}" bIns="{e(py)}" anchor="ctr"><a:spAutoFit/></a:bodyPr>'
    else:
        geom = ""
        # text-sized boxes: centering matches Figma's half-leading. The content placeholder grows with the
        # text, so it must stay top-anchored.
        anchor = ' anchor="t"' if item["ph"] == "obj" else ' anchor="ctr"'
        wrap = "square" if item["ph"] == "obj" else "none"
        bpr = f'<a:bodyPr wrap="{wrap}" {ZERO}{anchor}><a:noAutofit/></a:bodyPr>'
    if item["ph"] == "sldNum":
        body = '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="ja-JP"/><a:t>‹#›</a:t></a:fld></a:p>'
        prompt = ""
    else:
        body = paras(item["text"])
        prompt = "" if master else ' hasCustomPrompt="1"'
    style = lst(item["style"], item["color"], item.get("align", "l")) if item["ph"] != "obj" else "<a:lstStyle/>"
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f'<p:nvPr><p:ph {attr}{prompt}/></p:nvPr></p:nvSpPr><p:spPr>{xfrm(item["box"])}{geom}</p:spPr>'
            f'<p:txBody>{bpr}{style}{body}</p:txBody></p:sp>')


def guides_xml(pairs, uri, color):
    items = [("v" if a == "X" else "h", px * EMU, color) for a, px in pairs]
    return bp.guides(uri, items)


# ---------------------------------------------------------------- table style (Figma has no table primitive:
# the spec lives here and is rendered by PowerPoint; 0304 rule = no outer frame, horizontal rules only)
TABLE_STYLE_ID = "{6E5B4C3A-2D1F-4A8B-9C7D-0E1F2A3B4C5D}"


def table_styles():
    cap = STYLES["caption"]
    txt = lambda color, bold: (f'<a:tcTxStyle b="{"on" if bold else "off"}"><a:fontRef idx="minor"><a:schemeClr val="{SCHEME.get(color, color)}"/></a:fontRef>'
                               f'<a:schemeClr val="{SCHEME.get(color, color)}"/></a:tcTxStyle>')
    ln = lambda w, color: (f'<a:ln w="{e(w)}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="{SCHEME.get(color, color)}"/></a:solidFill></a:ln>')
    cell = lambda borders, fill: f'<a:tcStyle><a:tcBdr>{borders}</a:tcBdr>{fill}</a:tcStyle>'
    none = '<a:fill><a:noFill/></a:fill>'
    fill_of = lambda c: f'<a:fill><a:solidFill><a:schemeClr val="{SCHEME.get(c, c)}"/></a:solidFill></a:fill>'
    # borders not mentioned fall back to PowerPoint's default thin frame, so "no line" must be explicit.
    # order matters: left, right, top, bottom, insideH, insideV
    nl = "<a:ln><a:noFill/></a:ln>"
    inside = ("<a:left>" + nl + "</a:left><a:right>" + nl + "</a:right><a:top>" + nl + "</a:top><a:bottom>" + nl + "</a:bottom>"
              + "<a:insideH>" + ln(1, "accent5") + "</a:insideH><a:insideV>" + nl + "</a:insideV>")
    header_bdr = ("<a:left>" + nl + "</a:left><a:right>" + nl + "</a:right><a:top>" + nl + "</a:top>"
                  + "<a:bottom>" + ln(2, "accent1") + "</a:bottom><a:insideV>" + nl + "</a:insideV>")
    return (XML + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'def="{TABLE_STYLE_ID}"><a:tblStyle styleId="{TABLE_STYLE_ID}" styleName="Bigtree Lab 表">'
            # whole table: horizontal rules only, no outer frame (0304 rule)
            f'<a:wholeTbl>{cell(inside, none)}{txt("dk1", False)}</a:wholeTbl>'
            f'<a:band2H>{cell("", fill_of("lt2"))}</a:band2H>'
            # header row: accent1 band, white bold text, thicker rule underneath
            f'<a:firstRow>{txt("lt1", True)}{cell(header_bdr, fill_of("accent1"))}</a:firstRow>'
            # first column: bold row labels
            f'<a:firstCol>{txt("dk1", True)}{cell("", none)}</a:firstCol>'
            '</a:tblStyle></a:tblStyleLst>')


# ---------------------------------------------------------------- master / layouts
def master(n_layouts):
    bp._id[0] = 1
    body_layout = next(l for l in SPEC["layouts"] if l["name"] == "06_本文")
    it = {i["name"]: i for i in body_layout["items"]}
    # master placeholders: explicit types (a master body placeholder must be type="body"), no custom prompt
    shapes = [placeholder({**it["Title"], "text": "マスター タイトル"}, master=True),
              placeholder({**it["Content"], "ph": "body", "text": "マスター テキスト"}, master=True),
              placeholder(it["SlideNumber"], master=True)]
    b = STYLES["body"]
    lvl = lambda n, sz, bullet: (
        f'<a:lvl{n}pPr marL="{e(0 if n == 1 else 28 * (n - 1) + 8)}" indent="{0 if n == 1 else -e(20)}">'
        f'<a:lnSpc><a:spcPct val="{b["line"] * 1000}"/></a:lnSpc><a:spcBef><a:spcPts val="{600 if n == 1 else 300}"/></a:spcBef>'
        + (f'<a:buFont typeface="Noto Sans JP"/><a:buChar char="{bullet}"/>' if bullet else '<a:buNone/>')
        + f'<a:defRPr sz="{sz * 100}">{clr("dk1")}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl{n}pPr>')
    t = STYLES["title"]
    tx = (f'<p:txStyles><p:titleStyle><a:lvl1pPr><a:lnSpc><a:spcPct val="{t["line"] * 1000}"/></a:lnSpc>'
          f'<a:defRPr sz="{t["pt"] * 100}" b="1">{clr("dk1")}<a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle>'
          f'<p:bodyStyle>{lvl(1, b["pt"], None)}{lvl(2, b["pt"], "•")}{lvl(3, STYLES["caption"]["pt"] + 1, "◦")}</p:bodyStyle>'
          '<p:otherStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr></p:otherStyle></p:txStyles>')
    ids = "".join(f'<p:sldLayoutId id="{2147483649 + i}" r:id="rId{i + 1}"/>' for i in range(n_layouts))
    g = guides_xml(SPEC["master_guides"], "{27BBF7A9-308A-43DC-89C8-2F10F3537804}", "master")
    return (XML + f'<p:sldMaster {NS}><p:cSld><p:bg><p:bgPr>{clr("lt1")}<a:effectLst/></p:bgPr></p:bg>{bp.sptree(shapes)}</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            f'<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>{tx}{g}</p:sldMaster>')


def layout(spec):
    bp._id[0] = 1
    shapes, rels = [], [("slideLayout_master", None)]
    img = None
    for it in spec["items"]:
        k = it["kind"]
        if k == "picture":
            img = it["image"]
            shapes.append(picture(it, "rId2"))
        elif k == "logo":
            shapes.append(logo(it))
        elif k == "rect":
            shapes.append(rect(it))
        elif k == "text":
            shapes.append(text(it))
        elif k == "ph":
            shapes.append(placeholder(it))
    g = guides_xml(spec["guides"], "{DCECCB84-F9BA-43D5-87BE-67443E8EF086}", "layout")
    xml = (XML + f'<p:sldLayout {NS} preserve="1" userDrawn="1" showMasterSp="0"><p:cSld name="{spec["name"]}">'
           f'<p:bg><p:bgPr>{clr(spec["bg"])}<a:effectLst/></p:bgPr></p:bg>{bp.sptree(shapes)}</p:cSld>'
           f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{g}</p:sldLayout>')
    return xml, img


# ---------------------------------------------------------------- sample slides (one per layout)
SAMPLE_CONTENT = {
    "06_本文": [(0, "連載 #3「AIエージェント時代のnote執筆」が検索経由で安定して読まれている"),
               (1, "検索流入は前月比 142%。「note 自動投稿」「Claude Code note」が上位"),
               (0, "スキ率（スキ ÷ PV）は 4.8% → 6.1% に改善"),
               (0, "次の一手：連載 #4 は Office テンプレート編。9月中に公開する")],
    "07_本文_リードなし": [(0, "#1 はじめに：AI と一緒に note を書く ─ PV 2,140 ／ スキ 96"),
                      (0, "#2 挿絵は HTML で作る ─ PV 1,820 ／ スキ 121"),
                      (0, "#3 眠っていた才能 ─ PV 3,560 ／ スキ 214"),
                      (1, "検索流入が全体の 58%")],
}


def sample_slide(spec):
    bp._id[0] = 1
    shapes = []
    for it in spec["items"]:
        if it["kind"] != "ph":
            continue
        if it["ph"] == "obj":
            lines = SAMPLE_CONTENT.get(spec["name"], [])
            body = "".join(f'<a:p><a:pPr lvl="{lv}"/><a:r><a:rPr lang="ja-JP"/><a:t>{bp_esc(t)}</a:t></a:r></a:p>' for lv, t in lines)
            shapes.append(placeholder(it, on_slide=True, content=body))
        else:
            shapes.append(placeholder(it, on_slide=True))
    return XML + f'<p:sld {NS}><p:cSld>{bp.sptree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'


# ---------------------------------------------------------------- package
def rels_xml(items):
    return (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i + 1}" Type="{t}" Target="{tg}"/>' for i, (t, tg) in enumerate(items))
            + '</Relationships>')


def R(t):
    return f"{REL}/{t}"


def build(path, template, embed):
    files, over = {}, []
    lays = [layout(s) for s in SPEC["layouts"]]
    media = {}
    for i, (xml, img) in enumerate(lays, 1):
        files[f"ppt/slideLayouts/slideLayout{i}.xml"] = xml
        lr = [(R("slideLayout").replace("slideLayout", "slideMaster"), "../slideMasters/slideMaster1.xml")]
        if img:
            if img not in media:
                media[img] = f"image{len(media) + 1}{Path(img).suffix}"
                files[f"ppt/media/{media[img]}"] = (BRAND / img).read_bytes()
            lr.append((R("image"), f"../media/{media[img]}"))
        files[f"ppt/slideLayouts/_rels/slideLayout{i}.xml.rels"] = rels_xml(lr)
        over.append((f"/ppt/slideLayouts/slideLayout{i}.xml", f"{CT}.presentationml.slideLayout+xml"))
    files["ppt/slideMasters/slideMaster1.xml"] = master(len(lays))
    files["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = rels_xml(
        [(R("slideLayout"), f"../slideLayouts/slideLayout{i + 1}.xml") for i in range(len(lays))] + [(R("theme"), "../theme/theme1.xml")])
    files["ppt/theme/theme1.xml"] = bp.theme()

    pres_rels = [(R("slideMaster"), "slideMasters/slideMaster1.xml"), (R("theme"), "theme/theme1.xml"),
                 (R("presProps"), "presProps.xml"), (R("viewProps"), "viewProps.xml"), (R("tableStyles"), "tableStyles.xml")]
    sld_ids = ""
    if not template:
        for i, s in enumerate(SPEC["layouts"], 1):
            files[f"ppt/slides/slide{i}.xml"] = sample_slide(s)
            files[f"ppt/slides/_rels/slide{i}.xml.rels"] = rels_xml([(R("slideLayout"), f"../slideLayouts/slideLayout{i}.xml")])
            over.append((f"/ppt/slides/slide{i}.xml", f"{CT}.presentationml.slide+xml"))
            pres_rels.append((R("slide"), f"slides/slide{i}.xml"))
            sld_ids += f'<p:sldId id="{255 + i}" r:id="rId{len(pres_rels)}"/>'
    font_xml = ""
    if embed:
        n = 0
        for face, slots in bp.EMBED:
            refs = ""
            for slot, ttf in slots:
                n += 1
                files[f"ppt/fonts/font{n}.fntdata"] = bp.make_eot(bp.FONT_DIR / ttf)
                pres_rels.append((R("font"), f"fonts/font{n}.fntdata"))
                refs += f'<p:{slot} r:id="rId{len(pres_rels)}"/>'
            font_xml += f'<p:embeddedFont><p:font typeface="{face}" charset="-128"/>{refs}</p:embeddedFont>'
        font_xml = f"<p:embeddedFontLst>{font_xml}</p:embeddedFontLst>"
    main = "presentationml.template.main+xml" if template else "presentationml.presentation.main+xml"
    files["ppt/presentation.xml"] = (
        XML + f'<p:presentation {NS} saveSubsetFonts="0"{bp.EMB if embed else ""}>'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        + (f"<p:sldIdLst>{sld_ids}</p:sldIdLst>" if sld_ids else "")
        + f'<p:sldSz cx="{bp.W}" cy="{bp.H}"/><p:notesSz cx="6858000" cy="9144000"/>{font_xml}'
        '<p:defaultTextStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr><a:lvl1pPr><a:defRPr sz="1400">'
        f'{clr("dk1")}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>')
    files["ppt/_rels/presentation.xml.rels"] = rels_xml(pres_rels)
    files["ppt/presProps.xml"] = XML + f'<p:presentationPr {NS}/>'
    files["ppt/viewProps.xml"] = XML + f'<p:viewPr {NS}><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
    files["ppt/tableStyles.xml"] = table_styles()
    files["docProps/core.xml"] = (XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                                  'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Bigtree Lab template</dc:title>'
                                  '<dc:creator>Claude Code</dc:creator></cp:coreProperties>')
    files["docProps/app.xml"] = (XML + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                                 '<Application>Claude Code</Application></Properties>')
    files["_rels/.rels"] = rels_xml([(R("officeDocument"), "ppt/presentation.xml"),
                                     ("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml"),
                                     (R("extended-properties"), "docProps/app.xml")])
    over += [("/ppt/presentation.xml", f"{CT}.{main}"),
             ("/ppt/slideMasters/slideMaster1.xml", f"{CT}.presentationml.slideMaster+xml"),
             ("/ppt/theme/theme1.xml", f"{CT}.theme+xml"),
             ("/ppt/presProps.xml", f"{CT}.presentationml.presProps+xml"),
             ("/ppt/viewProps.xml", f"{CT}.presentationml.viewProps+xml"),
             ("/ppt/tableStyles.xml", f"{CT}.presentationml.tableStyles+xml"),
             ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
             ("/docProps/app.xml", f"{CT}.extended-properties+xml")]
    files["[Content_Types].xml"] = (
        XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="fntdata" ContentType="application/x-fontdata"/>'
        + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in over) + '</Types>')
    order = ["[Content_Types].xml", "_rels/.rels"] + [k for k in files if k not in ("[Content_Types].xml", "_rels/.rels")]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k in order:
            v = files[k]
            z.writestr(k, v if isinstance(v, bytes) else v.encode("utf-8"))
    print(f"wrote {path.relative_to(OUT.parent)} ({path.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    embed = "--no-embed" not in sys.argv
    build(OUT / "bigtree_lab.potx", template=True, embed=embed)
    build(OUT / "bigtree_lab_sample.pptx", template=False, embed=embed)
