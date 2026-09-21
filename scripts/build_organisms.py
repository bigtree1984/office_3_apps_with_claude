"""Figma organism export (design/figma_organisms.json) -> PowerPoint shapes, placed in the body layout.

Generic converter: RECTANGLE / ELLIPSE -> preset shapes, VECTOR / POLYGON -> custom geometry (paths copied as-is),
TEXT -> text boxes with the design-system text styles. Each organism becomes one named group.
Input : build/<slug>_sample.pptx (from build_potx_figma.py)
Output: build/<slug>_organisms.pptx
Usage : .venv/bin/python scripts/build_organisms.py
"""
import json
import re
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402
import build_potx_figma as bpf  # noqa: E402
import build_charts as bc  # noqa: E402  (zip / rels helpers)

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND") or ROOT / "bigtree").resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT") or ROOT / "build").resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
ORG = json.loads((BRAND / "design/figma_organisms.json").read_text())["organisms"]
STYLE_KEY = {"本文": "body", "リード文": "lead", "補足・グラフ・表": "caption", "最小（ページ番号等）": "min",
             "スライドタイトル": "title", "セクション見出し": "section", "表紙タイトル": "cover"}
e, clr = bpf.e, bpf.clr
ALIGN = {"LEFT": "l", "CENTER": "ctr", "RIGHT": "r"}


def fill(it):
    if not it.get("fill"):
        return "<a:noFill/>"
    op = it.get("fillOpacity", 1)
    return clr(it["fill"], op if op < 0.999 else None)


def path_xml(it, ox, oy):
    """Figma path data (M/L/C/Z, frame coords) -> DrawingML path relative to the shape box."""
    toks = re.findall(r"[MLCZ]|-?\d*\.?\d+(?:e-?\d+)?", it["path"])
    x0, y0 = it["x"], it["y"]
    w, h = max(it["w"], 0.5), max(it["h"], 0.5)
    P = lambda x, y: f'<a:pt x="{e(float(x) - x0)}" y="{e(float(y) - y0)}"/>'
    out, i = "", 0
    while i < len(toks):
        c = toks[i]
        if c == "M":
            out += f"<a:moveTo>{P(toks[i + 1], toks[i + 2])}</a:moveTo>"; i += 3
        elif c == "L":
            out += f"<a:lnTo>{P(toks[i + 1], toks[i + 2])}</a:lnTo>"; i += 3
        elif c == "C":
            out += f"<a:cubicBezTo>{P(toks[i + 1], toks[i + 2])}{P(toks[i + 3], toks[i + 4])}{P(toks[i + 5], toks[i + 6])}</a:cubicBezTo>"; i += 7
        elif c == "Z":
            out += "<a:close/>"; i += 1
        else:
            i += 1
    fill_attr = ' fill="none"' if it.get("open") else ""
    return (f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{e(w)}" h="{e(h)}"{fill_attr}>{out}</a:path></a:pathLst></a:custGeom>')


def shape(it, ox, oy):
    box = [it["x"] + ox, it["y"] + oy, max(it["w"], 0.5), max(it["h"], 0.5)]
    tb = ' txBox="1"' if it["type"] == "TEXT" else ""
    nv = f'<p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{it["name"]}"/><p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
    if it["type"] == "TEXT":
        key = STYLE_KEY[it["style"]]
        paras = "".join(f'<a:p><a:r><a:rPr lang="ja-JP"/><a:t>{bpf.bp_esc(line)}</a:t></a:r></a:p>' for line in it["text"].split("\n"))
        return (f'<p:sp>{nv}<p:spPr>{bpf.xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
                f'<p:txBody><a:bodyPr wrap="square" {bpf.ZERO} anchor="ctr"><a:noAutofit/></a:bodyPr>'
                f'{bpf.lst(key, it["fill"], ALIGN[it.get("align", "LEFT")])}{paras}</p:txBody></p:sp>')
    if it["type"] in ("RECTANGLE", "ELLIPSE"):
        geom = f'<a:prstGeom prst="{"rect" if it["type"] == "RECTANGLE" else "ellipse"}"><a:avLst/></a:prstGeom>'
    else:
        geom = path_xml(it, ox, oy)
    line = (f'<a:ln w="{e(it["strokeW"])}" cap="flat">{clr(it["stroke"])}<a:miter lim="800000"/></a:ln>' if it.get("stroke")
            else "<a:ln><a:noFill/></a:ln>")
    return f'<p:sp>{nv}<p:spPr>{bpf.xfrm(box)}{geom}{fill(it)}{line}</p:spPr></p:sp>'


def group(name, org, ox, oy):
    kids = "".join(shape(it, ox, oy) for it in org["items"])
    x, y, w, h = e(ox), e(oy), e(org["w"]), e(org["h"])
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{bp.nid()}" name="organism/{name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/><a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/></a:xfrm></p:grpSpPr>'
            f'{kids}</p:grpSp>')


SLIDES = [  # (organism, title, lead)  タイトルは論点、リード文で答える（design/CONTENT_RULES.md）
    ("O-B22 ビッグナンバー", "今月の到達点", "PV は伸びたが、フォロワーはまだ動いていない"),
    ("O-B11 KPIハイライト", "今月の主要指標", "スキ率が 6.1% に改善（+1.3pt）"),
    ("O-B7 乗算式フロー", "スキ数の分解", "PV より、スキ率の改善が効いている"),
    ("O-B15 ピラミッド／レイヤー", "note 執筆を支える4つの層", "下の層が安定しているほど、上の層に時間を使える"),
    ("O-B18 包含（入れ子）", "3つの層の関係", "毎回変わるのは、いちばん内側だけ"),
    ("O-B5 プロセスフロー（グルーピング）", "1本の記事ができるまで", "準備に時間をかけるほど、下書きは速くなる"),
    ("O-B8 グリッド一覧", "執筆の6工程", "型が決まると、毎回の判断が減る"),
    ("O-B4 プロセスフロー", "支援の進め方", "最後の「運用と改善」に時間を残す"),
    ("O-B13 タイムライン", "公開日の流れ", "翌週の記録までが1セット"),
    ("O-B16 関係・因果図", "フォロワーが増える道筋", "起点は検索流入の増加にある"),
    ("O-B6 ポジショニングマップ", "この連載の立ち位置", "手間はかかるが、一次情報で差がつく"),
    ("O-B20 コールアウト注釈", "ダッシュボードの見どころ", "数字より先に「どこを見るか」を決めておく"),
    ("O-B21 ツリー", "デザインシステムの構成", "上の層を変えると、下は全部作り直せる"),
    ("O-B14 人物紹介", "この連載の登場人物", "読者は非エンジニアが中心"),
]

# 構造レイアウト（04 構造 Light / 05 構造 Dark）に置く Organism。リード文が無く、見出しだけが上にある
STRUCTURE_SLIDES = [  # (organism, layout, title)
    ("O-S1 目次", "04_構造_Light", "目次"),
    ("O-S2 写真ギャラリー", "04_構造_Light", "9月の記録"),
]


def main():
    body = next(l for l in bpf.SPEC["layouts"] if l["name"] == "06_本文")
    it = {i["name"]: i for i in body["items"]}
    cx, cy, cw, ch = it["Content"]["box"]
    f = bc.read_zip(OUT / bp.out_name("_sample.pptx"))
    ct = f["[Content_Types].xml"].decode()
    pres = f["ppt/presentation.xml"].decode()
    prels = f["ppt/_rels/presentation.xml.rels"].decode()
    n = len([k for k in f if re.match(r"ppt/slides/slide\d+\.xml$", k)])
    layout_no = [l["name"] for l in bpf.SPEC["layouts"]].index("06_本文") + 1
    for i, (oname, title, lead) in enumerate(SLIDES, 1):
        org = ORG[oname]
        bp._id[0] = 1
        shapes = [bpf.placeholder({**it["Title"], "text": title}, on_slide=True),
                  bpf.placeholder({**it["Lead"], "text": lead}, on_slide=True),
                  group(oname, org, cx + (cw - org["w"]) / 2, cy + (ch - org["h"]) / 2),
                  bpf.placeholder(it["SlideNumber"], on_slide=True)]
        sn = n + i
        f[f"ppt/slides/slide{sn}.xml"] = (bp.XML + f'<p:sld {bp.NS}><p:cSld>{bp.sptree(shapes)}</p:cSld>'
                                          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>').encode()
        f[f"ppt/slides/_rels/slide{sn}.xml.rels"] = bpf.rels_xml([(bpf.R("slideLayout"), f"../slideLayouts/slideLayout{layout_no}.xml")]).encode()
        rid = f"rIdOrg{i}"
        prels = bc.add_rel(prels, rid, "slide", f"slides/slide{sn}.xml")
        pres = pres.replace("</p:sldIdLst>", f'<p:sldId id="{400 + i}" r:id="{rid}"/></p:sldIdLst>')
        ct = bc.add_override(ct, f"/ppt/slides/slide{sn}.xml", f"{bp.CT}.presentationml.slide+xml")
    # 構造レイアウト（目次・写真ギャラリー）。リード文が無く、見出しの下がそのまま本文エリア
    for j, (oname, layout_name, title) in enumerate(STRUCTURE_SLIDES, 1):
        spec = next(l for l in bpf.SPEC["layouts"] if l["name"] == layout_name)
        sit = {i["name"]: i for i in spec["items"]}
        org = ORG[oname]
        bp._id[0] = 1
        shapes = [bpf.placeholder({**sit["Title"], "text": title}, on_slide=True),
                  group(oname, org, 72, 152),
                  bpf.placeholder(sit["SlideNumber"], on_slide=True)]
        sn = n + len(SLIDES) + j
        f[f"ppt/slides/slide{sn}.xml"] = (bp.XML + f'<p:sld {bp.NS}><p:cSld>{bp.sptree(shapes)}</p:cSld>'
                                          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>').encode()
        lno = [l["name"] for l in bpf.SPEC["layouts"]].index(layout_name) + 1
        f[f"ppt/slides/_rels/slide{sn}.xml.rels"] = bpf.rels_xml([(bpf.R("slideLayout"), f"../slideLayouts/slideLayout{lno}.xml")]).encode()
        rid = f"rIdStruct{j}"
        prels = bc.add_rel(prels, rid, "slide", f"slides/slide{sn}.xml")
        pres = pres.replace("</p:sldIdLst>", f'<p:sldId id="{450 + j}" r:id="{rid}"/></p:sldIdLst>')
        ct = bc.add_override(ct, f"/ppt/slides/slide{sn}.xml", f"{bp.CT}.presentationml.slide+xml")
    f["[Content_Types].xml"], f["ppt/presentation.xml"], f["ppt/_rels/presentation.xml.rels"] = ct.encode(), pres.encode(), prels.encode()
    out = OUT / bp.out_name("_organisms.pptx")
    out.write_bytes(bc.zip_bytes(f))
    print(f"wrote {out.relative_to(OUT.parent)}")


if __name__ == "__main__":
    main()
