"""Figma organism export (design/figma_organisms.json) -> PowerPoint shapes, placed in the body layout.

Generic converter: RECTANGLE / ELLIPSE -> preset shapes, VECTOR / POLYGON -> custom geometry (paths copied as-is),
TEXT -> text boxes with the design-system text styles. Each organism becomes one named group.
Input : build/<slug>_sample.pptx (from build_potx_figma.py)
Output: build/<slug>_organisms.pptx
Usage : .venv/bin/python scripts/build_organisms.py
"""
import json
import json
import re
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402
import build_potx_figma as bpf
import fit_text as ft  # noqa: E402  (はみ出し検査。§8【変えない】でセット運用と決めている)  # noqa: E402
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


FIXED = []          # 自動で色を直した箇所（最後にまとめて報告する）


def under_fill(items, upto, box):
    """その文字の下に敷かれている色を、重なり順に混ぜながら求める。

    **一番下はスライドの地色**（白決め打ちにすると、暗地に変えた人の図解で文字が消える）。
    """
    cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
    under = bpf.BODY_BG
    for other in items[:upto]:
        if other["type"] == "TEXT":
            continue
        if other["x"] <= cx <= other["x"] + other["w"] and other["y"] <= cy <= other["y"] + other["h"]:
            under = bp.blend(other.get("fill", "dk1"), under, other.get("fillOpacity", 1))
    return under


def text_fill(it, items, idx):
    """図解の文字色。**白/黒の指定は、下地の明るさから決め直す。**

    理由：JSON に「白」と書かれた値は、作者の濃いブランド色を前提にしている。
    明るい色に差し替えた人は、そのままだと文字が読めなくなる（色を変えた人が必ず踏む）。
    accent などの意図のある色は、そのまま尊重する。
    """
    want = it["fill"]
    if want not in ("lt1", "dk1"):
        return want
    bg = under_fill(items, idx, [it["x"], it["y"], it["w"], it["h"]])
    good = bp.text_on(bg)
    if good != want:
        FIXED.append(f'{it["name"]}（{want} → {good}）')
    return good


def shape(it, ox, oy, items=None, idx=0):
    box = [it["x"] + ox, it["y"] + oy, max(it["w"], 0.5), max(it["h"], 0.5)]
    tb = ' txBox="1"' if it["type"] == "TEXT" else ""
    nv = f'<p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{it["name"]}"/><p:cNvSpPr{tb}/><p:nvPr/></p:nvSpPr>'
    if it["type"] == "TEXT":
        key = STYLE_KEY[it["style"]]
        ok, need, avail = ft.fits_box(it["text"], key, it["w"], it["h"])
        if not ok:
            ft.warn_box(it["name"], need, avail)
        it = {**it, "fill": text_fill(it, items, idx)} if items is not None else it
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
    kids = "".join(shape(it, ox, oy, org["items"], i) for i, it in enumerate(org["items"]))
    x, y, w, h = e(ox), e(oy), e(org["w"]), e(org["h"])
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{bp.nid()}" name="organism/{name}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/><a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/></a:xfrm></p:grpSpPr>'
            f'{kids}</p:grpSp>')


# スライドの題名とリード文はブランド側（design/samples.json）に置く。
# スクリプトに書くと、ブランドを差し替えた人の資料に作者の文章が残る
_S = json.loads((BRAND / "design/samples.json").read_text()) if (BRAND / "design/samples.json").exists() else {}
SLIDES = [tuple(x) for x in _S.get("organism_slides", [])]
STRUCTURE_SLIDES = [tuple(x) for x in _S.get("structure_slides", [])]


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
    if FIXED:
        print(f"（下地に合わせて文字色を自動で直しました: {len(FIXED)} か所 — {', '.join(FIXED[:4])}{' …' if len(FIXED) > 4 else ''}）")
    print(f"wrote {out.relative_to(OUT.parent)}")


if __name__ == "__main__":
    main()
