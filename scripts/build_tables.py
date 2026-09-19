"""Add table slides (O-B12 データテーブル / O-B1 比較テーブル) using the template's table style.

Tables are real PowerPoint tables, so rows/columns stay editable and the look comes from
the style defined in the template (ppt/tableStyles.xml), not from per-cell formatting.
Input : build/bigtree_lab_organisms.pptx    Output: build/bigtree_lab_tables.pptx
Usage : .venv/bin/python scripts/build_tables.py
"""
import re
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402
import build_potx_figma as bpf  # noqa: E402
import build_charts as bc  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND", ROOT / "bigtree")).resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT", ROOT / "build")).resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
e = bpf.e

DATA_TABLE = [["記事", "公開日", "PV", "スキ", "スキ率"],
              ["#1 はじめに：AI と一緒に note を書く", "09-02", "2,140", "96", "4.5%"],
              ["#2 挿絵は HTML で作る", "09-09", "1,820", "121", "6.6%"],
              ["#3 眠っていた才能", "09-16", "3,560", "214", "6.0%"],
              ["#4 Office テンプレート編（予定）", "09-30", "—", "—", "—"]]
COMPARE = [["", "AI に任せきり", "人が全部書く", "今回のやり方"],
           ["下書きの速さ", "速い", "遅い", "速い"],
           ["一次情報", "入らない", "入る", "入る"],
           ["見た目の統一", "毎回ぶれる", "手間がかかる", "トークンで自動"],
           ["読者への価値", "薄い", "高い", "高い"]]


def table(rows, x, y, w, col_w=None, row_h=56, first_col=True, band=True, body_align="r"):
    ncol = len(rows[0])
    col_w = col_w or [w / ncol] * ncol
    grid = "".join(f'<a:gridCol w="{e(cw)}"/>' for cw in col_w)
    cap = bpf.STYLES["caption"]
    trs = ""
    for ri, row in enumerate(rows):
        tcs = ""
        for ci, v in enumerate(row):
            algn = "l" if ci == 0 else body_align
            if ri == 0:
                algn = "l" if ci == 0 else "ctr"
            tcs += (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="{algn}">'
                    f'<a:lnSpc><a:spcPts val="{round(cap["pt"] * cap["line"])}"/></a:lnSpc></a:pPr>'
                    f'<a:r><a:rPr lang="ja-JP" sz="{cap["pt"] * 100}"/><a:t>{bpf.bp_esc(v)}</a:t></a:r></a:p></a:txBody>'
                    f'<a:tcPr marL="{e(12)}" marR="{e(12)}" marT="{e(6)}" marB="{e(6)}" anchor="ctr"/></a:tc>')
        trs += f'<a:tr h="{e(row_h)}">{tcs}</a:tr>'
    look = f'<a:tblPr firstRow="1" firstCol="{1 if first_col else 0}" bandRow="{1 if band else 0}"><a:tableStyleId>{bpf.TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>'
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{bp.nid()}" name="Table"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{e(x)}" y="{e(y)}"/><a:ext cx="{e(w)}" cy="{e(row_h * len(rows))}"/></p:xfrm>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl>{look}<a:tblGrid>{grid}</a:tblGrid>{trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')


def main():
    body = next(l for l in bpf.SPEC["layouts"] if l["name"] == "06_本文")
    it = {i["name"]: i for i in body["items"]}
    cx, cy, cw, ch = it["Content"]["box"]
    f = bc.read_zip(OUT / "bigtree_lab_organisms.pptx")
    ct, pres = f["[Content_Types].xml"].decode(), f["ppt/presentation.xml"].decode()
    prels = f["ppt/_rels/presentation.xml.rels"].decode()
    n = len([k for k in f if re.match(r"ppt/slides/slide\d+\.xml$", k)])
    layout_no = [l["name"] for l in bpf.SPEC["layouts"]].index("06_本文") + 1
    slides = [(DATA_TABLE, "記事別の PV とスキ（2026年9月）", "連載 #3 が PV・スキとも最も多い",
               [560, 140, 200, 180, 216], "r"),        # numbers: right aligned
              (COMPARE, "3つの書き方を比べる", "速さと一次情報は両立できる", None, "ctr")]  # words: centered
    for i, (rows, title, lead, col_w, body_align) in enumerate(slides, 1):
        bp._id[0] = 1
        tbl = table(rows, cx, cy + 16, cw, col_w, body_align=body_align)
        shapes = [bpf.placeholder({**it["Title"], "text": title}, on_slide=True),
                  bpf.placeholder({**it["Lead"], "text": lead}, on_slide=True),
                  tbl,
                  bpf.placeholder(it["SlideNumber"], on_slide=True)]
        sn = n + i
        f[f"ppt/slides/slide{sn}.xml"] = (bp.XML + f'<p:sld {bp.NS}><p:cSld>{bp.sptree(shapes)}</p:cSld>'
                                          '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>').encode()
        f[f"ppt/slides/_rels/slide{sn}.xml.rels"] = bpf.rels_xml([(bpf.R("slideLayout"), f"../slideLayouts/slideLayout{layout_no}.xml")]).encode()
        rid = f"rIdTbl{i}"
        prels = bc.add_rel(prels, rid, "slide", f"slides/slide{sn}.xml")
        pres = pres.replace("</p:sldIdLst>", f'<p:sldId id="{500 + i}" r:id="{rid}"/></p:sldIdLst>')
        ct = bc.add_override(ct, f"/ppt/slides/slide{sn}.xml", f"{bp.CT}.presentationml.slide+xml")
    f["[Content_Types].xml"], f["ppt/presentation.xml"], f["ppt/_rels/presentation.xml.rels"] = ct.encode(), pres.encode(), prels.encode()
    out = OUT / "bigtree_lab_tables.pptx"
    out.write_bytes(bc.zip_bytes(f))
    print(f"wrote {out.relative_to(OUT.parent)}")


if __name__ == "__main__":
    main()
