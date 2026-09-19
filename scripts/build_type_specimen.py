"""Type-scale specimen deck: the same 4 slides at 3 candidate sizes (body 12 / 14 / 16 pt).

Scale = body x 1.25^n for each role, rounded to whole pt. Values are candidates, not decisions.
Output: build/type_specimen.pptx
Usage : .venv/bin/python scripts/build_type_specimen.py
"""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402

RATIO = 1.25
ROLES = [  # (key, label, step from body)
    ("cover", "表紙タイトル", 4), ("section", "セクション見出し", 3), ("title", "スライドタイトル", 2),
    ("lead", "リード文", 1), ("body", "本文", 0), ("caption", "補足・グラフ・表", -1), ("min", "最小（ページ番号等）", -2),
]
CASES = [("A", "控えめ", 12), ("B", "標準", 14), ("C", "大きめ", 16)]


def scale(body):
    return {k: round(body * RATIO ** s) for k, _, s in ROLES}


def E(pt):
    return int(pt * 12700)


def run(text, size, color="tx1", bold=False, black=False):
    font = ('<a:latin typeface="Noto Sans JP Black"/><a:ea typeface="Noto Sans JP Black"/>' if black
            else '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/>')
    b = ' b="1"' if bold else ""
    return (f'<a:r><a:rPr lang="ja-JP" sz="{size * 100}"{b}>'
            + f'<a:solidFill><a:schemeClr val="{color}"/></a:solidFill>{font}</a:rPr><a:t>{bp_esc(text)}</a:t></a:r>')


def bp_esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def para(runs, algn="l", bullet=None, lvl=0, before=0, line=120):
    bu = (f'<a:buFont typeface="Arial"/><a:buChar char="{bullet}"/>' if bullet else '<a:buNone/>')
    ind = f' marL="{E(14 + lvl * 16)}" indent="{-E(12)}"' if bullet else ""
    return (f'<a:p><a:pPr algn="{algn}" lvl="{lvl}"{ind}><a:lnSpc><a:spcPct val="{line * 1000}"/></a:lnSpc>'
            f'<a:spcBef><a:spcPts val="{before * 100}"/></a:spcBef>{bu}</a:pPr>{runs}</a:p>')


def tb(name, x, y, w, h, paras, anchor="t", fill=None):
    f = f'<a:solidFill><a:schemeClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    ins = 'lIns="0" tIns="0" rIns="0" bIns="0"' if not fill else f'lIns="{E(6)}" tIns="{E(3)}" rIns="{E(6)}" bIns="{E(3)}"'
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr>{bp.xfrm(E(x), E(y), E(w), E(h))}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{f}</p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" {ins} anchor="{anchor}"><a:noAutofit/></a:bodyPr><a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')


def rect(name, x, y, w, h, color):
    return bp.rect(name, (E(x), E(y), E(w), E(h)), color)


def tag(case, s):
    k, label, body = case
    return tb("CaseTag", 520, 8, 192, 16, [para(run(f"案{k}（{label}）本文 {body}pt", 9, "dk2", bold=True), algn="r")])


def footer(s, n, color="dk2"):
    return [tb("Copyright", 36, 384, 300, 12, [para(run("© 2026 Bigtree Lab", s["min"], color))]),
            tb("SlideNumber", 600, 384, 84, 12, [para(run(str(n), s["min"], color), algn="r")])]


def slides_for(case, start_no):
    s = scale(case[2])
    out = []
    # 1 cover
    bp._id[0] = 1
    out.append((1, bp.slide([
        rect("Decor_Band", 0, 0, 14, 405, "accent1"), tag(case, s),
        tb("Title", 54, 130, 620, 110, [para(run("EC 運用レポート 2026年9月", s["cover"], black=True), line=110)], anchor="b"),
        rect("Decor_Line", 54, 250, 120, 3, "accent6"),
        tb("Subtitle", 54, 262, 620, 40, [para(run("広告費の配分見直しと、次の4週間の打ち手", s["lead"], "dk2"))]),
        tb("Meta", 54, 340, 620, 20, [para(run("2026-09-19 ｜ Bigtree Lab ｜ だいき君", s["caption"], "dk2"))]),
    ])))
    # 2 section (dark)
    bp._id[0] = 1
    out.append((2, bp.slide([
        tag(case, s),
        tb("Number", 72, 120, 300, 40, [para(run("01", s["lead"], "accent2", bold=True))]),
        tb("Title", 72, 160, 576, 90, [para(run("直近4週間の振り返り", s["section"], "bg1", black=True), line=110)]),
        rect("Decor_Line", 72, 262, 60, 3, "accent2"),
        *footer(s, start_no + 1, "bg1"),
    ])))
    # 3 body: text
    bp._id[0] = 1
    body_p = [para(run("広告経由のセッションは 4 週連続で増加し、W38 は 6,100 件だった", s["body"]), bullet="•", before=0, line=140),
              para(run("CVR は 2.1% → 2.4% に改善。ただし W38 はやや反落", s["body"]), bullet="•", before=6, line=140),
              para(run("楽天RPP の ROAS が最も高く、W38 は 510%", s["body"]), bullet="•", before=6, line=140),
              para(run("配分を 5% 移すと、月あたり約 12 万円の売上増を見込む（試算）", s["body"], "dk2"), bullet="◦", lvl=1, before=4, line=140),
              para(run("自然流入は安定。検索順位の変動は小さい", s["body"]), bullet="•", before=6, line=140)]
    out.append((1, bp.slide([
        tag(case, s),
        tb("Title", 36, 24, 648, 36, [para(run("広告経由の流入が伸び、CVR も改善した", s["title"], bold=True))], anchor="b"),
        rect("Decor_Line", 36, 64, 648, 1, "accent1"),
        tb("Lead", 36, 74, 648, 28, [para(run("楽天RPP への配分を増やすのが、次の4週間の最優先", s["lead"], "accent1", bold=True))]),
        tb("Body", 36, 112, 648, 230, body_p),
        tb("Note", 36, 352, 648, 18, [para(run("※ 数値はダミー。税抜・速報値。出典：BigQuery 週次集計", s["caption"], "dk2"))]),
        *footer(s, start_no + 2),
    ])))
    # 4 body: chart + table mock at caption size
    bp._id[0] = 1
    shapes = [tag(case, s),
              tb("Title", 36, 24, 648, 36, [para(run("週次セッションと ROAS", s["title"], bold=True))], anchor="b"),
              rect("Decor_Line", 36, 64, 648, 1, "accent1"),
              tb("Lead", 36, 74, 648, 28, [para(run("W38 は広告経由が過去最高", s["lead"], "accent1", bold=True))])]
    # bar chart mock (x 36..396, plot y 120..300)
    cats, ad, org = ["W35", "W36", "W37", "W38"], [5200, 5600, 5400, 6100], [7280, 7502, 7555, 8130]
    px, py, pw, ph, top = 76, 120, 320, 190, 9000
    for i, g in enumerate(range(0, 10000, 3000)):
        yy = py + ph - ph * g / top
        shapes.append(tb(f"YLabel{i}", 36, yy - 7, 36, 14, [para(run(f"{g:,}", s["caption"], "dk2"), algn="r")], anchor="ctr"))
        shapes.append(rect(f"Grid{i}", px, yy, pw, 0.5, "bg2"))
    gw = pw / 4
    for i, c in enumerate(cats):
        for j, (v, clr) in enumerate([(ad[i], "accent1"), (org[i], "accent2")]):
            h = ph * v / top
            shapes.append(rect(f"Bar{i}{j}", px + gw * i + gw * 0.15 + j * gw * 0.36, py + ph - h, gw * 0.33, h, clr))
        shapes.append(tb(f"XLabel{i}", px + gw * i, py + ph + 4, gw, 14, [para(run(c, s["caption"], "dk2"), algn="ctr")]))
    shapes.append(tb("Legend", px, py + ph + 24, pw, 14, [para(run("■ ", s["caption"], "accent1") + run("広告経由　", s["caption"])
                                                               + run("■ ", s["caption"], "accent2") + run("自然流入", s["caption"]), algn="ctr")]))
    # table mock (x 420..684)
    rows = [["チャネル", "ROAS"], ["Google", "395%"], ["Meta", "330%"], ["楽天RPP", "510%"]]
    rh = s["caption"] * 2.2
    for ri, r in enumerate(rows):
        fill = "accent1" if ri == 0 else ("bg2" if ri % 2 == 0 else None)
        clr = "bg1" if ri == 0 else "tx1"
        shapes.append(tb(f"Row{ri}", 420, 124 + ri * rh, 264, rh,
                         [para(run(r[0], s["caption"], clr, bold=ri == 0) + run("　" * 1, s["caption"], clr))], anchor="ctr", fill=fill))
        shapes.append(tb(f"Val{ri}", 560, 124 + ri * rh, 118, rh, [para(run(r[1], s["caption"], clr, bold=ri == 0), algn="r")], anchor="ctr"))
    shapes.append(tb("Caption", 420, 124 + 4 * rh + 6, 264, 30, [para(run("表1：チャネル別 ROAS（W38）", s["caption"], "dk2"))]))
    shapes += footer(s, start_no + 3)
    out.append((1, bp.slide(shapes)))
    return out


def overview():
    bp._id[0] = 1
    head = "".join(f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{para(run(t, 11, "bg1", bold=True), algn=("l" if i == 0 else "ctr"))}</a:txBody>'
                   f'<a:tcPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:tcPr></a:tc>'
                   for i, t in enumerate(["役割"] + [f"案{k}（{l}）" for k, l, _ in CASES]))
    body = ""
    for key, label, step in ROLES:
        cells = [label] + [f"{scale(b)[key]} pt" for _, _, b in CASES]
        body += '<a:tr h="304800">' + "".join(
            f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{para(run(t, 11, "tx1", bold=(key == "body")), algn=("l" if i == 0 else "ctr"))}</a:txBody><a:tcPr/></a:tc>'
            for i, t in enumerate(cells)) + '</a:tr>'
    tbl = (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{bp.nid()}" name="ScaleTable"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
           f'<p:xfrm><a:off x="{E(36)}" y="{E(100)}"/><a:ext cx="{E(648)}" cy="{E(250)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
           f'<a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="{E(216)}"/>' + "".join(f'<a:gridCol w="{E(144)}"/>' for _ in CASES)
           + f'</a:tblGrid><a:tr h="304800">{head}</a:tr>{body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>')
    return (1, bp.slide([
        tb("Title", 36, 24, 648, 36, [para(run("文字サイズ 3案の比較（試し書き）", 22, bold=True))], anchor="b"),
        rect("Decor_Line", 36, 64, 648, 1, "accent1"),
        tb("Lead", 36, 72, 648, 22, [para(run("どの案も「本文 × 1.25 のべき乗」。役割どうしの比率は同じで、全体の大きさだけが違う", 12, "dk2"))]),
        tbl]))


def main():
    bp.layouts = lambda: [bp.layout("00_Blank", [], [], show_master=False),
                          bp.layout("00_Blank_Dark", [], [], bg="tx1", show_master=False)]
    deck = [overview()]
    for case in CASES:
        deck += slides_for(case, len(deck) + 1)
    bp.slides = lambda: deck
    bp.build(bp.OUT / "type_specimen.pptx", template=False, embed=False)
    for k, l, b in CASES:
        print(k, l, scale(b))


if __name__ == "__main__":
    main()
