"""Native chart capability test: one chart definition -> companion xlsx + PPTX + DOCX.

Rule (REQUIREMENTS D3): the companion Excel (data1, graph1, data2, graph2 ...) is the source of truth;
charts embedded in PPTX / DOCX are copies regenerated from it.

Inputs : build/test_sample.pptx, build/bigtree_lab_sample.docx (from build_potx.py / build_dotx.py)
Outputs: build/charts/report_charts.xlsx, build/bigtree_lab_charts.pptx, build/bigtree_lab_charts.docx
Usage  : .venv/bin/python scripts/build_charts.py
"""
import io
import re
import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx  # noqa: E402  (shared theme)
import build_potx_figma as bpf  # noqa: E402  (レイアウトの並び順)

BODY_LAYOUT = [l["name"] for l in bpf.SPEC["layouts"]].index("06_本文") + 1

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND", ROOT / "bigtree")).resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT", ROOT / "build")).resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
C_NS = ('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "application/vnd.openxmlformats-officedocument"

# ---------------------------------------------------------------- data (dummy weekly report)
CHARTS = [  # fmt2 は複合グラフの第2軸の書式

    dict(name="graph1", data="data1", kind="bar", title="週次セッション（流入経路別）", fmt="#,##0",
         cats=["W35", "W36", "W37", "W38"],
         series=[("広告経由", [5200, 5600, 5400, 6100]), ("自然流入", [7280, 7502, 7555, 8130])],
         highlight=None),                      # automatic colors: accent1, accent2 ...
    dict(name="graph2", data="data2", kind="line", title="ROAS（チャネル別）", fmt="0%",
         cats=["W35", "W36", "W37", "W38"],
         series=[("Google", [3.88, 4.01, 4.12, 3.95]), ("Meta", [2.90, 3.10, 2.70, 3.30]), ("楽天RPP", [4.50, 4.20, 4.80, 5.10])],
         highlight="楽天RPP"),                  # one series in accent6 (emphasis), the rest muted
    dict(name="graph3", data="data3", kind="bar", title="週次セッション（W38 の広告経由だけ強調）", fmt="#,##0",
         cats=["W35", "W36", "W37", "W38"],
         series=[("広告経由", [5200, 5600, 5400, 6100]), ("自然流入", [7280, 7502, 7555, 8130])],
         highlight=None, pale=True, point=("広告経由", "W38")),  # everything pale, one data point solid + outlined
    dict(name="graph4", data="data4", kind="bar100", title="流入元の構成比（2026/09）", fmt="0%",
         cats=["2026/09"],
         series=[("検索", [0.58]), ("SNS", [0.21]), ("note 内", [0.13]), ("その他", [0.08])],
         highlight=None),   # 100% 積み上げ横棒。構成比は図形ではなくグラフで作る（データだから）
    dict(name="graph5", data="data5", kind="combo", title="セッションと CVR", fmt="#,##0", fmt2="0.0%",
         cats=["W35", "W36", "W37", "W38"],
         series=[("セッション（件）", [12480, 13102, 12955, 14230]), ("CVR（%）", [0.021, 0.023, 0.024, 0.022])],
         highlight="CVR（%）", unit2=0.005),   # 棒＝量（グレー）、折れ線＝率（強調色）。2軸
]

PALE = '<a:lumMod val="40000"/><a:lumOff val="60000"/>'  # theme tint, stays linked to the palette


def col(i):
    return "ABCDEFGHIJ"[i]


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


# ---------------------------------------------------------------- chart XML (shared by xlsx / pptx / docx)
TXT = ('<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="{sz}"><a:solidFill><a:schemeClr val="{clr}"/></a:solidFill>'
       '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr>')


def dlbls(kind, n, highlighted, fmt):
    """値ラベル（CHART_RULES.md §3）。棒は6カテゴリ以下なら全点、折れ線は両端だけ。"""
    txt = TXT.format(sz=1100, clr="tx1")
    if kind == "bar" and n <= 6:
        return (f'<c:dLbls>{txt}<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/>'
                '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    if kind == "bar100":   # 帯の中に % を置く
        return (f'<c:dLbls>{TXT.format(sz=1100, clr="tx1")}<c:dLblPos val="ctr"/><c:showLegendKey val="0"/><c:showVal val="1"/>'
                '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    if kind in ("line", "combo-line"):
        # 直接ラベル：最後の点に「系列名＋値」。強調系列は最初の点にも値を出す
        pts = [n - 1] + ([0] if highlighted else [])
        body = ""
        for i in sorted(pts):
            ser_name = '<c:showSerName val="1"/>' if i == n - 1 else '<c:showSerName val="0"/>'
            body += (f'<c:dLbl><c:idx val="{i}"/>{txt}<c:dLblPos val="r" />'
                     f'<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>{ser_name}'
                     '<c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>')
        return (f'<c:dLbls>{body}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>'
                '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>')
    return ""


def chart_xml(ch, sheet, embedded):
    ch.setdefault("fmt2", None)
    """CHART_RULES.md に沿ってグラフ XML を組み立てる。kind: bar / line / bar100 / combo"""
    n = len(ch["cats"])
    kind = ch["kind"]
    q = f"'{sheet}'" if not sheet.isascii() else sheet
    cat_ref = (f'<c:cat><c:strRef><c:f>{q}!$A$2:$A${n + 1}</c:f><c:strCache><c:ptCount val="{n}"/>'
               + "".join(f'<c:pt idx="{i}"><c:v>{esc(v)}</c:v></c:pt>' for i, v in enumerate(ch["cats"]))
               + '</c:strCache></c:strRef></c:cat>')
    series_xml = []
    for si, (sname, vals) in enumerate(ch["series"]):
        c = col(si + 1)
        line_like = kind == "line" or (kind == "combo" and si > 0)
        fmt = ch["fmt2"] if (kind == "combo" and si > 0) else ch["fmt"]
        dpt, sppr = "", ""
        if ch.get("pale"):          # 全体を淡く、1点だけ濃く（強調の単位＝1点）
            sppr = (f'<c:spPr><a:solidFill><a:schemeClr val="accent{si + 1}">{PALE}</a:schemeClr></a:solidFill>'
                    '<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>')
            if ch.get("point") and ch["point"][0] == sname:
                pi = ch["cats"].index(ch["point"][1])
                dpt = (f'<c:dPt><c:idx val="{pi}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>'
                       f'<c:spPr><a:solidFill><a:schemeClr val="accent{si + 1}"/></a:solidFill>'
                       '<a:ln w="28575"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:ln></c:spPr></c:dPt>')
        elif kind == "combo":       # 棒＝量はグレー、折れ線＝率は強調色
            sppr = ('<c:spPr><a:ln w="34925" cap="rnd"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill><a:round/></a:ln></c:spPr>'
                    if si else '<c:spPr><a:solidFill><a:schemeClr val="accent4"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>')
        elif ch.get("highlight"):   # 強調する系列だけ accent6、ほかはグレー
            clr, w = ("accent6", 34925) if sname == ch["highlight"] else ("accent4", 19050)
            sppr = (f'<c:spPr><a:ln w="{w}" cap="rnd"><a:solidFill><a:schemeClr val="{clr}"/></a:solidFill><a:round/></a:ln></c:spPr>')
        marker = '<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker>' if line_like else ""
        lbl = dlbls("combo-line" if line_like else kind, n, sname == ch.get("highlight"), fmt)
        series_xml.append(
            f'<c:ser><c:idx val="{si}"/><c:order val="{si}"/>'
            f'<c:tx><c:strRef><c:f>{q}!${c}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(sname)}</c:v></c:pt></c:strCache></c:strRef></c:tx>'
            f'{sppr}{marker}{dpt}{lbl}{cat_ref}'
            f'<c:val><c:numRef><c:f>{q}!${c}$2:${c}${n + 1}</c:f><c:numCache><c:formatCode>{fmt}</c:formatCode><c:ptCount val="{n}"/>'
            + "".join(f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>' for i, v in enumerate(vals))
            + '</c:numCache></c:numRef></c:val>'
            + ('<c:smooth val="0"/>' if line_like else "") + '</c:ser>')
    sers = "".join(series_xml)

    if kind == "combo":
        plot = (f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{series_xml[0]}'
                '<c:gapWidth val="80"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>'
                f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{series_xml[1]}'
                '<c:marker val="1"/><c:axId val="1003"/><c:axId val="1004"/></c:lineChart>')
    elif kind == "bar100":
        plot = (f'<c:barChart><c:barDir val="bar"/><c:grouping val="percentStacked"/><c:varyColors val="0"/>{sers}'
                '<c:gapWidth val="60"/><c:overlap val="100"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>')
    elif kind == "bar":
        plot = (f'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>{sers}'
                '<c:gapWidth val="80"/><c:overlap val="-10"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>')
    else:
        plot = (f'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{sers}'
                '<c:marker val="1"/><c:axId val="1001"/><c:axId val="1002"/></c:lineChart>')

    grid = ('<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:schemeClr val="accent5"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>')
    noline = '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
    horiz = kind == "bar100"
    # 値ラベルを出した棒グラフは、縦軸の目盛りを消す（軸かラベルのどちらか一方）
    val_deleted = 1 if (kind == "bar" and n <= 6) or horiz else 0
    def cat_ax(axid, crossax, delete=0):
        return ('<c:catAx><c:axId val="%s"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
                '<c:delete val="%s"/><c:axPos val="%s"/><c:numFmt formatCode="General" sourceLinked="1"/>'
                '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
                '<c:spPr><a:ln w="9525"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:ln></c:spPr>%s'
                '<c:crossAx val="%s"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
                % (axid, delete, "l" if horiz else "b", TXT.format(sz=1100, clr="tx2"), crossax))
    def val_ax(axid, crossax, fmt, delete=0, crosses="autoZero", show_grid=True, zero=False, unit=None):
        # 棒は必ず 0 起点（CHART_RULES.md §1）
        scaling = '<c:orientation val="minMax"/>' + ('<c:min val="0"/>' if zero else "")
        return ('<c:valAx><c:axId val="%s"/><c:scaling>%s</c:scaling><c:delete val="%s"/>'
                '<c:axPos val="%s"/>%s<c:numFmt formatCode="%s" sourceLinked="0"/><c:majorTickMark val="none"/>'
                '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>%s%s<c:crossAx val="%s"/><c:crosses val="%s"/>'
                '<c:crossBetween val="between"/>%s</c:valAx>'
                % (axid, scaling, delete, "b" if horiz else "l", grid if show_grid else "", fmt,
                   noline, TXT.format(sz=1100, clr="tx2"), crossax, crosses,
                   f'<c:majorUnit val="{unit}"/>' if unit else ""))
    if kind == "combo":
        axes = (cat_ax(1001, 1002) + val_ax(1002, 1001, ch["fmt"], zero=True)
                + val_ax(1004, 1003, ch["fmt2"], crosses="max", show_grid=False, unit=ch.get("unit2"))
                + cat_ax(1003, 1004, delete=1))
    else:
        axes = cat_ax(1001, 1002) + val_ax(1002, 1001, ch["fmt"], delete=val_deleted, zero=(kind == "bar"))
    # 直接ラベルを置く折れ線・複合では凡例を出さない（CHART_RULES.md §5）
    legend = ("" if kind in ("line", "combo")
              else '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + TXT.format(sz=1100, clr="tx1") + '</c:legend>')
    ext = '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>' if embedded else ""
    return (XML + f'<c:chartSpace {C_NS}><c:date1904 val="0"/><c:lang val="ja-JP"/><c:roundedCorners val="0"/>'
            f'<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>{plot}{axes}</c:plotArea>{legend}'
            '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
            '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' + TXT.format(sz=1100, clr="tx1") + f'{ext}</c:chartSpace>')


# ---------------------------------------------------------------- xlsx writer
def sheet_xml(ch):
    fmt_style = 2 if "%" in ch["fmt"] else 1
    rows = [[""] + [s for s, _ in ch["series"]]] + [[c] + [v[i] for _, v in ch["series"]] for i, c in enumerate(ch["cats"])]
    out = ""
    for ri, row in enumerate(rows, 1):
        cells = ""
        for ci, v in enumerate(row):
            ref = f"{col(ci)}{ri}"
            if isinstance(v, (int, float)):
                cells += f'<c r="{ref}" s="{fmt_style}"><v>{v}</v></c>'
            elif v != "":
                style = ' s="3"' if ri == 1 else ""
                cells += f'<c r="{ref}" t="inlineStr"{style}><is><t>{esc(v)}</t></is></c>'
        out += f'<row r="{ri}">{cells}</row>'
    return (XML + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="4" width="14" customWidth="1"/></cols>'
            f'<sheetData>{out}</sheetData></worksheet>')


STYLES = (XML + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Noto Sans JP"/><family val="2"/><scheme val="minor"/></font>'
          '<font><b/><sz val="11"/><color theme="1"/><name val="Noto Sans JP"/><family val="2"/><scheme val="minor"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
          '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
          '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


def xlsx(order):
    """order: [(kind, name, xml)] with kind 'sheet' (worksheet xml) or 'chart' (chart xml -> chartsheet), kept in this order."""
    files, ct, wb_rels, sheets_xml = {}, [], [], ""
    ws_i = cs_i = 0
    for i, (kind, name, body) in enumerate(order, 1):
        if kind == "sheet":
            ws_i += 1
            files[f"xl/worksheets/sheet{ws_i}.xml"] = body
            wb_rels.append(("worksheet", f"worksheets/sheet{ws_i}.xml"))
            ct.append((f"/xl/worksheets/sheet{ws_i}.xml", f"{CT}.spreadsheetml.worksheet+xml"))
        else:
            cs_i += 1
            files[f"xl/chartsheets/sheet{cs_i}.xml"] = (
                XML + '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                f'xmlns:r="{REL}"><sheetViews><sheetView zoomToFit="1" workbookViewId="0"/></sheetViews><drawing r:id="rId1"/></chartsheet>')
            files[f"xl/chartsheets/_rels/sheet{cs_i}.xml.rels"] = rels([("drawing", f"../drawings/drawing{cs_i}.xml")])
            files[f"xl/drawings/drawing{cs_i}.xml"] = (
                XML + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:absoluteAnchor><xdr:pos x="0" y="0"/>'
                '<xdr:ext cx="8670000" cy="6300000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>'
                f'<xdr:cNvPr id="2" name="{name}"/><xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr></xdr:nvGraphicFramePr>'
                '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="{REL}" r:id="rId1"/></a:graphicData></a:graphic>'
                '</xdr:graphicFrame><xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>')
            files[f"xl/drawings/_rels/drawing{cs_i}.xml.rels"] = rels([("chart", f"../charts/chart{cs_i}.xml")])
            files[f"xl/charts/chart{cs_i}.xml"] = body
            wb_rels.append(("chartsheet", f"chartsheets/sheet{cs_i}.xml"))
            ct += [(f"/xl/chartsheets/sheet{cs_i}.xml", f"{CT}.spreadsheetml.chartsheet+xml"),
                   (f"/xl/drawings/drawing{cs_i}.xml", f"{CT}.drawing+xml"),
                   (f"/xl/charts/chart{cs_i}.xml", f"{CT}.drawingml.chart+xml")]
        sheets_xml += f'<sheet name="{esc(name)}" sheetId="{i}" r:id="rId{i}"/>'
    n = len(wb_rels)
    wb_rels += [("styles", "styles.xml"), ("theme", "theme/theme1.xml")]
    files["xl/workbook.xml"] = (XML + f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{REL}">'
                                f'<bookViews><workbookView/></bookViews><sheets>{sheets_xml}</sheets></workbook>')
    files["xl/_rels/workbook.xml.rels"] = rels(wb_rels)
    files["xl/styles.xml"] = STYLES
    files["xl/theme/theme1.xml"] = build_potx.theme()
    files["_rels/.rels"] = rels([("officeDocument", "xl/workbook.xml")])
    ct += [("/xl/workbook.xml", f"{CT}.spreadsheetml.sheet.main+xml"), ("/xl/styles.xml", f"{CT}.spreadsheetml.styles+xml"),
           ("/xl/theme/theme1.xml", f"{CT}.theme+xml")]
    files["[Content_Types].xml"] = content_types(ct)
    return zip_bytes(files)


# ---------------------------------------------------------------- package helpers
def rels(items, start=1):
    return (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i + start}" Type="{REL}/{t}" Target="{tg}"/>' for i, (t, tg) in enumerate(items))
            + '</Relationships>')


def content_types(overrides, extra_defaults=""):
    return (XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>' + extra_defaults
            + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides) + '</Types>')


def zip_bytes(files):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for k in ["[Content_Types].xml"] + [k for k in files if k != "[Content_Types].xml"]:
            v = files[k]
            z.writestr(k, v if isinstance(v, bytes) else v.encode("utf-8"))
    return buf.getvalue()


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def add_rel(rels_xml, rid, rtype, target, external=False):
    mode = ' TargetMode="External"' if external else ""
    return rels_xml.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{REL}/{rtype}" Target="{target}"{mode}/></Relationships>')


def add_override(ct_xml, part, ctype):
    return ct_xml.replace("</Types>", f'<Override PartName="{part}" ContentType="{ctype}"/></Types>')


def add_default(ct_xml, ext, ctype):
    if f'Extension="{ext}"' in ct_xml:
        return ct_xml
    return ct_xml.replace("<Override", f'<Default Extension="{ext}" ContentType="{ctype}"/><Override', 1)


def embedded_parts(files, prefix, ch, i):
    """Write chart{i}.xml + its embedded workbook under {prefix}/charts and {prefix}/embeddings."""
    files[f"{prefix}/charts/chart{i}.xml"] = chart_xml(ch, "Sheet1", embedded=True).encode()
    files[f"{prefix}/charts/_rels/chart{i}.xml.rels"] = (
        XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{REL}/package" Target="../embeddings/Microsoft_Excel_Worksheet{i}.xlsx"/></Relationships>').encode()
    files[f"{prefix}/embeddings/Microsoft_Excel_Worksheet{i}.xlsx"] = xlsx([("sheet", "Sheet1", sheet_xml(ch))])


# ---------------------------------------------------------------- PPTX
def inject_pptx(src, dst):
    f = read_zip(src)
    ct = f["[Content_Types].xml"].decode()
    pres = f["ppt/presentation.xml"].decode()
    prels = f["ppt/_rels/presentation.xml.rels"].decode()
    n_slides = len([k for k in f if re.match(r"ppt/slides/slide\d+\.xml$", k)])
    W, H, M = build_potx.W, build_potx.H, build_potx.M
    for i, ch in enumerate(CHARTS, 1):
        embedded_parts(f, "ppt", ch, i)
        sn = n_slides + i
        build_potx._id[0] = 1
        frame = (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{build_potx.nid()}" name="{ch["name"]}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
                 f'<p:xfrm><a:off x="{M}" y="1828800"/><a:ext cx="{W - 2 * M}" cy="{H - 1828800 - M}"/></p:xfrm>'
                 '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                 f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>')
        slide = build_potx.slide([
            build_potx.s_ph("Title", 'type="title"', [(0, ch["title"])]),
            build_potx.s_ph("Lead", 'type="body" sz="quarter" idx="13"', [(0, f"データは {ch['data']}（横の report_charts.xlsx が正）")]),
            frame, build_potx.s_num()])
        f[f"ppt/slides/slide{sn}.xml"] = slide.encode()
        f[f"ppt/slides/_rels/slide{sn}.xml.rels"] = (
            XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}/slideLayout" Target="../slideLayouts/slideLayout{BODY_LAYOUT}.xml"/>'
            f'<Relationship Id="rId2" Type="{REL}/chart" Target="../charts/chart{i}.xml"/></Relationships>').encode()
        rid = f"rIdChartSlide{i}"
        prels = add_rel(prels, rid, "slide", f"slides/slide{sn}.xml")
        pres = pres.replace("</p:sldIdLst>", f'<p:sldId id="{300 + i}" r:id="{rid}"/></p:sldIdLst>')
        ct = add_override(ct, f"/ppt/slides/slide{sn}.xml", f"{CT}.presentationml.slide+xml")
        ct = add_override(ct, f"/ppt/charts/chart{i}.xml", f"{CT}.drawingml.chart+xml")
    ct = add_default(ct, "xlsx", f"{CT}.spreadsheetml.sheet")
    f["[Content_Types].xml"], f["ppt/presentation.xml"], f["ppt/_rels/presentation.xml.rels"] = ct.encode(), pres.encode(), prels.encode()
    dst.write_bytes(zip_bytes(f))
    print(f"wrote {dst.relative_to(OUT.parent)}")


# ---------------------------------------------------------------- DOCX
def inject_docx(src, dst):
    f = read_zip(src)
    ct = f["[Content_Types].xml"].decode()
    doc = f["word/document.xml"].decode()
    drels = f["word/_rels/document.xml.rels"].decode()
    text_w_emu = (11906 - 2 * 1134) * 635
    cx, cy = int(text_w_emu * 0.9), int(text_w_emu * 0.9 * 0.55)
    body = ('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>グラフ</w:t></w:r></w:p>'
            '<w:p><w:r><w:t xml:space="preserve">ネイティブグラフ。右クリック →「データの編集」で埋め込みの Excel が開く。正は横に置いた report_charts.xlsx。</w:t></w:r></w:p>')
    for i, ch in enumerate(CHARTS, 1):
        embedded_parts(f, "word", ch, i)
        rid = f"rIdChart{i}"
        drels = add_rel(drels, rid, "chart", f"charts/chart{i}.xml")
        ct = add_override(ct, f"/word/charts/chart{i}.xml", f"{CT}.drawingml.chart+xml")
        body += (f'<w:p><w:pPr><w:pStyle w:val="FigureBlock"/></w:pPr><w:r><w:drawing>'
                 '<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
                 f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{100 + i}" name="{ch["name"]}"/>'
                 '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                 f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{rid}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
                 '<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t xml:space="preserve">図 </w:t></w:r>'
                 '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ 図 \\* ARABIC </w:instrText></w:r>'
                 f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{i + 1}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
                 f'<w:r><w:t xml:space="preserve">：{esc(ch["title"])}（{ch["name"]}）</w:t></w:r></w:p>')
    doc = doc.replace("<w:sectPr>", body + "<w:sectPr>", 1)
    ct = add_default(ct, "xlsx", f"{CT}.spreadsheetml.sheet")
    f["[Content_Types].xml"], f["word/document.xml"], f["word/_rels/document.xml.rels"] = ct.encode(), doc.encode(), drels.encode()
    dst.write_bytes(zip_bytes(f))
    print(f"wrote {dst.relative_to(OUT.parent)}")


# ---------------------------------------------------------------- main
if __name__ == "__main__":
    (OUT / "charts").mkdir(parents=True, exist_ok=True)
    order = []  # data1, graph1, data2, graph2 ...
    for ch in CHARTS:
        order += [("sheet", ch["data"], sheet_xml(ch)), ("chart", ch["name"], chart_xml(ch, ch["data"], embedded=False))]
    (OUT / "charts/report_charts.xlsx").write_bytes(xlsx(order))
    print("wrote build/charts/report_charts.xlsx")
    inject_pptx(OUT / "bigtree_lab_sample.pptx", OUT / "bigtree_lab_charts.pptx")
    inject_docx(OUT / "bigtree_lab_sample.docx", OUT / "bigtree_lab_charts.docx")
