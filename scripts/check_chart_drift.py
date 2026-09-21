"""Compare charts embedded in a PPTX / DOCX with the companion Excel (source of truth).

Charts are matched by name: a chart frame named "graphN" in the document <-> sheet "dataN" in the Excel.
Both the chart's value cache (what is drawn) and its embedded workbook are checked.
Usage: .venv/bin/python scripts/check_chart_drift.py <companion.xlsx> <doc.pptx|doc.docx> [...]
"""
import io
import posixpath
import re
import os
import sys
import zipfile


def read_sheet_values(xlsx_bytes, sheet_name=None):
    """Return {cell_ref: value} for one sheet (first sheet if name is None). Handles inline and shared strings."""
    z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    sheets = re.findall(r'<sheet [^>]*name="([^"]*)"[^>]*r:id="([^"]*)"', wb)
    name, rid = next((s for s in sheets if sheet_name is None or s[0] == sheet_name), (None, None))
    if rid is None:
        return None
    target = re.search(rf'Id="{rid}"[^>]*Target="([^"]*)"', rels) or re.search(rf'Target="([^"]*)"[^>]*Id="{rid}"', rels)
    path = "xl/" + target.group(1).lstrip("/").replace("xl/", "")
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = [re.sub(r"<[^>]+>", "", si) for si in re.findall(r"<si>(.*?)</si>", z.read("xl/sharedStrings.xml").decode(), re.S)]
    out = {}
    for ref, attrs, body in re.findall(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', z.read(path).decode(), re.S):
        if 't="s"' in attrs:
            out[ref] = shared[int(re.search(r"<v>(.*?)</v>", body).group(1))]
        elif 't="inlineStr"' in attrs:
            out[ref] = re.sub(r"<[^>]+>", "", body)
        else:
            m = re.search(r"<v>(.*?)</v>", body)
            out[ref] = float(m.group(1)) if m else None
    return out


def chart_cache(chart_xml):
    """Return {cell_ref: value} reconstructed from the chart's own caches (what the chart actually draws)."""
    out = {}
    for f, cache in re.findall(r"<c:f>([^<]*)</c:f>\s*<c:(?:str|num)Cache>(.*?)</c:(?:str|num)Cache>", chart_xml, re.S):
        m = re.search(r"\$([A-Z]+)\$(\d+)(?::\$([A-Z]+)\$(\d+))?", f)
        colname, r0 = m.group(1), int(m.group(2))
        vertical = m.group(3) is None or m.group(3) == colname
        for idx, v in re.findall(r'<c:pt idx="(\d+)">\s*<c:v>([^<]*)</c:v>', cache):
            i = int(idx)
            ref = f"{colname}{r0 + i}" if vertical else f"{chr(ord(colname) + i)}{r0}"
            try:
                out[ref] = float(v)
            except ValueError:
                out[ref] = v
    return out


def doc_charts(path):
    """Yield (chart_name, chart_xml, embedded_xlsx_bytes) for every chart in a PPTX / DOCX."""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    hosts = [n for n in names if re.match(r"(ppt/slides/slide\d+|word/document)\.xml$", n)]
    for host in hosts:
        xml = z.read(host).decode()
        rels_path = posixpath.join(posixpath.dirname(host), "_rels", posixpath.basename(host) + ".rels")
        rels = z.read(rels_path).decode()
        # frame name ... c:chart r:id  (pptx: cNvPr name, docx: docPr name)
        frames = re.findall(r"<p:graphicFrame>.*?</p:graphicFrame>|<w:drawing>.*?</w:drawing>", xml, re.S)
        for frame in frames:
            chart = re.search(r'<c:chart [^>]*r:id="([^"]+)"', frame)
            if not chart:
                continue
            name, rid = re.search(r'(?:cNvPr|docPr) id="\d+" name="([^"]*)"', frame).group(1), chart.group(1)
            tgt = re.search(rf'Id="{rid}"[^>]*Target="([^"]*)"', rels).group(1)
            chart_path = posixpath.normpath(posixpath.join(posixpath.dirname(host), tgt))
            chart_xml = z.read(chart_path).decode()
            crels = z.read(posixpath.join(posixpath.dirname(chart_path), "_rels", posixpath.basename(chart_path) + ".rels")).decode()
            emb = re.search(r'Target="([^"]*\.xlsx)"', crels)
            emb_bytes = z.read(posixpath.normpath(posixpath.join(posixpath.dirname(chart_path), emb.group(1)))) if emb else None
            yield name, chart_xml, emb_bytes


def diff(label, truth, other):
    bad = []
    for ref, v in sorted(truth.items()):
        o = other.get(ref)
        same = (abs(v - o) < 1e-9) if isinstance(v, float) and isinstance(o, float) else v == o
        if not same:
            bad.append(f"    {label} {ref}: Excel={v!r}  資料内={o!r}")
    return bad


def main(xlsx_path, docs):
    xbytes = open(xlsx_path, "rb").read()
    ok = True
    for doc in docs:
        print(f"== {doc}")
        for name, chart_xml, emb in doc_charts(doc):
            m = re.match(r"graph(\d+)$", name)
            if not m:
                print(f"  {name}: 対応する dataN が名前から分からないためスキップ")
                continue
            truth = read_sheet_values(xbytes, f"data{m.group(1)}")
            if truth is None:
                print(f"  {name}: Excel に data{m.group(1)} シートが無い")
                ok = False
                continue
            problems = diff("表示中の値", truth, chart_cache(chart_xml))
            if emb:
                problems += diff("埋め込みExcel", truth, read_sheet_values(emb))
            if problems:
                ok = False
                print(f"  {name} ⚠ Excel（data{m.group(1)}）とずれています")
                print("\n".join(problems))
            else:
                print(f"  {name} ✓ Excel（data{m.group(1)}）と一致")
    return 0 if ok else 1


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(
            "使い方: .venv/bin/python scripts/check_chart_drift.py <正となる Excel> <確認する pptx / docx> [...]\n"
            "  例: .venv/bin/python scripts/check_chart_drift.py build/charts/report_charts.xlsx build/bigtree_lab_charts.pptx\n"
            "  グラフ枠の名前（graphN）とシート名（dataN）を突き合わせ、値のずれを報告します。")
    sys.exit(main(sys.argv[1], sys.argv[2:]))
