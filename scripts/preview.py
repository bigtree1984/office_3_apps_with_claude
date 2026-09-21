"""レイアウトと図解を1枚の HTML に描く簡易プレビュー（Figma が無くても検討できるように）。

**目的は「見る」ことではなく「指させる」こと。**
エージェントが作ったものを人が見て、「そこの 2px」「この文字が長い」と**要素の名前で**返せるようにする。
名前と座標が分からないと、フィードバックが「なんか違う」で終わってしまう。

できること
  - `design/figma_layouts.json` のレイアウトと `design/figma_organisms.json` の図解を、実寸の比率で描く
  - 要素にカーソルを乗せると**名前と座標**が出る。クリックすると、そのまま貼れる形でコピーされる
  - ガイド（余白・セーフゾーン）の表示切り替え
  - **文字のはみ出しを警告**（`fit_text.py` と同じ計算）。レイアウトのプレースホルダは折り返しオフなので
    「幅を超えたら警告」、図解の中の文字は折り返すので「枠の行数を超えたら警告」

できないこと（意図的に小さく作ってあります）
  - Office の実際の描画とは違います。行間の癖・表の罫線・グラフは **PPTX にして確認**してください
  - 画像のマスク（曲線の切り抜き）は近似です
  - **足りない機能は、この一式を使う人がエージェントに足させる前提**です。ここは出発点です

使い方：.venv/bin/python scripts/preview.py [--open]
"""
import base64
import json
import os
import sys
import webbrowser
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp       # noqa: E402  (tokens)
import fit_text as ft         # noqa: E402  (はみ出しの判定)

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND", ROOT / "bigtree")).resolve()
OUT = Path(os.environ.get("OFFICE3_OUT", ROOT / "build")).resolve()
OUT.mkdir(parents=True, exist_ok=True)

TOKENS = bp._TOKENS
COLORS = TOKENS["colors"]
ROLES = TOKENS["type_pptx"]["roles"]
PX = 2                                    # 2px = 1pt（Figma 1440px フレーム）
STYLE_KEY = {"本文": "body", "リード文": "lead", "補足・グラフ・表": "caption", "最小（ページ番号等）": "min",
             "スライドタイトル": "title", "セクション見出し": "section", "表紙タイトル": "cover"}
ALIGN = {"LEFT": "left", "CENTER": "center", "RIGHT": "right", "l": "left", "ctr": "center", "r": "right"}


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def hexc(name, opacity=1):
    h = COLORS.get(name, "888888")
    return f"#{h}" if opacity >= 1 else f"#{h}{int(opacity * 255):02x}"


def font_css(role):
    r = ROLES[role]
    return (f'font-size:{r["pt"] * PX}px;font-weight:{700 if r["weight"] == "Bold" else 400};'
            f'line-height:{r["line"]}%;')


def data_uri(path):
    if not path.exists():
        return ""
    mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "svg": "image/svg+xml"}
    return f'data:{mime.get(path.suffix.lstrip(".").lower(), "application/octet-stream")};base64,' \
           + base64.b64encode(path.read_bytes()).decode()


def el(name, x, y, w, h, inner="", css="", warn=""):
    """1つの要素。名前と座標を data 属性で持たせ、カーソルとクリックで拾えるようにする。"""
    return (f'<div class="el{" warn" if warn else ""}" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;{css}" '
            f'data-name="{esc(name)}" data-box="{x:.0f}, {y:.0f}, {w:.0f}, {h:.0f}"'
            + (f' data-warn="{esc(warn)}"' if warn else "") + f'>{inner}</div>')


def text_el(name, box, text, role, color, align="left", opacity=1, wrap=False):
    """wrap=True（図解の中）は枠の高さまで折り返せる。False（レイアウトのプレースホルダ）は1行で溢れる。"""
    x, y, w, h = box
    if wrap:
        ok, need, avail = ft.fits_box(text, role, w, h)
        warn = "" if ok else f"{need}行必要（枠は{avail}行）"
    else:
        ok, over, ratio = ft.fits(text, role, w)
        warn = "" if ok else f"はみ出し {over:.0f}pt（枠の {ratio * 100:.0f}%）"
    inner = f'<span class="t">{esc(text).replace(chr(10), "<br>")}</span>'
    css = (f'color:{hexc(color, opacity)};{font_css(role)}text-align:{ALIGN.get(align, "left")};'
           'display:flex;align-items:center;' + ("" if wrap else "white-space:pre;"))
    return el(name, x, y, w, h, inner, css, warn)


def layout_html(lay, assets):
    body = ""
    for it in lay["items"]:
        kind, name = it.get("kind"), it["name"]
        if kind == "picture":
            x, y, w, h = it["fill_box"]
            src = assets.get(it["image"], "")
            body += el(name, x, y, w, h,
                       f'<img src="{src}" style="width:100%;height:100%;object-fit:cover">',
                       "overflow:hidden;")
        elif kind == "logo":
            x, y, w, h = it["box"]
            src = assets.get("logo", "")
            body += el(name, x, y, w, h, f'<img src="{src}" style="width:100%;height:100%;object-fit:contain">')
        elif kind == "rect":
            x, y, w, h = it["box"]
            body += el(name, x, y, w, h, "", f'background:{hexc(it.get("color", "dk1"), it.get("opacity", 1))};')
        elif kind in ("ph", "text"):
            if it.get("ph") == "sldNum":
                txt = "3"
            else:
                txt = it.get("text", "")
            if not txt:
                x, y, w, h = it["box"]
                body += el(name, x, y, w, h, '<span class="empty">（中身はスライドで入る）</span>',
                           "border:1px dashed rgba(0,0,0,.2);color:#999;font-size:20px;")
                continue
            e = text_el(name, it["box"], txt, it.get("style", "body"), it.get("color", "dk1"), it.get("align", "l"))
            if it.get("pill"):
                e = e.replace('class="el', 'class="el pill', 1)
            body += e
    guides = ""
    for kind, v in (lay.get("guides", []) + TOKENS.get("master_guides", [])):
        guides += (f'<div class="guide" style="left:{v}px;top:0;width:1px;height:100%"></div>' if kind == "X"
                   else f'<div class="guide" style="left:0;top:{v}px;width:100%;height:1px"></div>')
    return body, guides


def organism_html(org):
    body = ""
    for it in org["items"]:
        x, y, w, h = it["x"], it["y"], it["w"], it["h"]
        t, name, fill, op = it["type"], it["name"], it.get("fill", "dk1"), it.get("fillOpacity", 1)
        if t == "TEXT":
            body += text_el(name, [x, y, w, h], it.get("text", ""), STYLE_KEY.get(it.get("style"), "body"),
                            fill, it.get("align", "LEFT"), op, wrap=True)
        elif t == "ELLIPSE":
            body += el(name, x, y, w, h, "", f'background:{hexc(fill, op)};border-radius:50%;')
        elif t in ("VECTOR", "POLYGON") and it.get("path"):
            # パスはフレーム座標なので、そのまま SVG に流し込む
            stroke = (f'stroke="{hexc(it["stroke"])}" stroke-width="{it.get("strokeW", 1)}"'
                      if it.get("stroke") else "")
            fill_attr = "none" if it.get("open") else hexc(fill, op)
            body += (f'<svg class="el vec" viewBox="0 0 {org["w"]} {org["h"]}" '
                     f'style="left:0;top:0;width:{org["w"]}px;height:{org["h"]}px" '
                     f'data-name="{esc(name)}" data-box="{x:.0f}, {y:.0f}, {w:.0f}, {h:.0f}">'
                     f'<path d="{esc(it["path"])}" fill="{fill_attr}" {stroke}/></svg>')
        else:
            body += el(name, x, y, w, h, "", f'background:{hexc(fill, op)};')
    return body


def main():
    spec = json.loads((BRAND / "design/figma_layouts.json").read_text())
    orgs = json.loads((BRAND / "design/figma_organisms.json").read_text())["organisms"]
    fw, fh = spec["frame"]
    TOKENS["master_guides"] = spec.get("master_guides", [])
    assets = {"logo": data_uri(BRAND / "assets/logo/260920_BT_logo.svg")}
    for lay in spec["layouts"]:
        for it in lay["items"]:
            if it.get("kind") == "picture":
                assets[it["image"]] = data_uri(BRAND / it["image"])

    cards = ""
    for lay in spec["layouts"]:
        body, guides = layout_html(lay, assets)
        cards += (f'<section class="card" data-kind="layout"><h2>{esc(lay["name"])}'
                  f'<small>レイアウト ｜ {fw}×{fh}px</small></h2>'
                  f'<div class="frame" style="width:{fw}px;height:{fh}px;background:{hexc(lay["bg"])}">'
                  f'{body}<div class="guides">{guides}</div></div></section>')
    for name, org in orgs.items():
        cards += (f'<section class="card" data-kind="organism"><h2>{esc(name)}'
                  f'<small>図解 ｜ {org["w"]}×{org["h"]}px</small></h2>'
                  f'<div class="frame" style="width:{org["w"]}px;height:{org["h"]}px;background:{hexc("lt1")}">'
                  f'{organism_html(org)}</div></section>')

    html = TEMPLATE.replace("{{CARDS}}", cards).replace("{{FONT}}", TOKENS["font"]["family"])
    path = OUT / "preview.html"
    path.write_text(html, encoding="utf-8")
    print(f"wrote {path.relative_to(OUT.parent)}  （ブラウザで開いてください）")
    if "--open" in sys.argv:
        webbrowser.open(path.as_uri())


TEMPLATE = """<!DOCTYPE html>
<html lang="ja"><head><meta charset="utf-8"><title>プレビュー</title>
<style>
 body{margin:0;background:#eceae5;font-family:"{{FONT}}",system-ui,sans-serif;color:#222}
 header{position:sticky;top:0;z-index:50;background:#fff;border-bottom:1px solid #ddd;padding:10px 20px;
        display:flex;gap:18px;align-items:center;flex-wrap:wrap}
 header b{font-size:15px} label{font-size:13px;user-select:none;cursor:pointer}
 #tip{margin-left:auto;font-size:13px;color:#555;font-family:ui-monospace,monospace}
 main{padding:24px;display:flex;flex-direction:column;gap:28px;align-items:flex-start}
 .card h2{font-size:14px;margin:0 0 8px;font-weight:700}
 .card h2 small{font-weight:400;color:#777;margin-left:10px}
 .frame{position:relative;transform-origin:top left;box-shadow:0 2px 12px rgba(0,0,0,.12);overflow:hidden}
 .el{position:absolute;box-sizing:border-box}
 .el .t{display:block;width:100%}
 .el.pill{border:2px solid currentColor;border-radius:999px;justify-content:center}
 .empty{font-size:20px}
 svg.el{position:absolute;overflow:visible}
 .guides{position:absolute;inset:0;pointer-events:none;display:none}
 .guide{position:absolute;background:rgba(200,60,60,.55)}
 body.guides-on .guides{display:block}
 body.outline-on .el{outline:1px solid rgba(0,120,255,.45)}
 .el:hover{outline:2px solid #0078ff !important}
 .el.warn{outline:2px solid #e2773a}
 .el.warn::after{content:attr(data-warn);position:absolute;right:0;top:100%;font-size:18px;color:#e2773a;
                 font-weight:700;white-space:nowrap}
 body.warn-off .el.warn{outline:none} body.warn-off .el.warn::after{display:none}
</style></head><body class="guides-on warn-on">
<header>
 <b>プレビュー</b>
 <label><input type="checkbox" id="g" checked> ガイド</label>
 <label><input type="checkbox" id="o"> 要素の枠</label>
 <label><input type="checkbox" id="w" checked> はみ出し警告</label>
 <label>表示倍率 <input type="range" id="z" min="30" max="100" value="62"> <span id="zv">62%</span></label>
 <span id="tip">要素にカーソルを乗せると名前と座標。クリックでコピー（エージェントにそのまま貼れます）</span>
</header>
<main>{{CARDS}}</main>
<script>
 const tip = document.getElementById('tip');
 function zoom(v){
   document.querySelectorAll('.frame').forEach(f=>{
     f.style.transform = 'scale('+v+')';
     f.style.marginBottom = (f.offsetHeight*(v-1))+'px';
     f.style.marginRight = (f.offsetWidth*(v-1))+'px';
   });
   document.getElementById('zv').textContent = Math.round(v*100)+'%';
 }
 document.getElementById('z').oninput = e => zoom(e.target.value/100);
 document.getElementById('g').onchange = e => document.body.classList.toggle('guides-on', e.target.checked);
 document.getElementById('o').onchange = e => document.body.classList.toggle('outline-on', e.target.checked);
 document.getElementById('w').onchange = e => document.body.classList.toggle('warn-off', !e.target.checked);
 document.addEventListener('mouseover', e => {
   const el = e.target.closest('[data-name]'); if(!el) return;
   const card = el.closest('.card').querySelector('h2').firstChild.textContent;
   tip.textContent = card + ' / ' + el.dataset.name + '  [' + el.dataset.box + ']'
     + (el.dataset.warn ? '  ⚠ ' + el.dataset.warn : '');
 });
 document.addEventListener('click', e => {
   const el = e.target.closest('[data-name]'); if(!el) return;
   const card = el.closest('.card').querySelector('h2').firstChild.textContent;
   const s = card + ' の ' + el.dataset.name + '（x,y,w,h = ' + el.dataset.box + '）';
   navigator.clipboard.writeText(s).then(()=>{ tip.textContent = 'コピーしました → ' + s; });
 });
 zoom(0.62);
</script></body></html>
"""

if __name__ == "__main__":
    main()
