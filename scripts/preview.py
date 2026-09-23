"""レイアウトと図解を1枚の HTML に描く簡易プレビュー（Figma が無くても検討できるように）。

**目的は「見る」ことではなく「指させる」こと。**
エージェントが作ったものを人が見て、「そこの 2px」「この文字が長い」と**要素の名前で**返せるようにする。
名前と座標が分からないと、フィードバックが「なんか違う」で終わってしまう。

できること
  - `design/figma_layouts.json` のレイアウトと `design/figma_organisms.json` の図解を、実寸の比率で描く
  - 要素にカーソルを乗せると**名前と座標**が出る。クリックすると、そのまま貼れる形でコピーされる
  - ガイド（余白・セーフゾーン）の表示切り替え
  - **地色に対して文字が読めるかを警告**（WCAG のコントラスト比。**直すのは人**だが、危ない場所は指させる）
  - **文字のはみ出しを警告**（`fit_text.py` と同じ計算）。レイアウトのプレースホルダは折り返しオフなので
    「幅を超えたら警告」、図解の中の文字は折り返すので「枠の行数を超えたら警告」

できないこと（意図的に小さく作ってあります）
  - Office の実際の描画とは違います。行間の癖・表の罫線・グラフは **PPTX にして確認**してください
  - 画像のマスク（曲線の切り抜き）は近似です
  - **足りない機能は、この一式を使う人がエージェントに足させる前提**です。ここは出発点です

使い方：.venv/bin/python scripts/preview.py [--open]
"""
import base64
import re
import json
import os
import sys
import webbrowser
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp       # noqa: E402  (tokens)
import build_potx_figma as bpf  # noqa: E402  (Copyright / Meta のブランド置換を共有する)
import fit_text as ft         # noqa: E402  (はみ出しの判定)

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND") or ROOT / "bigtree").resolve()
OUT = Path(os.environ.get("OFFICE3_OUT") or ROOT / "build").resolve()
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
    # 見出しだけ別の書体（"font": "heading"）のときは、その書体で描く。出力側（テーマの +mj / +mn）と同じ分け方
    return (f"font-family:'{bp.role_family(r)}',serif;" if r.get("font") == "heading" else "") + (
            f'font-size:{r["pt"] * PX}px;font-weight:{700 if r["weight"] == "Bold" else 400};'
            f'line-height:{r["line"]}%;')


# ---------------------------------------------------------------- 関所① 配色：パレットの一覧と、並び順の見本
# PLAYBOOK §5「見た目は3か所で見せて確定させる」の①。色の値だけでなく、**どの枠に何を置いたか**と
# **白黒にしても区別できるか（L*）**、**地色の上で読めるか（コントラスト比）**を一緒に見せる。
# 数字を並べるだけでは判断できないので、自動配色の順に塗った棒と、明地・暗地の文字見本を添える。
SLOT_ROLE = {"dk1": "基本の文字", "lt1": "基本の地", "dk2": "補助の文字", "lt2": "補助の地",
             "accent1": "有彩色1（自動配色の1番目・図形の既定）", "accent2": "有彩色2（自動配色の2番目）",
             "accent3": "グレー濃", "accent4": "グレー中", "accent5": "グレー淡",
             "accent6": "強調（自動では使われない）", "hlink": "リンク", "folHlink": "訪問済みリンク"}


def lstar(color):
    r, g, b = rgb(color)
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    y = 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)
    return 116 * y ** (1 / 3) - 16 if y > 0.008856 else 903.3 * y


def palette_html():
    names = TOKENS.get("_colors_name", {})
    sw = ""
    for k, v in COLORS.items():
        ink = "#fff" if contrast("lt1", k) >= contrast("dk1", k) else f"#{COLORS['dk1']}"
        sw += (f'<div class="sw" data-name="{k}" data-box="#{v}"><div class="chip" style="background:#{v};color:{ink}">'
               f'{esc(names.get(k, ""))}</div><b>{k}</b> <code>#{v}</code><br><small>{esc(SLOT_ROLE.get(k, ""))}</small><br>'
               f'<small>L* {lstar(k):.0f} ｜ 白地 {contrast(k, "lt1"):.1f} ｜ 暗地 {contrast(k, "dk1"):.1f}</small></div>')
    # 自動配色の順（accent1→5）で塗った棒。強調だけ accent6
    cats, bars = ["W1", "W2", "W3", "W4"], ""
    vals = [[62, 70, 66, 80], [40, 46, 52, 50], [30, 28, 34, 36], [22, 24, 20, 26], [12, 14, 16, 15]]
    for i, c in enumerate(cats):
        grp = "".join(f'<div class="bar" style="height:{v[i] * 2}px;background:{hexc("accent6") if (n == 0 and i == 3) else hexc(f"accent{n + 1}")}"></div>'
                      for n, v in enumerate(vals))
        bars += f'<div class="grp">{grp}<span>{c}</span></div>'
    sample = lambda bg, fg, sub, acc: (
        f'<div class="ground" style="background:{hexc(bg)}"><div style="color:{hexc(fg)};{font_css("title")}">スライドタイトル（見出し書体）</div>'
        f'<div style="color:{hexc(acc)};{font_css("lead")}">リード文：結論を1行で</div>'
        f'<div style="color:{hexc(fg)};{font_css("body")}">本文 14pt。数字 12,480 と English text</div>'
        f'<div style="color:{hexc(sub)};{font_css("caption")}">補足 11pt ※ 数値はダミー</div>'
        f'<div style="color:{hexc("accent6")};{font_css("lead")}">強調（accent6）</div></div>')
    return ('<section class="card palette"><h2>関所① 配色<small>tokens.json の colors ｜ '
            'L* は明るさ（白黒にしたときの見分け）、比は地色とのコントラスト（文字なら 3.0 以上）</small></h2>'
            f'<div class="sws">{sw}</div><div class="row"><div><h3>グラフの自動配色（系列 1→5 の順。W4 の系列1だけ強調）</h3>'
            f'<div class="bars">{bars}</div></div>'
            f'<div><h3>明地（本文）</h3>{sample("lt1", "dk1", "dk2", "accent1")}</div>'
            f'<div><h3>暗地（扉・表紙）</h3>{sample("dk1", "lt1", "accent5", "accent2")}</div></div></section>')


def data_uri(path):
    if not path.exists():
        return ""
    mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "svg": "image/svg+xml"}
    return f'data:{mime.get(path.suffix.lstrip(".").lower(), "application/octet-stream")};base64,' \
           + base64.b64encode(path.read_bytes()).decode()


MONO_SVG = {}


def is_mono_svg(path):
    """SVG の中の塗りが1色以下か。出力側（build_potx_figma.logo）と同じ判定にする
    （fill 属性と style="fill:..." の両方を見る）。単色ならテーマ色で塗り替えられる。"""
    if not path.exists() or path.suffix.lower() != ".svg":
        return False
    t = path.read_text(errors="ignore")
    fills = set(re.findall(r'fill="(#[0-9A-Fa-f]{6})"', t)) | set(re.findall(r'fill:\s*(#[0-9A-Fa-f]{6})', t))
    return len(fills) <= 1


rgb, blend, contrast = bp.rgb, bp.blend, bp.contrast      # 判定は build_potx に集約


def contrast_warn(color, bg, role):
    """地色に対して文字が読めるか。**直すのは人**だが、どこが危ないかは指させるようにする。"""
    if not bg:
        return ""
    ratio = contrast(color, bg)
    # しきい値は 3.0。WCAG の本文基準は 4.5 だが、そこまで出すと補足文字がほぼ全部鳴って
    # 警告が無視されるようになる。ここで拾いたいのは「読めない」（地色と同系）で、
    # 細かい詰めは人が目で見て決める（notes/TEXT_LAYOUT_NOTES.md §10）
    if ratio < 3.0:
        return f"コントラスト {ratio:.1f}:1（3.0 未満＝読みにくい）"
    return ""


def el(name, x, y, w, h, inner="", css="", warn=""):
    """1つの要素。名前と座標を data 属性で持たせ、カーソルとクリックで拾えるようにする。"""
    return (f'<div class="el{" warn" if warn else ""}" style="left:{x}px;top:{y}px;width:{w}px;height:{h}px;{css}" '
            f'data-name="{esc(name)}" data-box="{x:.0f}, {y:.0f}, {w:.0f}, {h:.0f}"'
            + (f' data-warn="{esc(warn)}"' if warn else "") + f'>{inner}</div>')


def text_el(name, box, text, role, color, align="left", opacity=1, wrap=False, bg=None):
    """wrap=True（図解の中）は枠の高さまで折り返せる。False（レイアウトのプレースホルダ）は1行で溢れる。"""
    x, y, w, h = box
    cw = contrast_warn(color, bg, role)
    if wrap:
        ok, need, avail = ft.fits_box(text, role, w, h)
        warn = "" if ok else f"{need}行必要（枠は{avail}行）"
    else:
        ok, over, ratio = ft.fits(text, role, w)
        warn = "" if ok else f"はみ出し {over:.0f}pt（枠の {ratio * 100:.0f}%）"
    warn = "／".join(x for x in (warn, cw) if x)
    inner = f'<span class="t">{esc(text).replace(chr(10), "<br>")}</span>'
    css = (f'color:{hexc(color, opacity)};{font_css(role)}text-align:{ALIGN.get(align, "left")};'
           'display:flex;align-items:center;' + ("" if wrap else "white-space:pre;"))
    return el(name, x, y, w, h, inner, css, warn)


def layout_html(lay, assets, spec_w=1440, spec_h=810):
    body = ""
    bg = lay["bg"]
    for it in lay["items"]:
        kind, name = it.get("kind"), it["name"]
        if kind == "picture":
            src = assets.get(it["image"], "")
            if "mask" in it:
                x, y, w, h = it["fill_box"]
                body += el(name, x, y, w, h,
                           f'<img src="{src}" style="width:100%;height:100%;object-fit:cover">',
                           "overflow:hidden;")
            else:
                # マスク無し：box（見える範囲）の中に fill_box の位置で画像を置いて切り抜く（出力の srcRect と同じ）
                x, y, w, h = it.get("box") or it["fill_box"]
                fx, fy, fw, fh = it["fill_box"]
                body += el(name, x, y, w, h,
                           f'<img src="{src}" style="position:absolute;left:{fx - x}px;top:{fy - y}px;'
                           f'width:{fw}px;height:{fh}px;object-fit:cover">', "overflow:hidden;")
        elif kind == "logo":
            x, y, w, h = it["box"]
            key = it.get("svg", "logo")
            src = assets.get(key, "")
            if MONO_SVG.get(key):
                # 単色 SVG は出力側でテーマ色に塗り替えられる。プレビューでも同じ色にしないと、
                # 「プレビューでは赤、PowerPoint ではピンク」という食い違いが起きる。
                inner = (f'<div style="width:100%;height:100%;background:{hexc(it.get("color", "dk1"), it.get("opacity", 1))};'
                         f'-webkit-mask:url({src}) center/contain no-repeat;mask:url({src}) center/contain no-repeat"></div>')
            else:
                # 多色 SVG はそのまま貼るが、**透明度は単色と同じように効かせる**。
                # 理由：透かし（opacity 0.12）を 100% で描くと、プレビューがロゴに占領されて
                # レイアウトの判断ができない。出力側は alpha を書いているので、食い違いにもなる。
                op = it.get("opacity", 1)
                fade = "" if op >= 1 else f";opacity:{op}"
                inner = f'<img src="{src}" style="width:100%;height:100%;object-fit:contain{fade}">'
            body += el(name, x, y, w, h, inner)
        elif kind == "rect":
            # picture から差し替えた item は fill_box / fill のことがある（PLAYBOOK §3-8）
            x, y, w, h = it.get("box") or it["fill_box"]
            fill = it.get("color") or it.get("fill") or "dk1"
            if it.get("gradient"):
                g = it["gradient"]
                stops = ",".join(f"{hexc(c, a)} {pos * 100:.0f}%" for pos, c, a in g["stops"])
                # OOXML の角度（0＝左→右、90＝上→下）を CSS（90deg＝左→右、180deg＝上→下）に直す
                body += el(name, x, y, w, h, "", f'background:linear-gradient({g.get("angle", 0) + 90}deg,{stops});')
            else:
                body += el(name, x, y, w, h, "", f'background:{hexc(fill, it.get("opacity", 1))};')
        elif kind == "lines":
            segs = it["segments"]
            xs = [v for sg in segs for v in (sg[0], sg[2])]
            ys = [v for sg in segs for v in (sg[1], sg[3])]
            d = " ".join(f"M{a} {b} L{c} {dd}" for a, b, c, dd in segs)
            body += (f'<svg class="el vec" viewBox="0 0 {spec_w} {spec_h}" style="left:0;top:0;width:{spec_w}px;height:{spec_h}px" '
                     f'data-name="{esc(name)}" data-box="{min(xs):.0f}, {min(ys):.0f}, {max(xs) - min(xs):.0f}, {max(ys) - min(ys):.0f}">'
                     f'<path d="{d}" fill="none" stroke="{hexc(it.get("color", "dk1"))}" stroke-opacity="{it.get("opacity", 1)}" '
                     f'stroke-width="{it.get("w_px", 1)}"/></svg>')
        elif kind in ("ph", "text"):
            # 出力と同じ文字を見せる（§ ③ プレビューと実物がズレていた）。{{sample}} も samples.json から埋める
            # （埋めないと「{{sample}}」の長さで はみ出し判定をしてしまい、実際の文言の長さを確かめられない）
            it = bpf.with_brand(bpf.with_sample(it, lay["name"]))
            if it.get("ph") == "sldNum":
                txt = "3"
            else:
                txt = it.get("text", "")
            if not txt:
                x, y, w, h = it["box"]
                body += el(name, x, y, w, h, '<span class="empty">（中身はスライドで入る）</span>',
                           "border:1px dashed rgba(0,0,0,.2);color:#999;font-size:20px;")
                continue
            e = text_el(name, it["box"], txt, it.get("style", "body"), it.get("color", "dk1"), it.get("align", "l"), bg=bg)
            if it.get("pill"):
                e = e.replace('class="el', 'class="el pill', 1)
            body += e
    guides = ""
    for kind, v in (lay.get("guides", []) + TOKENS.get("master_guides", [])):
        guides += (f'<div class="guide" style="left:{v}px;top:0;width:1px;height:100%"></div>' if kind == "X"
                   else f'<div class="guide" style="left:0;top:{v}px;width:100%;height:1px"></div>')
    return body, guides


def under_fill(items, upto, box, frame_bg="lt1"):
    frame_bg = GROUND[0] or frame_bg
    """その文字の下に敷かれている色を、重なり順に混ぜながら求める。"""
    cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
    under = frame_bg
    for it in items[:upto]:
        if it["type"] == "TEXT":
            continue
        if not it.get("fill"):          # 塗りの無い図形（輪・縁取り）は下地にならない（出力側と同じ判定）
            continue
        if it["x"] <= cx <= it["x"] + it["w"] and it["y"] <= cy <= it["y"] + it["h"]:
            under = blend(it["fill"], under, it.get("fillOpacity", 1))
    return under


GROUND = [None]


def organism_html(org):
    GROUND[0] = org.get("bg")          # 暗地に置く図解（目次など）は、その地色で判定する
    body = ""
    for idx, it in enumerate(org["items"]):
        x, y, w, h = it["x"], it["y"], it["w"], it["h"]
        t, name, fill, op = it["type"], it["name"], it.get("fill", "dk1"), it.get("fillOpacity", 1)
        if t == "TEXT":
            body += text_el(name, [x, y, w, h], it.get("text", ""), STYLE_KEY.get(it.get("style"), "body"),
                            fill, it.get("align", "LEFT"), op, wrap=True,
                            bg=under_fill(org["items"], idx, [x, y, w, h]))
        elif t == "ELLIPSE":
            body += el(name, x, y, w, h, "", paint(it) + "border-radius:50%;")
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
            body += el(name, x, y, w, h, "", paint(it))
    return body


def paint(it):
    """四角・円の塗りと縁。**塗りが無い図形は透明**にする（出力側は fill が無ければ noFill。
    以前は黒で塗っていて、円環だけの図が黒い円に見えた）"""
    css = f'background:{hexc(it["fill"], it.get("fillOpacity", 1))};' if it.get("fill") else ""
    if it.get("stroke"):
        css += f'border:{it.get("strokeW", 1)}px solid {hexc(it["stroke"])};'
    return css


def main():
    spec = json.loads((BRAND / "design/figma_layouts.json").read_text())
    orgs = json.loads((BRAND / "design/figma_organisms.json").read_text())["organisms"]
    fw, fh = spec["frame"]
    TOKENS["master_guides"] = spec.get("master_guides", [])
    assets = {"logo": data_uri(BRAND / bp.BRAND_INFO["logo"])}
    MONO_SVG["logo"] = is_mono_svg(BRAND / bp.BRAND_INFO["logo"])
    for lay in spec["layouts"]:
        for it in lay["items"]:
            if it.get("kind") == "picture":
                assets[it["image"]] = data_uri(BRAND / it["image"])
            elif it.get("svg"):
                assets[it["svg"]] = data_uri(BRAND / it["svg"])
                MONO_SVG[it["svg"]] = is_mono_svg(BRAND / it["svg"])

    cards = palette_html()
    for lay in spec["layouts"]:
        body, guides = layout_html(lay, assets, fw, fh)
        cards += (f'<section class="card" data-kind="layout"><h2>{esc(lay["name"])}'
                  f'<small>レイアウト ｜ {fw}×{fh}px</small></h2>'
                  f'<div class="frame" style="width:{fw}px;height:{fh}px;background:{hexc(lay["bg"])}">'
                  f'{body}<div class="guides">{guides}</div></div></section>')
    for name, org in orgs.items():
        cards += (f'<section class="card" data-kind="organism"><h2>{esc(name)}'
                  f'<small>図解 ｜ {org["w"]}×{org["h"]}px</small></h2>'
                  f'<div class="frame" style="width:{org["w"]}px;height:{org["h"]}px;background:{hexc(org.get("bg", "lt1"))}">'
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
 svg.el{position:absolute;overflow:visible;pointer-events:none} svg.el path{pointer-events:stroke}
 .guides{position:absolute;inset:0;pointer-events:none;display:none}
 .guide{position:absolute;background:rgba(200,60,60,.55)}
 body.guides-on .guides{display:block}
 body.outline-on .el{outline:1px solid rgba(0,120,255,.45)}
 .el:hover{outline:2px solid #0078ff !important}
 .el.warn{outline:2px solid #e2773a}
 .palette{background:#fff;padding:16px 20px;box-shadow:0 2px 12px rgba(0,0,0,.08);max-width:1400px}
 .palette h3{font-size:13px;margin:14px 0 6px}
 .sws{display:grid;grid-template-columns:repeat(6,1fr);gap:12px;font-size:12px}
 .sw{cursor:pointer} .chip{height:64px;border-radius:4px;border:1px solid rgba(0,0,0,.08);display:flex;align-items:flex-end;padding:6px;font-size:13px;font-weight:700;margin-bottom:4px}
 .row{display:flex;gap:28px;flex-wrap:wrap}
 .bars{display:flex;gap:18px;align-items:flex-end;height:190px;padding:0 8px;border-bottom:1px solid #ccc}
 .grp{display:flex;gap:3px;align-items:flex-end;position:relative;padding-bottom:0}
 .grp span{position:absolute;bottom:-20px;left:0;right:0;text-align:center;font-size:12px;color:#666}
 .bar{width:16px}
 .ground{width:720px;padding:28px 36px;display:flex;flex-direction:column;gap:8px;white-space:nowrap;zoom:.58}
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
