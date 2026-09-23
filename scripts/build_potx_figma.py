"""Figma layout export (design/figma_layouts.json) + tokens (design/tokens.json) -> POTX and a sample PPTX.

Flow: Figma (SSOT) --use_figma export--> design/figma_layouts.json --this script--> build/<slug>.potx
Units: Figma px (1440x810) -> EMU at 6350 EMU/px (10in slide), font px/2 = pt.
Usage: .venv/bin/python scripts/build_potx_figma.py [--no-embed]
"""
import json
import math
import re
import os
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402  (theme, EOT embedding, guides helper)
import fit_text as ft  # noqa: E402  (はみ出し警告)

ROOT = Path(__file__).resolve().parent.parent
BRAND = Path(os.environ.get("OFFICE3_BRAND") or ROOT / "bigtree").resolve()   # brand values (tokens, assets, templates)
OUT = Path(os.environ.get("OFFICE3_OUT") or ROOT / "build").resolve()         # generated files (gitignored)
OUT.mkdir(parents=True, exist_ok=True)
SPEC = json.loads((BRAND / "design/figma_layouts.json").read_text())
# サンプルの文言（表紙・扉・本文）はブランド側に置く。スクリプトやレイアウト書き出しに
# 実文を持たせると、ブランドを差し替えた人の資料に作者の文章が残る
_SP = BRAND / "design/samples.json"
_SAMPLES = json.loads(_SP.read_text()) if _SP.exists() else {}
EMU = bp.W / SPEC["frame"][0]  # 6350
XML, NS, REL, CT = bp.XML, bp.NS, bp.REL, bp.CT
SCHEME = {"dk1": "tx1", "lt1": "bg1", "dk2": "tx2", "lt2": "bg2"}
STYLES = bp._TOKENS["type_pptx"]["roles"]   # 文字スタイルの正は tokens.json
BRAND_NAME = bp.BRAND_INFO["name"]         # ブランド名の正は tokens.json の brand

# 本文レイアウトの地色が、この資料が明地か暗地かの判断材料。
# 暗地にしたのにグラフ・表・図解が明地のままだと、文字が地色と同化して消える（実際に踏まれた）
# 本文レイアウトは `"role": "body"` の印で探す（無ければ従来の名前 "06_本文"）。
# 名前で探していたため、レイアウトを並べ替えたり名前を変えたりすると、マスター・グラフ・表・図解の
# 5か所が同時に「06_本文 が無い」で止まった。印なら名前も順番も自由に変えられる
BODY_LAYOUT = (next((l for l in SPEC["layouts"] if l.get("role") == "body"), None)
               or next(l for l in SPEC["layouts"] if l["name"] == "06_本文"))
BODY_LAYOUT_NO = SPEC["layouts"].index(BODY_LAYOUT) + 1     # slideLayoutN.xml の N
BODY_BG = BODY_LAYOUT.get("bg", "lt1")
# その地色の上で読める文字の**トークン名**を聞く。トークンの意味（dk1＝文字色）は
# 配色を反転した人でも保つので、名前で受け取るのが安全。
# 明るさで直に判定しようとすると、トークン名と実際の色が入れ替わったブランドで逆を引く（実際に踏んだ）
TEXT_ON_SLIDE = bp.text_on(BODY_BG)               # "dk1" か "lt1"
MUTED_ON_SLIDE = "dk2" if TEXT_ON_SLIDE == "dk1" else "lt2"
ON_DARK = bp.contrast("lt1", BODY_BG) > bp.contrast("dk1", BODY_BG)   # 地の上で lt1 のほうが読める＝暗地


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
    # 見出し書体の役割（"font": "heading"）はテーマの見出しフォント（+mj）を指す。書体名は直書きしない
    # （直書きすると、テーマのフォントを変えても文字が追随しない）
    mj = s.get("font") == "heading"
    font = ('<a:latin typeface="Noto Sans JP Black"/><a:ea typeface="Noto Sans JP Black"/>' if s["weight"] == "Black"
            else f'<a:latin typeface="+{"mj" if mj else "mn"}-lt"/><a:ea typeface="+{"mj" if mj else "mn"}-ea"/>')
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


# ---------------------------------------------------------------- logo geometry (SVG -> custGeom)
# 対応：viewBox のオフセット、複数の <path>、path ごとの fill、M/L/H/V/C/S（絶対・相対）と Z。
# 多色のロゴは「1パス＝1図形」で重ねる。fill が無い SVG は、レイアウト側の色（テーマ色）で塗る。
SVG_NUM = re.compile(r"[-+]?(?:\d*\.\d+|\d+)(?:[eE][-+]?\d+)?")
SVG_CMD = re.compile(r"([MmLlHhVvCcSsZz])([^MmLlHhVvCcSsZz]*)")


def svg_subpaths(d):
    """SVG の d 属性を、絶対座標のサブパス [(始点, [('L',p) | ('C',p1,p2,p3), ...]), ...] にする。"""
    subs, cur, start, pos, prev_c2 = [], None, (0.0, 0.0), (0.0, 0.0), None
    for cmd, body in SVG_CMD.findall(d):
        n = [float(v) for v in SVG_NUM.findall(body)]
        rel = cmd.islower()
        c = cmd.upper()
        if c == "Z":
            if cur:
                subs.append(cur)
                cur, pos, prev_c2 = None, start, None
            continue
        i = 0
        while i < len(n) or (c == "M" and i == 0):
            if c == "M":
                p = (n[i] + (pos[0] if rel else 0), n[i + 1] + (pos[1] if rel else 0))
                if cur:
                    subs.append(cur)
                cur, start, pos, prev_c2 = (p, []), p, p, None
                i += 2
                c = "L"          # M のあとに続く座標は L 扱い（SVG の仕様）
            elif c in ("L", "T"):
                p = (n[i] + (pos[0] if rel else 0), n[i + 1] + (pos[1] if rel else 0))
                cur[1].append(("L", p)); pos, prev_c2 = p, None; i += 2
            elif c == "H":
                p = (n[i] + (pos[0] if rel else 0), pos[1])
                cur[1].append(("L", p)); pos, prev_c2 = p, None; i += 1
            elif c == "V":
                p = (pos[0], n[i] + (pos[1] if rel else 0))
                cur[1].append(("L", p)); pos, prev_c2 = p, None; i += 1
            elif c == "C":
                ox, oy = pos if rel else (0.0, 0.0)
                p1, p2, p3 = ((n[i] + ox, n[i + 1] + oy), (n[i + 2] + ox, n[i + 3] + oy), (n[i + 4] + ox, n[i + 5] + oy))
                cur[1].append(("C", p1, p2, p3)); pos, prev_c2 = p3, p2; i += 6
            elif c == "S":
                ox, oy = pos if rel else (0.0, 0.0)
                p1 = (2 * pos[0] - prev_c2[0], 2 * pos[1] - prev_c2[1]) if prev_c2 else pos
                p2, p3 = (n[i] + ox, n[i + 1] + oy), (n[i + 2] + ox, n[i + 3] + oy)
                cur[1].append(("C", p1, p2, p3)); pos, prev_c2 = p3, p2; i += 4
            else:
                raise ValueError(f"未対応の SVG コマンド: {cmd}")
            if i >= len(n):
                break
    if cur:
        subs.append(cur)
    return subs


SVG_TF = re.compile(r"(matrix|translate|scale|rotate)\s*\(([^)]*)\)")


def tf_mul(m, n):
    """2つのアフィン変換 (a,b,c,d,e,f) を合成する（m のあとに n を内側で適用）。"""
    a, b, c, d, e, f = m
    A, B, C, D, E, F = n
    return (a * A + c * B, b * A + d * B, a * C + c * D, b * C + d * D, a * E + c * F + e, b * E + d * F + f)


def tf_parse(s):
    """transform 属性 -> アフィン変換。translate / scale / rotate / matrix に対応。"""
    m = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    for fn, body in SVG_TF.findall(s or ""):
        v = [float(x) for x in SVG_NUM.findall(body)]
        if fn == "translate":
            n = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0)
        elif fn == "scale":
            n = (v[0], 0, 0, v[1] if len(v) > 1 else v[0], 0, 0)
        elif fn == "rotate":
            r = math.radians(v[0])
            n = (math.cos(r), math.sin(r), -math.sin(r), math.cos(r), 0, 0)
            if len(v) == 3:   # 回転中心つき
                n = tf_mul(tf_mul((1, 0, 0, 1, v[1], v[2]), n), (1, 0, 0, 1, -v[1], -v[2]))
        else:
            n = tuple(v[:6])
        m = tf_mul(m, n)
    return m


def tf_apply(m, p):
    a, b, c, d, e, f = m
    return (a * p[0] + c * p[1] + e, b * p[0] + d * p[1] + f)


def logo_paths(rel=None):
    """ロゴ SVG を (幅, 高さ, [(fill|None, サブパス), ...]) にする。
    <g> の transform を積み、viewBox の原点を 0 に寄せた絶対座標で返す。
    rel を渡すと、ブランドの既定ロゴではなくその SVG（アイコン・フッターなど）を読む。"""
    path = BRAND / (rel or bp.BRAND_INFO["logo"])
    if not path.exists():
        # ロゴがまだ無くても、枠だけで一通り作れるようにする（あとで差し替える前提）
        print(f"（ロゴが無いので、四角い枠で代用します: {path}）", file=sys.stderr)
        return 100.0, 100.0, [(None, [((0.0, 0.0), [("L", (100.0, 0.0)), ("L", (100.0, 100.0)), ("L", (0.0, 100.0))])])]
    svg = path.read_text()
    vb = re.search(r'viewBox="([-\d.eE]+)[,\s]+([-\d.eE]+)[,\s]+([\d.eE]+)[,\s]+([\d.eE]+)"', svg)
    if not vb:
        raise ValueError(f'{bp.BRAND_INFO["logo"]} に viewBox がありません')
    vx, vy, vw, vh = [float(v) for v in vb.groups()]
    out, stack = [], [((1.0, 0.0, 0.0, 1.0, 0.0, 0.0), None)]   # (変換, 継承した fill)
    for tag in re.findall(r"<[^>]+>", svg):
        name = re.match(r"<\s*(/?)\s*([a-zA-Z]+)", tag)
        if not name:
            continue
        closing, el = name.group(1), name.group(2)
        if el == "g" and closing:
            if len(stack) > 1:
                stack.pop()
            continue
        tr = re.search(r' transform="([^"]*)"', tag)
        # Illustrator の SVG 書き出しは fill 属性ではなく style="fill:#RRGGBB" を書く。
        # 両方を見ないと、多色ロゴが1色に潰れたまま（警告も出ずに）出力される。
        fl = (re.search(r' fill="(#[0-9A-Fa-f]{6})"', tag)
              or re.search(r'style="[^"]*\bfill:\s*(#[0-9A-Fa-f]{6})', tag))
        m = tf_mul(stack[-1][0], tf_parse(tr.group(1))) if tr else stack[-1][0]
        fill = fl.group(1)[1:].upper() if fl else stack[-1][1]
        if el == "g":
            if not tag.rstrip().endswith("/>"):
                stack.append((m, fill))
            continue
        if el != "path":
            continue
        d = re.search(r' d="([^"]+)"', tag)
        if not d:
            continue
        subs = []
        for start, segs in svg_subpaths(d.group(1)):
            T = lambda p: (lambda q: (q[0] - vx, q[1] - vy))(tf_apply(m, p))
            subs.append((T(start), [(seg[0], *[T(p) for p in seg[1:]]) for seg in segs]))
        out.append((fill, subs))
    if not out:
        raise ValueError(f'{rel or bp.BRAND_INFO["logo"]} に <path> がありません')
    return vw, vh, out


LOGO_W, LOGO_H, LOGO_SHAPES = logo_paths()
LOGO_MULTICOLOR = len({f for f, _ in LOGO_SHAPES if f}) > 1
_SVG_CACHE = {}


def svg_shapes(rel):
    """レイアウトの item が svg を指していたら、その SVG を読む（同じものは使い回す）。"""
    if rel not in _SVG_CACHE:
        w, h, shapes = logo_paths(rel)
        _SVG_CACHE[rel] = (w, h, shapes, len({f for f, _ in shapes if f}) > 1)
    return _SVG_CACHE[rel]


def logo(item):
    """ロゴ。多色 SVG は path ごとに図形を重ねる（単色 SVG はレイアウトの色で塗る）。
    パス座標は図形の枠の EMU に直して書く（<a:path w/h> の換算に頼らない。頼ると
    その換算を見ないレンダラで小さく描かれる）。"""
    bw, bh = e(item["box"][2]), e(item["box"][3])
    w0, h0, shapes, multi = svg_shapes(item["svg"]) if item.get("svg") else (LOGO_W, LOGO_H, LOGO_SHAPES, LOGO_MULTICOLOR)
    # 枠の縦横比が SVG と違っても、SVG の比率は崩さない（プレビューの object-fit:contain と同じ）。
    # 崩す指定は事故なので警告する。ここを伸ばすと、プレビューと実物が違う見た目になる。
    sx = sy = min(bw / w0, bh / h0)
    if abs((bw / bh) - (w0 / h0)) > 0.01 * (w0 / h0):
        print(f'（{item["name"]}: 枠 {item["box"][2]:.0f}x{item["box"][3]:.0f} と SVG {w0:.0f}x{h0:.0f} の比率が違います。'
              f'中央に収めました。枠を {w0 / h0:.3f} の比率にしてください'
              f'／まとめて直すなら scripts/fit_logo_boxes.py）', file=sys.stderr)
    ox, oy = (bw - w0 * sx) / 2, (bh - h0 * sy) / 2   # 収めたぶんを中央に寄せる
    P = lambda p: f'<a:pt x="{int(p[0] * sx + ox)}" y="{int(p[1] * sy + oy)}"/>'
    out = []
    for n, (fill, subs) in enumerate(shapes):
        # SVG の <path> ひとつ＝ <a:path> ひとつ。サブパスは moveTo で continue する。
        # サブパスごとに <a:path> を分けると、文字の「A」の中の穴まで塗られる（穴が穴にならない）。
        body = ""
        for start, segs in subs:
            body += f'<a:moveTo>{P(start)}</a:moveTo>'
            for seg in segs:
                body += (f'<a:lnTo>{P(seg[1])}</a:lnTo>' if seg[0] == "L"
                         else f'<a:cubicBezTo>{P(seg[1])}{P(seg[2])}{P(seg[3])}</a:cubicBezTo>')
            body += '<a:close/>'
        paths = f'<a:path w="{bw}" h="{bh}">{body}</a:path>'
        if multi and fill:
            op = item.get("opacity")
            a = f'<a:alpha val="{int(op * 100000)}"/>' if op is not None else ""
            paint = f'<a:solidFill><a:srgbClr val="{fill}">{a}</a:srgbClr></a:solidFill>'
        else:
            paint = clr(item.get("color", "dk1"), item.get("opacity"))   # 多色 SVG では使われない
        name = item["name"] if len(shapes) == 1 else f'{item["name"]}_{n + 1}'
        out.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{name}" descr="{BRAND_NAME} ロゴ"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(item["box"])}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst>{paths}</a:pathLst></a:custGeom>{paint}<a:ln><a:noFill/></a:ln></p:spPr></p:sp>')
    return "".join(out)


# ---------------------------------------------------------------- key visual: picture cropped to the Figma curve
def picture(item, rid):
    if "mask" not in item:
        return picture_rect(item, rid)
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
    # 代替テキスト（スクリーンリーダーが読む）はレイアウト JSON の alt から。無ければ図の名前
    alt = bp_esc(item.get("alt") or item["name"])
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}" descr="{alt}"/>'
            '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/>{src}<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{xfrm(box)}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{e(box[2])}" h="{e(box[3])}">{segs}</a:path></a:pathLst></a:custGeom></p:spPr></p:pic>')


def picture_rect(item, rid):
    """曲線のマスクを使わない写真。`box`（見える範囲）の中に、`fill_box`（画像を置く位置と大きさ）の
    画像を切り抜いて入れる。fill_box は box より大きくてよく、はみ出した分が切り落とされる。
    切り抜きは srcRect で持つので、PowerPoint 上で人が「トリミング」から位置を直せる。"""
    box = item.get("box") or item["fill_box"]
    fx, fy, fw, fh = item["fill_box"]
    iw, ih = item["img_px"]
    sc = max(fw / iw, fh / ih)                      # fill_box を画像の比率のまま覆う（cover）
    dw, dh = iw * sc, ih * sc
    ix, iy = fx - (dw - fw) / 2, fy - (dh - fh) / 2
    l, t = (box[0] - ix) / dw, (box[1] - iy) / dh
    r, b = (ix + dw - (box[0] + box[2])) / dw, (iy + dh - (box[1] + box[3])) / dh
    src = f'<a:srcRect l="{int(l * 100000)}" t="{int(t * 100000)}" r="{int(r * 100000)}" b="{int(b * 100000)}"/>'
    alt = bp_esc(item.get("alt") or item["name"])
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}" descr="{alt}"/>'
            '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/>{src}<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')


def grad(g):
    """グラデーションの塗り。stops は [[位置0〜1, 色トークン, 不透明度0〜1], ...]、angle は度（0＝左→右、90＝上→下）。
    色はテーマ参照のまま（配色を変えると追随する）。写真の上に地色を溶かす用途を想定"""
    gs = "".join(f'<a:gs pos="{int(pos * 100000)}"><a:schemeClr val="{SCHEME.get(c, c)}"><a:alpha val="{int(a * 100000)}"/></a:schemeClr></a:gs>'
                 for pos, c, a in g["stops"])
    return f'<a:gradFill rotWithShape="1"><a:gsLst>{gs}</a:gsLst><a:lin ang="{int(g.get("angle", 0) * 60000)}" scaled="0"/></a:gradFill>'


def lines(item):
    """細い線の束（光の筋・ガラスの稜線など）。segments は [[x1, y1, x2, y2], ...]（フレームの px）。
    1つの図形にまとめる（線ごとに図形を作ると、選択や移動で人がばらばらにしてしまう）。
    線の太さは w_px（2px＝1pt）、色は color / opacity"""
    xs = [v for sgm in item["segments"] for v in (sgm[0], sgm[2])]
    ys = [v for sgm in item["segments"] for v in (sgm[1], sgm[3])]
    x0, y0 = min(xs), min(ys)
    w, h = max(max(xs) - x0, 1), max(max(ys) - y0, 1)
    P = lambda x, y: f'<a:pt x="{e(x - x0)}" y="{e(y - y0)}"/>'
    body = "".join(f'<a:moveTo>{P(a, b)}</a:moveTo><a:lnTo>{P(c, d)}</a:lnTo>' for a, b, c, d in item["segments"])
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm([x0, y0, w, h])}<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{e(w)}" h="{e(h)}" fill="none">{body}</a:path></a:pathLst></a:custGeom><a:noFill/>'
            f'<a:ln w="{e(item.get("w_px", 1))}">{clr(item.get("color", "dk1"), item.get("opacity"))}</a:ln></p:spPr></p:sp>')


# ---------------------------------------------------------------- shapes
def rect(item):
    # PLAYBOOK §3-8 は「キービジュアルが無ければ kind:"picture" を "rect" に置き換える」と案内している。
    # picture の item は box/color ではなく fill_box/fill を持つため、どちらの綴りも受ける。
    box = item.get("box") or item["fill_box"]
    color = item.get("color") or item.get("fill") or "lt2"
    # 透明度（opacity）とグラデーション（gradient）も受ける。以前は opacity を黙って捨てていた（プレビューだけ半透明）
    paint = grad(item["gradient"]) if item.get("gradient") else clr(color, item.get("opacity"))
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{bp.nid()}" name="{item["name"]}"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
            f'<p:spPr>{xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{paint}<a:ln><a:noFill/></a:ln></p:spPr>'
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


# Figma の書き出しにはブランドの文字列が入っている。名前で拾って tokens.json の brand で上書きする
# （そうしないと、トークンを差し替えても「© 2026 前のブランド名」が残る）
LAYOUT_TEXTS = _SAMPLES.get("layout_texts", {})


def with_sample(item, layout_name):
    """`{{sample}}` を samples.json の文言に置き換える。

    レイアウトの書き出しに実文を持たせると、ブランドを差し替えた人の表紙に
    作者の文章（「note トラフィックレポート」など）が残る。文言は samples.json に集約する。
    """
    if item.get("text") == "{{sample}}":
        return {**item, "text": LAYOUT_TEXTS.get(layout_name, {}).get(item["name"], "")}
    return item


def with_brand(item):
    """`{{brand.xxx}}` と書かれた文字を tokens.json の brand の値に置き換える。

    著作権表示や作成者名は、書き出し JSON にも Figma にも実物が書いてあると**同じ値が2か所**になり、
    JSON を直しても出力が変わらない（＝直したつもりで直っていない）状態になる。
    そこで JSON 側は差し込み口だけを残し、値は brand から取る。
    """
    text = item.get("text")
    if not text or "{{" not in text:
        return item
    for key, val in bp.BRAND_INFO.items():
        text = text.replace("{{brand.%s}}" % key, str(val))
    return {**item, "text": text}


def placeholder(item, on_slide=False, content=None, master=False):
    item = with_brand(item)
    attr = ph_attr(item)
    # 折り返しをオフにしてあるので、長すぎる文字は枠の外へ出る。作った時点で気づけるように警告する
    if item.get("text") and item["ph"] not in ("obj", "sldNum"):
        # 枠線つきの区分表示（pill）は左右の内側余白ぶん狭い。差し引かずに測ると、枠の線に文字が重なっても通ってしまう
        room = item["box"][2] - (2 * item["pill"]["pad"][0] if item.get("pill") else 0)
        ft.warn(item["text"], item["style"], room, where=item["name"])
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
    # 縞は「文字色を 8% 乗せる」。明地なら薄いグレー、暗地なら少し明るい帯になり、どちらでも成立する
    band_fill = (f'<a:fill><a:solidFill><a:schemeClr val="{SCHEME.get(TEXT_ON_SLIDE, TEXT_ON_SLIDE)}">'
                 '<a:alpha val="8000"/></a:schemeClr></a:solidFill></a:fill>')
    # 縞は「文字色を 8% 乗せる」。明地なら薄いグレー、暗地なら少し明るい帯になり、どちらでも成立する
    band_fill = (f'<a:fill><a:solidFill><a:schemeClr val="{SCHEME.get(TEXT_ON_SLIDE, TEXT_ON_SLIDE)}">'
                 '<a:alpha val="8000"/></a:schemeClr></a:solidFill></a:fill>')
    inside = ("<a:left>" + nl + "</a:left><a:right>" + nl + "</a:right><a:top>" + nl + "</a:top><a:bottom>" + nl + "</a:bottom>"
              + "<a:insideH>" + ln(1, "accent5") + "</a:insideH><a:insideV>" + nl + "</a:insideV>")
    header_bdr = ("<a:left>" + nl + "</a:left><a:right>" + nl + "</a:right><a:top>" + nl + "</a:top>"
                  + "<a:bottom>" + ln(2, "accent1") + "</a:bottom><a:insideV>" + nl + "</a:insideV>")
    return (XML + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            f'def="{TABLE_STYLE_ID}"><a:tblStyle styleId="{TABLE_STYLE_ID}" styleName="{BRAND_NAME} 表">'
            # whole table: horizontal rules only, no outer frame (0304 rule)
            f'<a:wholeTbl>{cell(inside, none)}{txt(TEXT_ON_SLIDE, False)}</a:wholeTbl>'
            f'<a:band2H>{cell("", band_fill)}</a:band2H>'
            # header row: accent1 band, white bold text, thicker rule underneath
            f'<a:firstRow>{txt(bp.text_on("accent1"), True)}{cell(header_bdr, fill_of("accent1"))}</a:firstRow>'
            # first column: bold row labels
            f'<a:firstCol>{txt(TEXT_ON_SLIDE, True)}{cell("", none)}</a:firstCol>'
            '</a:tblStyle></a:tblStyleLst>')


# ---------------------------------------------------------------- master / layouts
def master(n_layouts):
    bp._id[0] = 1
    body_layout = BODY_LAYOUT
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
        elif k == "lines":
            shapes.append(lines(it))
        elif k == "text":
            shapes.append(text(with_brand(with_sample(it, spec["name"]))))
        elif k == "ph":
            shapes.append(placeholder(with_sample(it, spec["name"])))
    g = guides_xml(spec["guides"], "{DCECCB84-F9BA-43D5-87BE-67443E8EF086}", "layout")
    xml = (XML + f'<p:sldLayout {NS} preserve="1" userDrawn="1" showMasterSp="0"><p:cSld name="{spec["name"]}">'
           f'<p:bg><p:bgPr>{clr(spec["bg"])}<a:effectLst/></p:bgPr></p:bg>{bp.sptree(shapes)}</p:cSld>'
           f'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>{g}</p:sldLayout>')
    return xml, img


# ---------------------------------------------------------------- sample slides (one per layout)
# サンプルの文章はブランド側（design/samples.json）に置く。
# **スクリプトに書くと、ブランドを差し替えた人の資料に作者の文章が残る**（実際に踏まれた）
SAMPLE_CONTENT = {k: [tuple(x) for x in v] for k, v in _SAMPLES.get("slides", {}).items()}


def sample_slide(spec):
    bp._id[0] = 1
    shapes = []
    for it in spec["items"]:
        if it["kind"] != "ph":
            continue
        if it["ph"] == "obj":
            lines = SAMPLE_CONTENT.get(spec["name"], [])
            body = "".join(f'<a:p><a:pPr lvl="{lv}"/><a:r><a:rPr lang="ja-JP"/><a:t>{bp_esc(t)}</a:t></a:r></a:p>' for lv, t in lines)
            shapes.append(placeholder(with_sample(it, spec["name"]), on_slide=True, content=body))
        else:
            shapes.append(placeholder(with_sample(it, spec["name"]), on_slide=True))
    return XML + f'<p:sld {NS}><p:cSld>{bp.sptree(shapes)}</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'


# ---------------------------------------------------------------- package
def rels_xml(items):
    return (XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="rId{i + 1}" Type="{t}" Target="{tg}"/>' for i, (t, tg) in enumerate(items))
            + '</Relationships>')


def R(t):
    return f"{REL}/{t}"


def build(path, template, embed, slides=None):
    """slides を渡すと、見本スライドの代わりにそのスライドを入れる（資料そのものを組むとき）。
    slides は [{"layout_no": N, "xml": スライドXML, "rels": [(関係の種類URI, 参照先)], "files": {パス: 中身},
    "ct": [(パーツ名, コンテンツタイプ)], "defaults": [(拡張子, コンテンツタイプ)]}, ...]。
    rels は rId2 から順に振られる（rId1 はレイアウト）。テンプレートと同じ部品（マスター・テーマ・表スタイル・
    フォント埋め込み）で組むので、資料とテンプレートの見た目がずれない"""
    files, over = {}, []
    extra_defaults = []
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
    if not template and slides is not None:
        for i, sl in enumerate(slides, 1):
            files[f"ppt/slides/slide{i}.xml"] = sl["xml"]
            files[f"ppt/slides/_rels/slide{i}.xml.rels"] = rels_xml(
                [(R("slideLayout"), f'../slideLayouts/slideLayout{sl["layout_no"]}.xml')] + list(sl.get("rels", [])))
            files.update(sl.get("files", {}))
            over += list(sl.get("ct", []))
            extra_defaults += [d for d in sl.get("defaults", []) if d not in extra_defaults]
            over.append((f"/ppt/slides/slide{i}.xml", f"{CT}.presentationml.slide+xml"))
            pres_rels.append((R("slide"), f"slides/slide{i}.xml"))
            sld_ids += f'<p:sldId id="{255 + i}" r:id="rId{len(pres_rels)}"/>'
    elif not template:
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
                files[f"ppt/fonts/font{n}.fntdata"] = bp.make_eot(bp.font_file(ttf.split("-")[-1].replace(".ttf", ""), face))
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
        # 既定の文字サイズもトークンから。ベタ書きすると、本文を変えてもここだけ古い値が残る
        + f'<p:defaultTextStyle><a:defPPr><a:defRPr lang="ja-JP"/></a:defPPr><a:lvl1pPr><a:defRPr sz="{int(STYLES["body"]["pt"] * 100)}">'
        f'{clr("dk1")}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle></p:presentation>')
    files["ppt/_rels/presentation.xml.rels"] = rels_xml(pres_rels)
    files["ppt/presProps.xml"] = XML + f'<p:presentationPr {NS}/>'
    files["ppt/viewProps.xml"] = XML + f'<p:viewPr {NS}><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
    files["ppt/tableStyles.xml"] = table_styles()
    files["docProps/core.xml"] = (XML + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                                  f'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{BRAND_NAME} template</dc:title>'
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
        + "".join(f'<Default Extension="{x}" ContentType="{c}"/>' for x, c in extra_defaults if x not in ("png", "jpg", "fntdata"))
        + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in over) + '</Types>')
    order = ["[Content_Types].xml", "_rels/.rels"] + [k for k in files if k not in ("[Content_Types].xml", "_rels/.rels")]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k in order:
            v = files[k]
            z.writestr(k, v if isinstance(v, bytes) else v.encode("utf-8"))
    print(f"wrote {path.relative_to(OUT.parent)} ({path.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    embed = "--no-embed" not in sys.argv
    build(OUT / bp.out_name(".potx"), template=True, embed=embed)
    build(OUT / bp.out_name("_sample.pptx"), template=False, embed=embed)
