"""文字が枠に収まるかを、フォントの実寸で判定する。

折り返しをオフにしてあるので（notes/TEXT_LAYOUT_NOTES.md §4）、長すぎる文字は枠の外へ出る。
生成時に警告を出して、気づかないまま提出するのを防ぐ。

単体でも使える：
  .venv/bin/python scripts/fit_text.py lead "リード文に入れたい文章"
"""
import functools
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402  (FONT_DIR / tokens)

PX_PER_PT = 2          # Figma 1440px = 720pt
WARNED = set()


@functools.lru_cache(maxsize=None)
def _font(weight):
    from fontTools.ttLib import TTFont
    family = bp.FONT.replace(" ", "")
    f = TTFont(bp.FONT_DIR / f"{family}-{weight}.ttf", lazy=True)
    return f.getBestCmap(), f["hmtx"], f["head"].unitsPerEm


def width_pt(text, weight, pt):
    """文字列の表示幅（pt）。フォントの字送り（advance width）から計算する。"""
    cmap, hmtx, upm = _font("Bold" if weight == "Bold" else "Regular")
    total = 0.0
    for ch in text:
        g = cmap.get(ord(ch))
        total += (hmtx[g][0] if g else upm) / upm
    return total * pt


def fits(text, style, box_px, roles=None):
    """(収まるか, はみ出し幅pt, 使用率) を返す。改行は行ごとに見る。"""
    roles = roles or bp._TOKENS["type_pptx"]["roles"]
    r = roles[style]
    box_pt = box_px / PX_PER_PT
    worst = max((width_pt(line, r["weight"], r["pt"]) for line in text.split("\n")), default=0.0)
    return worst <= box_pt, worst - box_pt, (worst / box_pt if box_pt else 0)


def warn(text, style, box_px, where="", roles=None):
    """収まらなければ標準エラーに警告（同じ内容は1回だけ）。"""
    if not text:
        return True
    ok, over, ratio = fits(text, style, box_px, roles)
    if not ok:
        key = (text, style, box_px)
        if key not in WARNED:
            WARNED.add(key)
            head = text.replace("\n", " ")[:24]
            print(f"⚠ はみ出し {where}［{style}］{over:.0f}pt 超過（{ratio * 100:.0f}%）: 「{head}…」", file=sys.stderr)
    return ok


def limits(roles=None):
    """役割ごとの上限文字数（全角だけ／半角英数だけ）の目安。"""
    roles = roles or bp._TOKENS["type_pptx"]["roles"]
    out = {}
    for key, r in roles.items():
        out[key] = {"pt": r["pt"], "weight": r["weight"],
                    "full_per_100px": 100 / PX_PER_PT / width_pt("あ", r["weight"], r["pt"]),
                    "half_per_100px": 100 / PX_PER_PT / width_pt("A", r["weight"], r["pt"])}
    return out


if __name__ == "__main__":
    style, text = sys.argv[1], sys.argv[2]
    box = float(sys.argv[3]) if len(sys.argv) > 3 else 1296
    ok, over, ratio = fits(text, style, box)
    print(f"{style} / 枠 {box:.0f}px: {'収まる' if ok else f'{over:.0f}pt はみ出す'}（使用率 {ratio * 100:.0f}%）")
    for key, v in limits().items():
        print(f"  {key:<8} {v['pt']:>4}pt {v['weight']:<8} 全角 {v['full_per_100px'] * box / 100:.0f} 文字 / 半角 {v['half_per_100px'] * box / 100:.0f} 文字")
