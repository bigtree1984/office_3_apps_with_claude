"""レイアウトのロゴ枠を、自分のロゴの縦横比に合わせる。

**なぜ要るか**：`figma_layouts.json` のロゴ枠は、作者のロゴ（縦長）の比率で作ってあります。
正方形や横長のロゴに差し替えると、ビルドのたびに全レイアウトで比率の警告が出ます
（黙って引き伸ばさずに中央へ収める仕様のため。`notes/` 参照）。
このスクリプトは、**枠の位置を保ったまま**比率だけ合わせ直します。

使い方：
  OFFICE3_BRAND=~/mybrand .venv/bin/python scripts/fit_logo_boxes.py          # 変更内容を表示するだけ
  OFFICE3_BRAND=~/mybrand .venv/bin/python scripts/fit_logo_boxes.py --write  # 実際に書き換える

  --anchor=tl（既定）… 左上を固定して合わせる。ページの角に置いたロゴ向き。
                       枠がフレームからはみ出す場合は、逆側の端を保ちます
  --anchor=center    … 中心を固定して合わせる。面の中央に置いた地紋（ウォーターマーク）向き

**枠は「もとの枠に収まる大きさ」になります。** 高さだけ保つと、横長のロゴでスライドの外へ出るためです。
そのぶん縦長の枠に横長のロゴを入れると小さくなるので、**最後は自分で枠を決め直してください**
（このスクリプトは警告を消すための下ごしらえで、レイアウトの意図までは決められません）。
"""
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp             # noqa: E402  (brand.logo)
import build_potx_figma as bpf      # noqa: E402  (SVG の読み取り)

BRAND = Path(os.environ.get("OFFICE3_BRAND") or Path(__file__).resolve().parent.parent / "bigtree").resolve()
FRAME_W, FRAME_H = bpf.SPEC["frame"]


def main():
    write = "--write" in sys.argv
    anchor = next((a.split("=")[1] for a in sys.argv if a.startswith("--anchor=")), "tl")
    path = BRAND / "design/figma_layouts.json"
    spec = json.loads(path.read_text())

    w0, h0, _ = bpf.logo_paths()                       # 既定ロゴの実寸
    print(f"ロゴ {bp.BRAND_INFO['logo']}: {w0:.0f}x{h0:.0f}（比率 {w0 / h0:.3f}）")

    changed = 0
    for lay in spec["layouts"]:
        for it in lay["items"]:
            if it.get("kind") != "logo":
                continue
            sw, sh, _ = bpf.svg_shapes(it["svg"])[:3] if it.get("svg") else (w0, h0, None)
            x, y, w, h = it["box"]
            if abs((w / h) - (sw / sh)) <= 0.01 * (sw / sh):
                continue
            # **もとの枠に収まる大きさにする**（高さだけ保つと、横長ロゴで枠がスライドの外へ出る。
            # 実際、比率 4.375 のワードマークで幅 2835px＝フレーム 1440px の2倍になった）
            k = min(w / sw, h / sh)
            nw, nh = sw * k, sh * k
            if anchor == "tl":
                nx, ny = x, y
                if nx + nw > FRAME_W:                  # 右寄せの枠は右端を保つ
                    nx = max(0, x + w - nw)
                if ny + nh > FRAME_H:
                    ny = max(0, y + h - nh)
            else:
                nx, ny = x + (w - nw) / 2, y + (h - nh) / 2
            print(f"  {lay['name']:<16} {it['name']:<12} {w:.1f}x{h:.1f} → {nw:.1f}x{nh:.1f}"
                  + ("" if anchor == "tl" else f"（中心を保持: x {x:.1f} → {nx:.1f}）"))
            it["box"] = [round(nx, 2), round(ny, 2), round(nw, 2), round(nh, 2)]
            changed += 1

    if not changed:
        print("すべて合っています。書き換えるところはありません。")
        return
    if write:
        path.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n")
        print(f"{changed} か所を書き換えました: {path}")
    else:
        print(f"{changed} か所が対象です。実際に書き換えるには --write を付けてください。")


if __name__ == "__main__":
    main()
