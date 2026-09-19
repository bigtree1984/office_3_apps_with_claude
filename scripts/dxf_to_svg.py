"""Fusion sketch DXF (LINE entities in the XZ plane) -> filled SVG.

Snaps endpoints, nodes all lines, polygonizes, and dissolves the faces into one filled shape.
Construction lines that stick out (dangling) do not form faces and drop out automatically.
Usage: .venv/bin/python scripts/dxf_to_svg.py in.dxf out.svg [#color]
"""
import os
import sys
import ezdxf
from shapely import set_precision
from shapely.geometry import LineString, MultiPolygon
from shapely.ops import polygonize, unary_union

src, dst = sys.argv[1], sys.argv[2]
color = sys.argv[3] if len(sys.argv) > 3 else "#1F4437"
msp = ezdxf.readfile(src).modelspace()
pts = [(e.dxf.start, e.dxf.end) for e in msp.query("LINE")]
top = max(max(a.z, b.z) for a, b in pts)
lines = [LineString([(a.x, top - a.z), (b.x, top - b.z)]) for a, b in pts if (a - b).magnitude > 1e-6]
noded = set_precision(unary_union([set_precision(l, 1e-3) for l in lines]), 1e-3)
faces = [f for f in polygonize(noded) if f.area > 1e-3]
shape = unary_union(faces)
geoms = list(shape.geoms) if isinstance(shape, MultiPolygon) else [shape]
minx, miny, maxx, maxy = shape.bounds
pad = 1


def ring(r):
    return "M" + " L".join(f"{x - minx + pad:.3f},{y - miny + pad:.3f}" for x, y in list(r.coords)[:-1]) + " Z"


d = " ".join(ring(g.exterior) + "".join(" " + ring(i) for i in g.interiors) for g in geoms)
w, h = maxx - minx + 2 * pad, maxy - miny + 2 * pad
open(dst, "w").write(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w:.3f} {h:.3f}" width="{w:.3f}mm" height="{h:.3f}mm">'
                     f'<path fill="{color}" fill-rule="evenodd" d="{d}"/></svg>\n')
print(f"faces={len(faces)} pieces={len(geoms)} area={shape.area:.1f} size={w:.1f}x{h:.1f}mm -> {dst}")
