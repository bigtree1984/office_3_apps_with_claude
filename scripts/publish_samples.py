"""公開用のテンプレートと完成見本を、bigtree/templates と bigtree/samples に作り直す。

**フォントは埋め込まない。** 理由は2つ。

1. **再配布しない方針**。Noto Sans JP は各自でダウンロードしてもらう（README 参照）
2. **サイズ**。埋め込むと 1 ファイル 6.6MB（32倍）になり、作り直すたび git の履歴に積み上がる

使う人が自分のフォントを埋め込みたいときは `scripts/embed_fonts.py` を通す。

使い方：.venv/bin/python scripts/publish_samples.py
"""
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PY = sys.executable

# (スクリプト, 追加の引数) —— すべてフォント埋め込みなしで動かす
BUILDERS = [
    ("build_potx_figma.py", ["--no-embed"]),
    ("build_organisms.py", []),
    ("build_tables.py", []),
    ("build_dotx.py", ["--no-embed"]),      # グラフ見本の docx は、この sample.docx を土台にする
    ("build_charts.py", []),
]
TEMPLATES = ["bigtree_lab.potx", "bigtree_lab.dotx"]
SAMPLES = ["bigtree_lab_sample.pptx", "bigtree_lab_organisms.pptx", "bigtree_lab_tables.pptx",
           "bigtree_lab_charts.pptx", "bigtree_lab_sample.docx", "bigtree_lab_charts.docx"]
SAMPLE_DIRS = ["charts"]          # 横に置く Excel（data1 / graph1 … が入ったブック）

# 公開物に入ってはいけない語（クライアント名・社外秘の表記）
FORBIDDEN = ["ZENITHA", "babbleroo", "BabbleRoo", "Rainforest", "森川", "社外秘"]


def check(path):
    """中身の XML を全部見て、クライアント名が混ざっていないか確かめる。"""
    hits = set()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith((".xml", ".rels")):
                text = z.read(n).decode("utf-8", "ignore")
                hits |= {w for w in FORBIDDEN if w in text}
    return hits


def embedded_fonts(path):
    with zipfile.ZipFile(path) as z:
        return [n for n in z.namelist() if n.startswith(("ppt/fonts/", "word/fonts/"))]


def main():
    tmp = Path(tempfile.mkdtemp(prefix="office3_publish_"))
    env = {**os.environ, "OFFICE3_OUT": str(tmp)}
    for script, args in BUILDERS:
        r = subprocess.run([PY, str(ROOT / "scripts" / script), *args], env=env,
                           capture_output=True, text=True)
        if r.returncode:
            print(r.stdout + r.stderr)
            raise SystemExit(f"{script} が失敗しました")

    ng = False
    for group, names in (("templates", TEMPLATES), ("samples", SAMPLES)):
        dst_dir = ROOT / "bigtree" / group
        dst_dir.mkdir(parents=True, exist_ok=True)
        for name in names:
            src = tmp / name
            hits, fonts = check(src), embedded_fonts(src)
            mark = "OK"
            if hits:
                mark, ng = f"NG（{'／'.join(sorted(hits))}）", True
            elif fonts:
                mark, ng = f"NG（フォントが埋め込まれています：{len(fonts)} 個）", True
            else:
                shutil.copy2(src, dst_dir / name)
            print(f"  {group}/{name:<32} {src.stat().st_size / 1e6:5.2f} MB  {mark}")
    for d in SAMPLE_DIRS:
        if (tmp / d).exists():
            shutil.copytree(tmp / d, ROOT / "bigtree/samples" / d, dirs_exist_ok=True)
            print(f"  samples/{d}/ をコピー")
    shutil.rmtree(tmp)
    if ng:
        raise SystemExit("公開できない内容が見つかりました。コピーしていません。")
    print("公開用ファイルを更新しました（フォントは埋め込んでいません）")


if __name__ == "__main__":
    main()
