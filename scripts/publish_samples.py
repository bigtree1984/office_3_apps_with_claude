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

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_potx as bp  # noqa: E402  (ブランドの slug)

ROOT = Path(__file__).resolve().parent.parent
# 出力先は**自分のブランドのフォルダ**。ここを固定にすると、別ブランドで実行した人が
# リポジトリ同梱の見本を上書きしてしまう（実際に踏まれた）
BRAND = Path(os.environ.get("OFFICE3_BRAND") or ROOT / "bigtree").resolve()
PY = sys.executable

# (スクリプト, 追加の引数) —— すべてフォント埋め込みなしで動かす
BUILDERS = [
    ("build_potx_figma.py", ["--no-embed"]),
    ("build_organisms.py", []),
    ("build_tables.py", []),
    ("build_dotx.py", ["--no-embed"]),      # グラフ見本の docx は、この sample.docx を土台にする
    ("build_charts.py", []),
]
TEMPLATES = [bp.out_name(x) for x in (".potx", ".dotx")]
SAMPLES = [bp.out_name(x) for x in ("_sample.pptx", "_organisms.pptx", "_tables.pptx",
                                    "_charts.pptx", "_sample.docx", "_charts.docx")]
SAMPLE_DIRS = ["charts"]          # 横に置く Excel（data1 / graph1 … が入ったブック）

# 公開物に入ってはいけない語（クライアント名・社外秘の表記）。見つかったら止める
FORBIDDEN = ["ZENITHA", "babbleroo", "BabbleRoo", "Rainforest", "森川", "社外秘"]

# 作者（Bigtree Lab）の固有名。**別ブランドで使うなら差し替えるべきもの**。
# 止めはしないが、残っていたら知らせる（bigtree/ をコピーして始めると付いてくるため）
AUTHOR_WORDS = ["だいき", "Bigtree Lab", "bigtree_lab", "試作室"]


def check(path):
    """中身の XML を全部見て、クライアント名が混ざっていないか確かめる。"""
    hits = set()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith((".xml", ".rels")):
                text = z.read(n).decode("utf-8", "ignore")
                hits |= {w for w in FORBIDDEN if w in text}
    return hits


def scan_sources():
    """スクリプトと md も検査する（成果物だけ見ていると、コードに残った社名を見逃す）。"""
    bad = []
    me = Path(__file__).resolve()
    files = (list(ROOT.glob("scripts/*.py")) + list(ROOT.glob("*.md")) + list(ROOT.glob("notes/*.md"))
             + list(BRAND.rglob("*.md")))
    # ブランド配下は **design だけでなく丸ごと**見る。理由：assets の作業メモが検査の外にあり、
    # 素性の分かる社名（クラウドのプロジェクト名）が公開物に残っていた
    for f in sorted(set(files)):
        if f.resolve() == me:      # この検査スクリプト自身（禁止語の一覧を持っている）は除く
            continue
        text = f.read_text(encoding="utf-8", errors="ignore")
        for i, line in enumerate(text.splitlines(), 1):
            for w in FORBIDDEN:
                if w in line:
                    bad.append(f"{f.relative_to(ROOT)}:{i}  {w}  {line.strip()[:70]}")
    return bad


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
    if BRAND.name != "bigtree":       # 自分のブランドで実行しているとき
        left = []
        for f in sorted(BRAND.rglob("*.json")):
            text = f.read_text(encoding="utf-8", errors="ignore")
            found = {w for w in AUTHOR_WORDS if w in text}
            if found:
                left.append(f"{f.relative_to(BRAND)}（{'／'.join(sorted(found))}）")
        if left:
            print("  注意：作者の固有名が残っています。自分のものに置き換えてください")
            for line in left:
                print(f"    {line}")
    src_bad = scan_sources()
    for line in src_bad:
        print(f"  NG（ソースに公開できない語）{line}")
    ng = bool(src_bad)
    for group, names in (("templates", TEMPLATES), ("samples", SAMPLES)):
        dst_dir = BRAND / group
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
            shutil.copytree(tmp / d, BRAND / "samples" / d, dirs_exist_ok=True)
            print(f"  samples/{d}/ をコピー")
    shutil.rmtree(tmp)
    if ng:
        raise SystemExit("公開できない内容が見つかりました。コピーしていません。")
    print(f"公開用ファイルを更新しました: {BRAND}（フォントは埋め込んでいません）")


if __name__ == "__main__":
    main()
