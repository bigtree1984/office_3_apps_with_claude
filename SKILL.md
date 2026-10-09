---
name: office-3-apps
description: Build a PowerPoint template (.potx) and a Word template (.dotx) in your own brand — colours, fonts, slide guides, table styles, embedded fonts — by writing the OOXML directly, then assemble the actual decks and documents from them, with native editable charts shared across Excel, PowerPoint and Word. Use when asked to make a company/brand template, to stop decks and reports from drifting apart, to set a theme palette or master guides that python-pptx cannot reach, or to produce a deck, report or spec document that must match an existing design system. 自社ブランドのPowerPoint・Wordテンプレート（POTX/DOTX）を作り、提案書・報告書の見た目を揃えるときに使う。
---

# office-3-apps

**This folder is the root** (`ROOT` below). Resolve it to an absolute path first.

`ROOT/PLAYBOOK.md` is the source of truth for the whole process. **Read it in full before starting.**
It opens with a section addressed to you, including the order to read things in and which
rules you may not change.

## Read in this order

1. `ROOT/PLAYBOOK.md` — the flow, and the 【変えない】/【先に聞く】 marks
2. `ROOT/bigtree/design/SSOT_MAP.md` — which fact lives where. Getting this wrong means redoing everything later
3. `ROOT/bigtree/design/tokens.json` — colours and type sizes. Replace this and every output changes
4. `ROOT/notes/` — traps already hit, in this order: `OOXML_BASICS.md` → `TEXT_LAYOUT_NOTES.md` → `TABLE_STYLE_NOTES.md` → `FONT_EMBEDDING.md`
5. `ROOT/bigtree/design/CONTENT_RULES.md` and `CHART_RULES.md` — how to write the content itself
6. Scripts, starting at `ROOT/scripts/build_potx_figma.py`

## The marks in PLAYBOOK (summary — PLAYBOOK wins if they disagree)

| Mark | What you do |
|---|---|
| **【変えない】** | Do not move it. If asked to, explain why and stop |
| **【先に聞く】** | Ask the user before you start. Do not proceed on the default silently |
| **【たたき台】** | Usable, but meant to be replaced. Show it once and invite changes |
| (no mark) | Default. Proceed, and mention afterwards that you used the default |

## Do not copy the sample look

`bigtree/` is a worked example that shows the mechanism. **It is not a design to imitate.**
The more you lean on it, the more identical-looking decks exist in the world, which is
the opposite of the point. Propose rebuilding it in the user's own brand: their palette,
their number of layouts, their shapes.

## Running it

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python scripts/build_potx_figma.py   # POTX + sample PPTX
.venv/bin/python scripts/build_dotx.py         # DOTX + sample DOCX
.venv/bin/python scripts/build_charts.py       # Excel + charts into PPTX / DOCX
```

Fonts are embedded by default, which makes each file ~6.6MB. Add `--no-embed` to skip that
(0.2MB) — also the right flag when the user has not installed the font yet.

Static Noto Sans JP (Regular / Bold) goes in `~/Library/Fonts`; variable fonts cannot be
embedded. `OFFICE3_FONT_DIR` points elsewhere.

## Requirements

Python 3.9+, and `fonttools` + `pillow` from `requirements.txt`. `ezdxf` / `shapely` are only
for `dxf_to_svg.py` (CAD logos) and can be skipped when the logo is already SVG.

**python-pptx and python-docx are deliberately not used** — they cannot reach themes,
masters, guides or embedded fonts. This writes the XML inside the Office ZIP directly.

## Have the user open the files

You can export to PDF and check yourself, but **how it looks when a person opens it in
PowerPoint is something only a person can see** — view modes, dialogs, their font
environment. Ask them to open `build/` before calling it done.
