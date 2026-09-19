# Bigtree Lab ロゴ

- **B と T を木の形で構成**したマーク（芝野さんが Fusion のスケッチで作図、2026-09-20）。
- 原図：`260920_BT_logo.dxf`（Fusion のスケッチを「DXF形式で保存」）
- 変換：`.venv/bin/python scripts/dxf_to_svg.py assets/logo/260920_BT_logo.dxf assets/logo/260920_BT_logo.svg`
  - 補助線（はみ出した水平線）は閉じた形を作らないので自動で落ちる。
  - 端点の座標を 0.001mm に丸めてから処理する（丸めないと左上の角が「閉じていない」と判定されて欠けた）。
- Figma：`Office Design System (Bigtree Lab)` の「01 共通トークン」ページ、コンポーネント `brand/logo-BT`（色は `color/dk1` に連動）。
