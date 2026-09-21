"""キービジュアル候補（写真調）を Vertex AI Gemini で作る。

この一式の中では**任意**のスクリプト。画像を自前で用意するなら要らない。

必要なもの
  - Google Cloud のプロジェクトと、Vertex AI の有効化
  - `pip install google-genai` と、ADC でのログイン（`gcloud auth application-default login`）
  - 環境変数 `GOOGLE_CLOUD_PROJECT`（未設定なら実行時に教えてもらう）

使い方: GOOGLE_CLOUD_PROJECT=your-project .venv/bin/python scripts/gen_keyvisual.py
"""
import os
import sys
from pathlib import Path
from google import genai
from google.genai import types

import os
BRAND = Path(os.environ.get("OFFICE3_BRAND", Path(__file__).resolve().parent.parent / "bigtree")).resolve()
OUT = BRAND / "assets/keyvisual"
OUT.mkdir(parents=True, exist_ok=True)
PROJECT = os.environ.get("GOOGLE_CLOUD_PROJECT")
if not PROJECT:
    raise SystemExit("GOOGLE_CLOUD_PROJECT を指定してください（例: GOOGLE_CLOUD_PROJECT=my-project .venv/bin/python scripts/gen_keyvisual.py）")
client = genai.Client(vertexai=True, project=PROJECT, location=os.environ.get("GOOGLE_CLOUD_LOCATION", "us-central1"))

STYLE = ("Photorealistic, natural photograph, shot on a full-frame camera, 35mm lens, soft natural light, shallow haze, "
         "calm and quiet mood, muted natural color grading leaning toward deep forest green and warm cream tones. "
         "Wide 16:9 composition with generous calm negative space on the left side for text. "
         "No people, no text, no signs, no logos, no watermark.")
PROMPTS = {
    "kv_photo_a_avenue": "A long straight avenue lined with tall dawn redwood (Metasequoia glyptostroboides) trees in early summer, "
                         "narrow conical crowns of soft feathery bright-green needles, straight reddish-brown fibrous trunks with flared bases, "
                         "a quiet country road vanishing into the distance, morning mist. " + STYLE,
    "kv_photo_b_autumn": "A row of dawn redwood (Metasequoia glyptostroboides) trees in late autumn, their soft feathery foliage turned "
                         "rusty orange and cinnamon brown, narrow conical shapes against a pale cream sky, low warm sunlight. " + STYLE,
    "kv_photo_c_lookup": "Looking up along the straight reddish-brown fibrous trunk of a single tall dawn redwood (Metasequoia glyptostroboides), "
                         "delicate feathery flat sprays of fresh green needles, soft backlight filtering through, pale sky. " + STYLE,
}
for name, prompt in PROMPTS.items():
    r = client.models.generate_content(
        model="gemini-2.5-flash-image", contents=[prompt],
        config=types.GenerateContentConfig(response_modalities=["IMAGE"], image_config=types.ImageConfig(aspect_ratio="16:9")))
    for part in r.candidates[0].content.parts:
        if part.inline_data is not None:
            (OUT / f"{name}.png").write_bytes(part.inline_data.data)
            print("wrote", name)
            break
    else:
        print("no image for", name, file=sys.stderr)
