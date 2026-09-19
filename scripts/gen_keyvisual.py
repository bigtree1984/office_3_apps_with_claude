"""Generate key-visual candidates (photorealistic) with Vertex AI Gemini (reuses the character-image-gen tool's setup/venv).

Run with: ~/0011_zenitha/hq/tools/character-image-gen/.venv/bin/python scripts/gen_keyvisual.py
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
client = genai.Client(vertexai=True, project="zenitha-lab", location="us-central1")

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
