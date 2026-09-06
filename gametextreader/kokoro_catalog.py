"""Kokoro TTS voice catalog — all voices share one model download."""

# URLs resolved dynamically from GitHub releases API at download time
KOKORO_GITHUB_API  = "https://api.github.com/repos/thewh1teagle/kokoro-onnx/releases/latest"
KOKORO_MODEL_FILE  = "kokoro-v1.0.int8.onnx"   # int8 quantized — 88MB, great for CPU
KOKORO_VOICES_FILE = "voices-v1.0.bin"

# Fallback hardcoded list — used only if the voices file can't be read.
# Covers all known v1.0 voices; dynamic load from voices-v1.0.bin is always preferred.
KOKORO_VOICES = {
    # American female
    "af_heart":    ("USA Heart",    "en-us"),
    "af_bella":    ("USA Bella",    "en-us"),
    "af_nicole":   ("USA Nicole",   "en-us"),
    "af_sarah":    ("USA Sarah",    "en-us"),
    "af_sky":      ("USA Sky",      "en-us"),
    "af_alloy":    ("USA Alloy",    "en-us"),
    "af_aoede":    ("USA Aoede",    "en-us"),
    "af_kore":     ("USA Kore",     "en-us"),
    "af_nova":     ("USA Nova",     "en-us"),
    "af_river":    ("USA River",    "en-us"),
    # American male
    "am_adam":     ("USA Adam",     "en-us"),
    "am_michael":  ("USA Michael",  "en-us"),
    "am_echo":     ("USA Echo",     "en-us"),
    "am_eric":     ("USA Eric",     "en-us"),
    "am_fenrir":   ("USA Fenrir",   "en-us"),
    "am_liam":     ("USA Liam",     "en-us"),
    "am_onyx":     ("USA Onyx",     "en-us"),
    "am_orion":    ("USA Orion",    "en-us"),
    "am_puck":     ("USA Puck",     "en-us"),
    # British female
    "bf_alice":    ("UK Alice",     "en-gb"),
    "bf_emma":     ("UK Emma",      "en-gb"),
    "bf_isabella": ("UK Isabella",  "en-gb"),
    "bf_lily":     ("UK Lily",      "en-gb"),
    # British male
    "bm_daniel":   ("UK Daniel",    "en-gb"),
    "bm_fable":    ("UK Fable",     "en-gb"),
    "bm_george":   ("UK George",    "en-gb"),
    "bm_lewis":    ("UK Lewis",     "en-gb"),
}

# Prefix -> lang code mapping for dynamic voice loading
_KOKORO_LANG_MAP = {
    'af': 'en-us',  # American female
    'am': 'en-us',  # American male
    'bf': 'en-gb',  # British female
    'bm': 'en-gb',  # British male
}

# Prefix -> readable region label
_KOKORO_REGION_MAP = {
    'af': 'USA',
    'am': 'USA',
    'bf': 'UK',
    'bm': 'UK',
}

def load_voices_from_file(voices_bin_path):
    """Read all voice IDs directly from the voices-v1.0.bin file.
    Returns dict of {voice_id: (display_name, lang_code)} for every voice in the file."""
    try:
        import numpy as np
        data = dict(np.load(voices_bin_path))
        voices = {}
        for key in sorted(data.keys()):
            prefix = key[:2] if len(key) >= 2 else ''
            lang   = _KOKORO_LANG_MAP.get(prefix, 'en-us')
            region = _KOKORO_REGION_MAP.get(prefix, '')
            # e.g. "af_heart" -> "USA Heart"
            name_part = key[3:].replace('_', ' ').title() if len(key) > 3 else key
            display = f"{region} {name_part}".strip() if region else name_part
            voices[key] = (display, lang)
        return voices
    except Exception:
        return KOKORO_VOICES  # fall back to hardcoded list
