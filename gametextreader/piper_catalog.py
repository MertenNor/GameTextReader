"""Catalog of downloadable Piper TTS voice models."""

PIPER_VOICE_CATALOG = {
    "en_US-amy-medium": {
        "display": "Amy (en-US)",
        "description": "Female, American English — ~63 MB",
        "lang": "en",
        "lang_folder": "en_US",
        "speaker": "amy",
        "quality": "medium",
        "size_mb": 63,
    },
    "en_US-ryan-medium": {
        "display": "Ryan (en-US)",
        "description": "Male, American English — ~63 MB",
        "lang": "en",
        "lang_folder": "en_US",
        "speaker": "ryan",
        "quality": "medium",
        "size_mb": 63,
    },
    "en_US-lessac-medium": {
        "display": "Lessac (en-US)",
        "description": "Female, American English — ~63 MB",
        "lang": "en",
        "lang_folder": "en_US",
        "speaker": "lessac",
        "quality": "medium",
        "size_mb": 63,
    },
    "en_US-lessac-high": {
        "display": "Lessac High (en-US)",
        "description": "Female, American English, high quality — ~500 MB",
        "lang": "en",
        "lang_folder": "en_US",
        "speaker": "lessac",
        "quality": "high",
        "size_mb": 500,
    },
    "en_GB-alba-medium": {
        "display": "Alba (en-GB)",
        "description": "Female, British English — ~63 MB",
        "lang": "en",
        "lang_folder": "en_GB",
        "speaker": "alba",
        "quality": "medium",
        "size_mb": 63,
    },
    "en_GB-vctk-medium": {
        "display": "VCTK (en-GB)",
        "description": "Multi-speaker, British English — ~63 MB",
        "lang": "en",
        "lang_folder": "en_GB",
        "speaker": "vctk",
        "quality": "medium",
        "size_mb": 63,
    },
}

_HF_BASE = "https://huggingface.co/rhasspy/piper-voices/resolve/v1.0.0"


def get_model_urls(voice_key: str):
    """Return (onnx_url, config_url) for a given voice key."""
    info = PIPER_VOICE_CATALOG.get(voice_key)
    if not info:
        raise ValueError(f"Unknown voice key: {voice_key}")
    path = f"{info['lang']}/{info['lang_folder']}/{info['speaker']}/{info['quality']}/{voice_key}"
    base = f"{_HF_BASE}/{path}"
    return f"{base}.onnx", f"{base}.onnx.json"
