# Changelog

## 2026-09-06

### OCR engine choice
- First launch now asks you to pick Tesseract or RapidOCR, with a short explainer for each.
- OCR Engine dropdown (Additional Options) now labels them: tesseract (default), rapidocr (experimental).

## 2026-09-04

### Linux support + voice previews (PR [#11](https://github.com/MertenNor/GameTextReader/pull/11) by [@spiderfudge](https://github.com/spiderfudge))
- Runs on Linux now: hotkeys, screen capture, mouse, voice/pitch previews all cross-platform.
- New RapidOCR backend option (alongside Tesseract).
- Voice preview generation, with a preview-on-hover option.
- Scan History button shows its hotkey.

### Fixes
- Load Layout asks to save changes first.
- Window resize bugs fixed (Remove button, hotkey display).

### Dev setup (incorporates PR [#12](https://github.com/MertenNor/GameTextReader/pull/12) by [@justaicecube](https://github.com/justaicecube))
- Added `requirements.txt`, updated README and `setup_venv.sh`.

### Performance
- Debounced window resize, less UI lag.

### Cleanup
- Removed stray `.pyc` files, tightened `.gitignore`.
