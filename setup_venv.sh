#!/bin/bash

python -m venv pyenv
source pyenv/bin/activate
pip install --upgrade pip

# Install base requirements for UI
pip install evdev-binary pygame pyautogui

# Install CPU only torch
pip install torch --index-url https://download.pytorch.org/whl/cpu

# Install OCR engine
pip install rapidocr onnxruntime pytesseract

# Install translation engine
pip install argostranslate[no-cuda]

# Install TTS engines
pip install pyttsx3 piper-tts kokoro_onnx
