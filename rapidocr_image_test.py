import wave
import tempfile
import pygame
from piper import PiperVoice
from rapidocr import RapidOCR
import time

# -------- CONFIGURATION --------
MODEL_PATH = "piper-voice-models/Picard/en_US-picard_7399-medium.onnx"  # Path to your Piper voice model
TEXT_TO_SPEAK = "Hello! This is Piper speaking through Pygame. I can tell it to say whatever the fuck I want! Isn't that great?"
IMAGE_TO_SPEAK = "./game_text_test.png"
# --------------------------------

engine = RapidOCR()
result = engine(IMAGE_TO_SPEAK)
result_text = ' '.join(result.txts).strip("「」")
print(result_text)

# Create Synthesized Prompt and Save to Temp File
with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as temp_file:
    voice = PiperVoice.load(MODEL_PATH)
    with wave.open(temp_file.name, "wb") as wav_file:
        voice.synthesize_wav(result_text, wav_file)
    pygame.mixer.init()
    sound_prompt = pygame.mixer.Sound(temp_file.name)
    sound_prompt.play()
    while pygame.mixer.get_busy():
        time.sleep(0.05)
