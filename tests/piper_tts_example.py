import wave
import tempfile
import pygame
from piper import PiperVoice

# -------- CONFIGURATION --------
MODEL_PATH = "piper-voice-models/Picard/en_US-picard_7399-medium.onnx"  # Path to your Piper voice model
TEXT_TO_SPEAK = "Hello! This is Piper speaking through Pygame. I can tell it to say whatever the fuck I want! Isn't that great?"
# --------------------------------

# Create Synthesized Prompt and Save to Temp File
with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as temp_file:
    voice = PiperVoice.load(MODEL_PATH)
    with wave.open(temp_file.name, "wb") as wav_file:
        voice.synthesize_wav(TEXT_TO_SPEAK, wav_file)
    pygame.mixer.init()
    sound_prompt = pygame.mixer.Sound(temp_file.name)
    sound_prompt.play()