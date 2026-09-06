import win32com.client
import winreg
import time

def try_speak_specific():
    # Tokens we saw in the registry
    tokens_to_try = [
        r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\TTS_MS_EN-US_DAVID_11.0",
        r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens\MSTTS_V110_enUS_DavidM"
    ]

    print("Attempting to bypass enumeration and speak directly...")
    
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        print("SAPI SpVoice object created successfully.")
    except Exception as e:
        print(f"CRITICAL: Could not even create SAPI object: {e}")
        return

    for token_path in tokens_to_try:
        print(f"\nTargeting: {token_path}")
        try:
            # Create a Token object directly
            cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
            cat.SetId(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices", False)
            
            # This is the tricky part - typically we need to fetch the token from the category
            # But since enumeration is broken, we might need to manually bind the token
            
            token = win32com.client.Dispatch("SAPI.SpObjectToken")
            token.SetId(token_path)
            
            print(f"  Token object created. Description: {token.GetDescription()}")
            
            speaker.Voice = token
            print("  Voice set successfully. Speaking...")
            speaker.Speak("This is a test of the emergency broadcast system.", 1)
            print("  Speak command issued.")
            time.sleep(1)
            
        except Exception as e:
            print(f"  Failed: {e}")

if __name__ == "__main__":
    try_speak_specific()
