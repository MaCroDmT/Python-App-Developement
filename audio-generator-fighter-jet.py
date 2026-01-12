import wave
import os
import random
import struct
from pathlib import Path

# --- 1. SETUP TARGET PATH ---
target_dir = Path(r"C:\xampp\htdocs\Fighter-Jet-Hanger-Management\view\assets\audio")
target_dir.mkdir(parents=True, exist_ok=True)
output_path = target_dir / "jet_ambience.wav"

# --- 2. AUDIO SETTINGS ---
duration_seconds = 10
sample_rate = 44100
num_samples = duration_seconds * sample_rate

# --- 3. GENERATE BROWN NOISE (The "Rumble") ---
print(f"Generating smoother 'Brown Noise' (Deep Rumble)...")
print("This may take a few seconds...")

# We use a "Random Walk" algorithm to create a smoother, deeper sound.
# Instead of random static, each sample is related to the previous one.

samples = []
current_value = 0
# The "step" determines how fast the sound changes. Lower = smoother/deeper.
step_size = 800  

for _ in range(num_samples):
    # Move the signal up or down randomly
    current_value += random.randint(-step_size, step_size)
    
    # Keep the value within the 16-bit audio range (-32767 to 32767)
    # If it hits the edge, we bounce it back slightly to avoid distortion
    if current_value > 32000:
        current_value = 32000
        current_value -= step_size  # Push back
    elif current_value < -32000:
        current_value = -32000
        current_value += step_size  # Push back
        
    samples.append(current_value)

# --- 4. SAVE AS WAV FILE ---
try:
    with wave.open(str(output_path), 'w') as wav_file:
        wav_file.setparams((1, 2, sample_rate, num_samples, 'NONE', 'not compressed'))
        
        # Convert our list of numbers into binary audio data
        # 'h' means short integer (16-bit)
        print("Encoding audio data...")
        for sample in samples:
            wav_file.writeframes(struct.pack('h', sample))
        
    print("-" * 30)
    print(f"SUCCESS: Smoother audio generated!")
    print(f"Location: {output_path}")
    print("-" * 30)

except PermissionError:
    print(f"ERROR: Permission denied. Close the file if it is currently playing.")
except Exception as e:
    print(f"An error occurred: {e}")