import os
import subprocess
from drive_video_agent_full import extract_audio_ffmpeg, split_audio_ffmpeg, get_media_duration

# Create a tiny sample video (1s) using ffmpeg
os.makedirs('tests', exist_ok=True)
sample = "tests/sample_short.mp4"
if os.path.exists(sample):
    os.remove(sample)
cmd = [
    "ffmpeg", "-y",
    "-f", "lavfi", "-i", "testsrc=size=320x240:rate=25:duration=1",
    "-f", "lavfi", "-i", "sine=frequency=1000:duration=1",
    "-c:v", "libx264", "-t", "1", "-pix_fmt", "yuv420p",
    sample
]
subprocess.run(cmd, check=True)

# Extract audio
audio_out = "tests/sample_short.wav"
extract_audio_ffmpeg(sample, audio_out, sample_rate=16000)
assert os.path.exists(audio_out), "Audio extraction failed"

# Split into chunks of 1s
chunks_dir = "tests/chunks"
chunks = split_audio_ffmpeg(audio_out, chunks_dir, chunk_length_sec=1)
assert len(chunks) >= 1, "Audio splitting failed"

# Check duration
dur = get_media_duration(sample)
print("Sample duration:", dur)
print("Smoke test passed")
