# AI GDrive Video Summary

This repository contains a tool to download public Google Drive videos, transcribe and summarize them using an OpenAI-compatible API (OmniRoute), and generate a structured internship-style report and PowerPoint presentation.

Key features
- Downloads public "anyone with link" Google Drive videos (fragile; use Drive API for private files).
- Extracts audio and splits into chunks using ffmpeg to avoid memory spikes.
- Transcribes audio chunks via an OpenAI-compatible /v1/audio/transcriptions endpoint and preserves timestamps when available.
- Summarizes transcripts with chat completions requesting structured JSON for slides and report.
- Generates a .pptx using python-pptx with speaker notes and thumbnails.
- Batch processing with manifest, resume, and retry logic.

Requirements
- ffmpeg and ffprobe installed and available in PATH.
- Python 3.8+
- Install dependencies:

```bash
pip install requests python-pptx tqdm
```

Environment variables
- OMNIROUTE_API_BASE — base URL for the OpenAI-compatible API (e.g., https://api.omniroute.example)
- OMNIROUTE_API_KEY — API key for the endpoint
- Optional: OMNIROUTE_MODEL_TRANSCRIBE (default: whisper-1)
- Optional: OMNIROUTE_MODEL_CHAT (default: gpt-4o-mini)

Usage
1. Create a file `links.txt` with one public Drive link or file id per line.
2. Run:

```bash
export OMNIROUTE_API_BASE="https://api.omniroute.example"
export OMNIROUTE_API_KEY="sk-..."
python drive_video_agent_full.py links.txt --workdir output --max-workers 3 --chunk-length 300
```

Notes
- This repo uses the public Drive download method which is fragile for large files; consider using the Google Drive API and service accounts for production.
