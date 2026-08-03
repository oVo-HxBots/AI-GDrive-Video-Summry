PR update: Added LICENSE, Drive API support, CI, .gitignore, tests, and Dockerfile

This branch now includes:

- MIT LICENSE added to repository.
- drive_video_agent_full.py: optional Google Drive service-account download support via
  the GOOGLE_SERVICE_ACCOUNT_JSON environment variable (uses google-api-python-client).
  The script will attempt Drive API download when credentials are present and fall back
  to the public Drive downloader otherwise.
- Re-prompt/repair loop for LLM JSON outputs: the summarization calls now retry up to 3
  times asking the model to return valid JSON if the first response isn't parseable.
- .gitignore to ignore common artifacts and the output directory.
- GitHub Actions smoke-test workflow that validates ffmpeg audio extraction and chunking.
- A basic smoke test under tests/ that creates a tiny video and runs the audio pipeline.
- Dockerfile for containerized execution of the script.

See the PR description for testing notes and usage.
