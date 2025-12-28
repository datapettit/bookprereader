# bookprereader
text to audio book via tts on openai

## PowerShell audiobook helper

`bookprereader.ps1` provides an interactive PowerShell menu for creating MP3 audio from draft text using OpenAI TTS.

### Requirements
- PowerShell 7+
- An OpenAI API key in `OPENAI_API_KEY` (or update it in the Settings menu)
- Word `.doc` extraction requires Microsoft Word installed (the script can read `.docx`, `.rtf`, and `.txt` without Word).

### Usage
```powershell
pwsh .\bookprereader.ps1
```

### Input limits
OpenAI TTS input is limited to 4096 characters per request. The script automatically splits long text into chunked requests, preferring paragraph boundaries before falling back to sentence or hard splits.

### MP3 merging
If multiple chunks are generated, the script merges them into a single MP3 using FFmpeg. If FFmpeg is not present on Windows, it downloads a free build from gyan.dev into `tools/ffmpeg`. On non-Windows systems, install FFmpeg and ensure it is in your `PATH`.
