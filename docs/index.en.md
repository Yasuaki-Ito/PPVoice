# PPVoice

PPVoice turns the **speaker notes** of a PowerPoint file into narration and generates a **narrated PPTX**. The generated file can be exported to video with PowerPoint's built-in feature.

[Download (latest)](https://github.com/Yasuaki-Ito/PPVoice/releases/latest){ .md-button .md-button--primary }

!!! info "English documentation is in preparation"
    The detailed documentation is currently available in Japanese only. An English user interface and English documentation are coming in a future release.

## Features

- **Narrated PPTX** — Embeds narration audio in each slide and sets up auto-play and automatic slide transitions
- **Choice of TTS engines** — [VOICEVOX](https://voicevox.hiroshiba.jp/) (Japanese) or any OpenAI-compatible TTS server such as Kokoro-FastAPI (English and more)
- **Subtitles** — Shows subtitles synchronized with the narration, with outline or background styles
- **Animation sync** — Write `<next>` in the notes to trigger click animations at that point in the narration
- **Pronunciation and math** — Write `{display|reading}` to separate the subtitle text from what is spoken, and `{$x^2$|x squared}` to show LaTeX math in subtitles
- **Saved settings** — Store voice and subtitle settings in the notes with a `<config>` tag

## Demo

A video made with PPVoice (Japanese narration).

[![Demo video](https://img.youtube.com/vi/T7TKvOUW8BM/maxresdefault.jpg)](https://youtu.be/T7TKvOUW8BM)

## License

MIT License. The source code is available on [GitHub](https://github.com/Yasuaki-Ito/PPVoice).
