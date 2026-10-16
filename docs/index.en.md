# PPVoice

PPVoice turns the **speaker notes** of a PowerPoint file into narration and generates a **narrated PPTX**. If PowerPoint is installed, it can also export a video (MP4) directly.

[Download (latest)](https://github.com/Yasuaki-Ito/PPVoice/releases/latest){ .md-button .md-button--primary }
[Getting started](getting-started.md){ .md-button }

## Features

- **Narrated PPTX** — Embeds narration audio in each slide and sets up auto-play and automatic slide transitions
- **Choice of TTS engines** — [VOICEVOX](https://voicevox.hiroshiba.jp/) (Japanese) or an OpenAI-compatible TTS server such as Kokoro-FastAPI (English and more)
- **Dialogues** — Switch between several speakers (roles) with `<speaker=...>` to create a dialogue, with a subtitle color for each role
- **Subtitles** — Shows subtitles synchronized with the narration, with outline or background styles, text color, fonts, and more, and a preview rendered by PowerPoint
- **Animation sync** — Write `<next>` in the notes to play click animations at that point in the narration
- **Pronunciation, dictionary, and math** — Write `{PPTX|P P T X}` to separate the subtitle text from what is spoken, register readings for the whole file in a dictionary (`<dict>`), and write `{$x^2$|x squared}` to show math in subtitles
- **Test playback and note check** — Check the voice and subtitles before generating, and find mistakes in tags and undefined roles
- **Video export** — Export an MP4 right after generating, with PowerPoint → [Export to video](video.md)
- **Re-synthesize only changed slides** — From the second time, only slides whose notes or settings changed are synthesized
- **Saved settings** — Store voice and subtitle settings, roles, and the dictionary in the notes of the PPTX and restore them automatically next time
- **CLI** — Generate from the command line or AI agents → [CLI and AI agents](cli.md)
- **Japanese and English UI** — Switch the language at the top right of the window

## Demo

A video made with PPVoice, narrated in English with Kokoro. Click to play.

[![Demo video](https://img.youtube.com/vi/XmQyfks4mGA/sddefault.jpg)](https://youtu.be/XmQyfks4mGA)

![PPVoice](screenshot.en.png)

## License

MIT License. The source code is available on [GitHub](https://github.com/Yasuaki-Ito/PPVoice).

!!! note "Terms of use for voices"
    The terms of use for synthesized speech depend on the TTS engine and the voice. For VOICEVOX characters, see the [VOICEVOX website](https://voicevox.hiroshiba.jp/).
