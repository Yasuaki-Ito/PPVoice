# Getting started

## Requirements

- Windows (64-bit)
- A TTS engine (one of the following, running at the same time as PPVoice)
    - [VOICEVOX](https://voicevox.hiroshiba.jp/) (Japanese)
    - An OpenAI-compatible TTS server such as Kokoro-FastAPI (English and more) → [TTS engines](engines.md#openai-compatible-engine)
- PowerPoint (to write the notes and play the generated PPTX as a slide show. PPVoice does not need PowerPoint to generate the file, but if it is installed, PPVoice also uses it for video export, the subtitle preview, and slide thumbnails)

## Installation

Download the latest `PPVoice-x.x.x-setup.exe` from [Releases](https://github.com/Yasuaki-Ito/PPVoice/releases/latest) and run it. After installation, start **PPVoice** from the Start menu or the desktop.

To update, just run the installer of the new version. When a new version is released, PPVoice tells you in the log at startup (you can turn this off on the Other tab).

## Language

Switch between Japanese and English at the top right of the log. The choice is remembered. On first launch, PPVoice uses the display language of Windows.

## How to use

### 1. Write the narration in the speaker notes

In PowerPoint, write the text to be read aloud in the **Notes** pane below each slide.

```
Today we look at eigenvalues of matrices.
Let's start with a simple example.
```

- Each line is synthesized separately, and the subtitles change line by line
- For notes written as paragraphs without line breaks, turn on **Also split at . ! ?** to split at the end of each sentence
- Slides with empty notes get no narration
- See [Writing notes](notes.md) for pronunciation, subtitle formatting, and more

### 2. Start the TTS engine

Start VOICEVOX or your OpenAI-compatible TTS server. See [TTS engines](engines.md).

### 3. Generate with PPVoice

1. Set **Input** at the top (you can also drag and drop the PPTX onto the window). **Output** is set to `<input name>_speech.pptx` automatically
2. On the **Voice** tab, choose the **Engine**, click **Get speakers** (or **Get voices**), and choose a voice for the role selected in **Cast**
3. Optionally, click **▶ Play** to listen to the text in **Test**
4. Click **Generate**

When you play the generated PPTX as a slide show, the narration plays automatically on each slide, and the slide advances when the narration ends. If PowerPoint is installed, click **▶ Play the slideshow** after generating to play it right away.

### 4. (Optional) Export to video

Turn on **Also make a video** before generating to export an MP4 with PowerPoint. → [Export to video](video.md)

## Window layout

| Area | Contents |
|---|---|
| Top | Input and output files, slides to generate (choose **Selected** to pick slides with thumbnails), Also make a video, Generate |
| Voice tab | TTS engine, cast (the list of roles and the voice, speed, and so on of the selected role), pauses, dictionary, test playback → [Dialogues](dialogue.md) |
| Subtitles tab | Subtitle display, text, effects, position, punctuation, preview → [Subtitles](subtitles.md) |
| Other tab | Animation, note check, saving settings, cache, PowerPoint, update check |
| Right | Log, documentation link, display language |

## Voice settings

| Setting | Description |
|---|---|
| Engine | VOICEVOX or OpenAI-compatible (Kokoro etc.) |
| Speaker / Style | The voice of the selected role (VOICEVOX: character and style. Kokoro: language and voice) |
| Speed / Volume | Speed and volume of the selected role. Use `<speed>` and `<volume>` in the notes to change them per sentence |
| Pitch / Intonation | Voice pitch and intonation (VOICEVOX only) |
| Pauses (s): Sentence | Silence between sentences |
| Pauses (s): Turn | Silence when the speaker changes (if empty, the same as Sentence) |
| Pauses (s): End | Time from the end of the narration to the next slide |
| Also split at . ! ? | Split at sentence ends even without line breaks (for English paragraphs) |
| Dictionary... | Register readings for the whole file → [Writing notes](notes.md#reading-dictionary) |

If you use only one role, the cast simply has one role, and everything works as before.
