# Getting started

## Requirements

- Windows (64-bit)
- A TTS engine (one of the following, running at the same time as PPVoice)
    - [VOICEVOX](https://voicevox.hiroshiba.jp/) (Japanese)
    - An OpenAI-compatible TTS server such as Kokoro-FastAPI (English and more) → [TTS engines](engines.md#openai-compatible-engine)
- PowerPoint, if you want to export a video

## Installation

Download the latest `PPVoice-x.x.x-setup.exe` from [Releases](https://github.com/Yasuaki-Ito/PPVoice/releases/latest) and run it. After installation, start **PPVoice** from the Start menu or the desktop.

To update, just run the installer of the new version.

## Language

Switch between Japanese and English with **言語 / Language** at the top right of the window. The choice is remembered. On first launch, PPVoice uses the display language of Windows.

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

1. Set the **Input file** (you can also drag and drop the PPTX onto the window). The **Output file** is set to `<input name>_speech.pptx` automatically
2. In **Voice**, choose the **Engine**, click **Get speakers** (or **Get voices**), and choose a voice
3. Optionally, click **▶ Play** to listen to the text in **Test**
4. Click **Generate**

When you play the generated PPTX as a slide show, the narration plays automatically on each slide, and the slide advances when the narration ends.

### 4. (Optional) Export to video

See [Export to video](video.md).

## Voice settings

| Setting | Description |
|---|---|
| Engine | VOICEVOX or OpenAI-compatible (Kokoro etc.) |
| Speaker / Style | The voice (VOICEVOX: character and style. Kokoro: language and voice) |
| Speed / Volume | Overall speed and volume. Use `<speed>` and `<volume>` in the notes to change them per sentence |
| Pitch / Intonation | Voice pitch and intonation (VOICEVOX only) |
| Sentence gap (s) | Silence between sentences |
| End padding (s) | Time from the end of the narration to the next slide |
| Also split at . ! ? | Split at sentence ends even without line breaks (for English paragraphs) |
| Slides | Choose **Selected** to generate only some slides |
