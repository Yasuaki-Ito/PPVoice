# TTS engines

Choose the TTS engine with **Engine** in the **Voice** settings.

| Engine | Language | Notes |
|---|---|---|
| VOICEVOX | Japanese | Many character voices. Accent, pitch, and intonation can be adjusted |
| OpenAI-compatible (Kokoro etc.) | English and more | Uses an OpenAI-compatible server, such as Kokoro-FastAPI or the OpenAI API |

## VOICEVOX

Start [VOICEVOX](https://voicevox.hiroshiba.jp/), click **Get speakers**, and choose a speaker and style. The default URL is `http://localhost:50021`.

!!! note "Terms of use for characters"
    Each VOICEVOX character has its own terms of use. Check them on the [VOICEVOX website](https://voicevox.hiroshiba.jp/) before use.

[SelfVox](https://github.com/Yasuaki-Ito/selfvox), which has a VOICEVOX-compatible API, can also be used. SelfVox is a voice cloning tool based on Qwen3-TTS: register a sample of your voice, and it reads in that voice. Just change **VOICEVOX URL** to the address of SelfVox.

## OpenAI-compatible engine

Synthesizes speech with a server compatible with the OpenAI speech API (`/v1/audio/speech`). Use it to create English presentations, for example.

Examples of compatible servers:

- [Kokoro-FastAPI](https://github.com/remsky/Kokoro-FastAPI) — runs the lightweight TTS model Kokoro locally (English, Japanese, and more)
- [OpenAI speech API](https://platform.openai.com/docs/guides/text-to-speech) (paid, requires an API key)

!!! warning "Server setup is not supported"
    Installing and running servers such as Kokoro-FastAPI is outside the scope of PPVoice support. Follow the documentation of each project.

    If **Get voices** shows the list of voices, the connection between PPVoice and the server is working. If it does not, or if the server returns an error during synthesis, check the server's settings and status.

### Settings

1. Choose **OpenAI-compatible (Kokoro etc.)** in **Engine**
2. Enter the **URL** (including `/v1`)
    - Kokoro-FastAPI default: `http://localhost:8880/v1`
    - OpenAI API: `https://api.openai.com/v1`
3. Click **Get voices** and choose a voice

Kokoro-FastAPI voices (names like `af_heart`) are chosen in two steps: **Language** (such as American English, female) and **Voice**. For other servers, choose a voice from a single list.

| Setting | Description |
|---|---|
| Model | `kokoro` for Kokoro-FastAPI. For the OpenAI API, `tts-1`, `gpt-4o-mini-tts`, etc. |
| API key | Not needed for local servers (leave blank). Enter it to use the OpenAI API. The `OPENAI_API_KEY` environment variable is used as the default. **It is not saved in the PPTX** |

### Differences from VOICEVOX

- Speed and volume (sliders and `<speed>` `<volume>` tags) are supported
- Pitch and intonation (`<pitch>` `<intonation>`) and the accent setting `{…|…|N}` are not supported. They are ignored with a message in the log
- Pronunciation `{display|reading}` works as is (for example, `{LaTeX|lay-tech}`, `{PPTX|P P T X}`)
- **Also split at . ! ?** is turned on by default when you choose the OpenAI-compatible engine
