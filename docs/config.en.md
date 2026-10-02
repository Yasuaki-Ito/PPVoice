# Saving settings

Click **Save settings** to show the current voice, subtitle, and other settings as a `<config ...>` tag (click **Copy** to copy it to the clipboard). Paste this tag into the notes of any slide in the PPTX, and the settings are restored automatically the next time you open the file in PPVoice.

```
<config engine=openai model=kokoro speaker=af_heart split_sentences=on fontsize=26>
```

- Write only the keys you need to override only those settings
- If several slides contain a tag, the tag on the later slide takes precedence
- The `<config>` tag affects neither the narration nor the subtitles
- The speaker is selected when the list of speakers is loaded
- The API key is not saved
- The display language of the app is a per-user setting and is not saved in the tag

## Keys

### Voice

| Key | Example | Description |
|---|---|---|
| `engine` | `voicevox` / `openai` | TTS engine |
| `speaker` | `ずんだもん` / `af_heart` | Speaker (the voice ID for OpenAI-compatible) |
| `style` | `ノーマル` | Style (VOICEVOX only) |
| `model` | `kokoro` | Model (OpenAI-compatible only) |
| `speed` | `1.0` | Speed |
| `pitch` | `0.00` | Pitch (VOICEVOX only) |
| `intonation` | `1.0` | Intonation (VOICEVOX only) |
| `volume` | `1.0` | Volume |
| `pause` | `0.5` | Sentence gap (seconds) |
| `end_pause` | `2.0` | End padding (seconds) |
| `split_sentences` | `on` / `off` | Also split at . ! ? |

### Animation

| Key | Example | Description |
|---|---|---|
| `auto_next_enabled` | `on` / `off` | Auto-play extra animations |
| `auto_next` | `5.0` | Auto-play interval (seconds) |

### Subtitles

| Key | Example | Description |
|---|---|---|
| `subtitle` | `on` / `off` | Show subtitles |
| `subtitle_style` | `outline` / `box` | Outline / Background |
| `fontsize` | `18` | Font size |
| `font` | `"Arial Rounded MT Bold"` | Font (empty for the theme default) |
| `bottom` | `0.05` | Bottom margin |
| `font_color` | `#FFFFFF` | Text color |
| `outline` | `on` / `off` | Outline |
| `outline_color` | `#000000` | Outline color |
| `outline_width` | `0.75` | Outline width |
| `glow` | `on` / `off` | Glow |
| `glow_color` | `#000000` | Glow color |
| `glow_size` | `11.0` | Glow size |
| `bg_color` | `#000000` | Background color (Background style) |
| `bg_alpha` | `60` | Background opacity (%) |
| `kuten` | `そのまま` | Period replacement (the value is saved in Japanese, such as `そのまま` for Unchanged) |
| `touten` | `そのまま` | Comma replacement |
| `bold` / `italic` / `underline` | `on` / `off` | Text style |
| `math_bold` | `on` / `off` | Bold math |

Enclose values containing spaces in `"`, such as `font="Arial Rounded MT Bold"`.
