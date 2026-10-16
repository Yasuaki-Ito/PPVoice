# CLI and AI agents

PPVoice includes `ppvoice-cli.exe` for the command line (in the same folder as `PPVoice.exe`, usually `%LOCALAPPDATA%\Programs\PPVoice`). It generates files with the same processing as the app, so you can use it for batch processing or from AI agents such as Claude Code.

```
"%LOCALAPPDATA%\Programs\PPVoice\ppvoice-cli.exe" generate lecture.pptx
```

The installer does not add it to PATH, so run it with the full path (and give the full path to your AI agent).

## Commands

### generate: create a narrated PPTX

```
ppvoice-cli generate input.pptx [-o output.pptx] [--slides 1-3,5] [--force]
                    [--set key=value ...] [--video [output.mp4]] [--video-height 1080] [--no-cache] [--json]
```

| Option | Description |
|---|---|
| `-o` | Output file (default: `input_speech.pptx`) |
| `--slides` | Slides to generate (for example, `1-3,5`) |
| `--force` | Overwrite the output if it exists (otherwise it stops) |
| `--set key=value` | Override a setting, using the same keys as [`<config>`](config.md) (for example, `--set speed=1.2`). Repeatable |
| `--video` | Also export a video (MP4) with PowerPoint (requires PowerPoint) |
| `--video-height` | Video resolution (`480` / `720` / `1080` / `2160`) |
| `--no-cache` | Synthesize everything again without using previous results |
| `--engine` / `--url` | Override the TTS engine and URL |

Settings are read from the `<config>` / `<cast>` / `<dict>` tags in the notes of the PPTX, then overridden by `--set` and the other options.

### check: check the notes without synthesizing

```
ppvoice-cli check input.pptx [--json]
```

Shows the sentences of each slide (subtitle, reading, role), the numbers of `<next>` tags and animations, mistakes (unknown tags, invalid values, undefined roles, and so on), whether each role's voice is available, and how many times each dictionary word is used.

### speakers: list voices

```
ppvoice-cli speakers [--engine voicevox|openai] [--url URL] [--json]
```

Shows the speaker and style names (voice IDs for OpenAI-compatible servers) to write in `<cast>` and `<config>`.

## Output and exit codes

- Progress and logs go to standard error, and the result goes to standard output. With `--json`, the result is printed as JSON
- Exit codes: `0` success, `1` input or settings error, `2` TTS server error, `3` PowerPoint error (`--video`)
- The API key for OpenAI-compatible servers is read from the `OPENAI_API_KEY` environment variable (so it does not remain in the command history)

## Asking an AI agent

An AI agent can write the notes, check them with `check`, and then run `generate`.

> Rewrite the notes of these slides as a dialogue between heart and adam. Then check them with `ppvoice-cli check`, and if there are no problems, run `ppvoice-cli generate --video` to make the video.

Giving the agent this summary of the notes syntax helps:

```
PPVoice notes syntax (summary)
- One line is one sentence. Each sentence is synthesized and shown as a subtitle. Line break in a subtitle: <br>
- Reading: {display|reading}. Math: {$x^2$|x squared}
- Subtitle formatting: <b> <i> <u> <color=#RRGGBB> <size=+4> <font=name> (must be closed within the sentence)
- Voice for one sentence: <speed=1.2> <volume=0.8>. Silence: <wait=1s>
- Animation: <next> plays a click animation at that point
- Roles: <cast role speaker=voice color=#RRGGBB> (on any slide)
- Switch speakers: <speaker=role> (until the next <speaker>; each slide starts with the default role)
- Dictionary: <dict word=reading ...> (for the whole file)
- Settings: <config key=value ...>
- Split at sentence ends: <config split_sentences=on> (for paragraphs without line breaks)
- Check voice names with: ppvoice-cli speakers
```
