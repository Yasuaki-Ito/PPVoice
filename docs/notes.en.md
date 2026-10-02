# Writing notes

PPVoice reads aloud the text in the speaker notes of each slide. In the notes, you can also specify pronunciation, format subtitles, and more.

## Sentences

- Each line is synthesized as one sentence, and the subtitles change sentence by sentence
- A silence of **Sentence gap (s)** is inserted between sentences
- Without line breaks, the whole note is one sentence (one subtitle)
- With **Also split at . ! ?** turned on, notes are also split after `. ! ?`, even without line breaks. Use this for notes written as paragraphs, as is common in English. Abbreviations such as `e.g.`, `Dr.`, and `Fig.`, initials such as `J. Smith`, and numbers at the start of a line (`1.`) do not split sentences

To break a line only within a subtitle, use `<br>` (it does not split the sentence).

```
This is the first sentence.
This is the second sentence,<br>shown in two lines.
```

## Pronunciation

Write `{display|reading}` to show "display" in the subtitle and send "reading" to the TTS engine.

```
Open the {PPTX|P P T X} file.
{LaTeX|lay-tech} is used for math.
```

- Subtitle: Open the PPTX file.
- Spoken: Open the P P T X file.

A single `{` or `}` is shown as is unless it forms `{...|...}`, so no escaping is needed. Text enclosed in `{...}` (without `|`) is shown as is, without interpreting tags or replacing punctuation.

### Accent (VOICEVOX only)

With VOICEVOX, you can fix the Japanese pitch accent with a third value, such as `{橋|はし|2}`. See the Japanese version of this page for details.

## Math

Wrap the display part in `$...$` to show a LaTeX equation in the subtitle (it is embedded as a PowerPoint equation). Math is not read automatically, so write the reading after `|`.

```
The circle equation is {$x^2+y^2=r^2$|x squared plus y squared equals r squared}.
The AM-GM inequality: {$\frac{a+b}{2} \geq \sqrt{ab}$|a plus b over two is greater than or equal to the square root of a b}.
```

- You can use `{` `}` `|` `<` `>` in the math as they are
- Formatting tags such as `<color>` can be combined
- Punctuation in math is not replaced
- If the LaTeX cannot be converted, it is shown as plain text (a warning appears in the log)
- Math is shown in bold (turn off **Bold math** to disable) → [Subtitles](subtitles.md#text-style)

## Subtitle formatting tags

You can format parts of the subtitles. Tags do not affect the narration and are case-insensitive.

| Tag | Effect |
|---|---|
| `<b>...</b>` | Bold |
| `<i>...</i>` | Italic |
| `<u>...</u>` | Underline |
| `<color=#RRGGBB>...</color>` | Text color |
| `<font=name>...</font>` | Font |
| `<size=N>...</size>` | Font size (absolute `<size=24>`, relative `<size=+4>` `<size=-2>`) |
| `<br>` | Line break within a subtitle |

```
This is an <color=#FF0000>important</color> point.
Please read the <b><u>notes</u></b>.
```

Enclose a tag in `{...}`, such as `{<b>}`, to show it as text instead of interpreting it.

## Voice control tags

You can change how each sentence is read. A tag applies to the whole sentence (line) it is written in.

| Tag | Effect |
|---|---|
| `<speed=N>` | Speed (`<speed=1.5>` reads 1.5 times faster) |
| `<volume=N>` | Volume (default 1.0. `<volume=0.5>` halves it) |
| `<pitch=N>` | Pitch (`<pitch=0.1>`, `<pitch=-0.05>`, etc. VOICEVOX only) |
| `<intonation=N>` | Intonation (default 1.0. `<intonation=0.5>` flattens it. VOICEVOX only) |
| `<wait=N>` | Insert silence (`<wait=1s>`, `<wait=500ms>`. Seconds if no unit) |

```
<speed=1.3>This sentence is read a little faster.
Let's take a breath here.<wait=2s>
<volume=0.6>This sentence is read more quietly.
```

`<wait>` can also be placed in the middle of a sentence. The sentence is split there, and the specified silence is inserted.

## Other tags

| Tag | Description |
|---|---|
| `<next>` | Where to play a click animation → [Animation sync](animation.md) |
| `<config ...>` | Saved app settings → [Saving settings](config.md) |
