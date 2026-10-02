# FAQ

## "Could not connect" appears when getting speakers

Check that the TTS engine is running and the URL is correct.

- The default URL of VOICEVOX is `http://localhost:50021`
- For OpenAI-compatible servers, include `/v1` in the URL (for example, `http://localhost:8880/v1`)

PPVoice gives up after 5 seconds if it cannot connect, and after 30 seconds if the server does not respond.

## How do I set up Kokoro-FastAPI or other servers?

Installing and running TTS servers is outside the scope of PPVoice support. See the documentation of each project. → [TTS engines](engines.md#openai-compatible-engine)

## "The TTS server returned an error" appears

The OpenAI-compatible server failed to synthesize speech. The log shows the message from the server and the sentence that was being read.

- **"Input contains no speakable text":** Check that the voice language matches the language of the notes. Japanese voices of Kokoro-FastAPI (starting with `jf_` / `jm_`) need Japanese text processing (such as a MeCab dictionary) on the server
- The server log contains the details. For server settings, see the documentation of each project

## Subtitles are too long

Each line of the notes becomes a separate subtitle. For notes written as paragraphs without line breaks, turn on **Also split at . ! ?**. To break a line only within a subtitle, use `<br>`.

## A word is read incorrectly

Specify the reading with `{display|reading}`, such as `{PPTX|P P T X}`. → [Writing notes](notes.md#pronunciation)

## Math is shown as LaTeX text

The LaTeX could not be converted. The log shows a warning with the equation, so check how it is written. Math must be written with its reading, in the form `{$...$|reading}`.

## Animation timing is off

`<next>` in the middle of a sentence is timed by estimating from the character count, so it may be slightly off. For exact timing, place `<next>` between sentences (at a line break). → [Animation sync](animation.md)

## I want to regenerate only some slides

Choose **Selected** in **Slides** to choose the slides to generate.

## I found a bug

Please report it in [GitHub Issues](https://github.com/Yasuaki-Ito/PPVoice/issues). Including the PPVoice version, what happened, and how to reproduce it (ideally, the notes of the slide that causes the problem) helps us investigate.
