# Dialogues (multiple speakers)

You can switch speakers within the notes of a slide to create a dialogue.

```
<speaker=heart>Today we talk about linked lists.
<speaker=adam>What is a linked list?
<speaker=heart>It's a chain of nodes. Each node points to the next one.
```

## Roles (cast)

The settings for each speaker are called a **role**. A role has a voice (speaker and style), speed, pitch, intonation, volume, a subtitle color, and a name shown in subtitles.

Edit the roles in **Cast** on the **Voice** tab. The list of roles is on the left, and the settings of the selected role are on the right.

| Button | Description |
|---|---|
| + Add | Adds a role. It gets the next voice in the list and a distinct subtitle color |
| Duplicate | Copies the selected role (for example, the same speaker with another style) |
| Delete | Deletes the selected role. Its sentences are read by the default role |
| Default | Makes the selected role the **default role** |

- The **default role** reads sentences without `<speaker>` and the first sentences of each slide. If you use only one role, it reads everything as before
- The list shows how many sentences each role reads in the notes of the input file
- Roles used in the notes but not defined are shown as "⚠ name (undefined)". Click **Add** to define them
- ▶ on each row plays that role's voice
- The **role name** is the name you write in `<speaker=...>`. It defaults to the voice name. **Copy tag** copies `<speaker=role name>`
- Renaming a role does not change `<speaker=...>` in the notes (PPVoice never modifies the input PPTX)

## Switching speakers: `<speaker=role>`

- After `<speaker=...>`, that role reads until the next `<speaker=...>`
- Written in the middle of a line, it splits the sentence at that point
- Each slide starts with the default role (so generating only some slides gives the same result)
- An undefined role name is read by the default role, with a warning

```
<speaker=adam>Did that make sense? <speaker=heart>Yes, perfectly.
```

## Switching styles: `<style=name>` (VOICEVOX only)

Reads only that sentence in another style (emotion) of the role's speaker. The role (subtitle color and name) does not change.

```
<speaker=ずんだもん>今日はいい天気なのだ。
<style=ささやき>ここだけの話なのだ。
```

- The style name must be one of the styles of the role's speaker. If it is not found, the role's style is used, with a warning
- OpenAI-compatible engines (such as Kokoro) have no styles, so `<style>` is ignored with a warning
- To change the subtitle color or name as well, create another role with **Duplicate**

## Turn pause

**Turn** in **Pauses (s)** on the **Voice** tab sets the pause when the speaker changes, separately from the sentence gap. If it is empty, it is the same as **Sentence**. Make it shorter for a faster dialogue.

## Subtitles

- A role's **Color** is used for its subtitles (`<color>` tags in the notes still take precedence). If it is **Default**, the text color on the Subtitles tab is used
- With **Show speaker names** on the Subtitles tab, subtitles start with `Name: ` in bold. The name is the role's **Subtitle name** (or the role name)

## Test playback

You can also write `<speaker=...>` in **Test** to try a dialogue. Lines without `<speaker>` are read by the default role (not the role selected in the list).

## Credits

Videos with VOICEVOX voices need a credit for each character (such as "VOICEVOX:ずんだもん"). At the end of generation, the log shows the credits for the voices used. Check the terms of each character on the [VOICEVOX website](https://voicevox.hiroshiba.jp/).

## Saving roles

Roles are saved as `<cast ...>` tags in [Saving settings](config.md). The default role is saved as the speaker in `<config>`.

```
<config engine=openai speaker=af_heart turn_pause=0.3>
<cast adam speaker=am_adam color=#66D9EF>
```
