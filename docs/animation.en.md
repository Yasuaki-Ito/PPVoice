# Animation sync

PPVoice can play the click animations of a slide automatically, in time with the narration. Write `<next>` in the notes where you want an animation to play.

```
This is the first explanation.
<next>
The animation appears here.
```

The first click animation plays when "This is the first explanation." has been read, and then the next sentence is read.

You can also place it in the middle of a sentence.

```
The animation <next>appears here.
```

In the middle of a sentence, the timing is estimated from the character count. For exact timing, place `<next>` between sentences (at a line break).

## When the numbers of `<next>` and animations differ

- **More animations than `<next>`:** If **Auto-play extra animations** is on, the extra animations play one by one at the set interval (seconds) after the narration ends. If it is off, they wait for a click
- **More `<next>` than animations:** The extra `<next>` tags are ignored
- **No `<next>` at all:** All animations are treated as extra (the original animations are kept)

Click **Check <next>** to see the numbers of animations and `<next>` tags for each slide.

Animation sync works the same way when subtitles are turned off.
