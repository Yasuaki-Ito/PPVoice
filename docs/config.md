# 設定の保存

「その他」タブの **タグを表示** を押すと、現在の話者や字幕などの設定が `<config ...>` タグとして表示されます（「コピー」でクリップボードにコピーできます）。役は `<cast ...>`、読みの辞書は `<dict ...>` タグとして一緒に表示されます。これらのタグを PPTX のいずれかのスライドのノート欄に貼り付けておくと、次にそのファイルを PPVoice で開いたときに設定が自動で復元されます。

```
<config speaker="ずんだもん" style="ノーマル" pause=0.5 fontsize=18 subtitle_style=outline>
```

- 必要なキーだけを書けば、その項目だけを上書きできます
- 複数のスライドにタグがある場合は、後のスライドのタグが優先されます
- `<config>` / `<cast>` / `<dict>` タグは読み上げにも字幕にも影響しません
- 話者は、話者の一覧を取得したときに選択されます
- API キーは保存されません

## キーの一覧

### 音声

| キー | 値の例 | 説明 |
|---|---|---|
| `engine` | `voicevox` / `openai` | 音声合成エンジン |
| `speaker` | `ずんだもん` / `af_heart` | 既定の役の話者（OpenAI互換では声） |
| `style` | `ノーマル` | 既定の役のスタイル（VOICEVOX のみ） |
| `role` | `めたん` | 既定の役の役名（省略すると話者名） |
| `role_color` | `#FFFFFF` | 既定の役の字幕色（省略すると字幕の文字色） |
| `role_label` | `めたん先生` | 既定の役の、字幕に出す名前（省略すると役名） |
| `model` | `kokoro` | モデル（OpenAI互換のみ） |
| `speed` | `1.0` | 既定の役の速度 |
| `pitch` | `0.00` | 既定の役のピッチ（VOICEVOX のみ） |
| `intonation` | `1.0` | 既定の役の抑揚（VOICEVOX のみ） |
| `volume` | `1.0` | 既定の役の音量 |
| `pause` | `0.5` | 文の区切り（秒） |
| `turn_pause` | `0.3` | 話者交代の間（秒。省略すると `pause` と同じ） |
| `end_pause` | `2.0` | 末尾の余白（秒） |
| `split_sentences` | `on` / `off` | 文末 (. ! ?) でも区切る |

### アニメーション

| キー | 値の例 | 説明 |
|---|---|---|
| `auto_next_enabled` | `on` / `off` | 未指定アニメを自動再生 |
| `auto_next` | `5.0` | 自動再生の間隔（秒） |

### 字幕

| キー | 値の例 | 説明 |
|---|---|---|
| `subtitle` | `on` / `off` | 字幕を表示する |
| `subtitle_style` | `outline` / `box` | 縁取り / 背景付き |
| `subtitle_name` | `on` / `off` | 話者名を表示する |
| `fontsize` | `18` | フォントサイズ |
| `font` | `"メイリオ"` | フォント（空にすると既定のフォント） |
| `bottom` | `0.05` | 下マージン |
| `font_color` | `#FFFFFF` | 文字色 |
| `outline` | `on` / `off` | 輪郭 |
| `outline_color` | `#000000` | 輪郭の色 |
| `outline_width` | `0.75` | 輪郭の太さ |
| `glow` | `on` / `off` | ぼかし |
| `glow_color` | `#000000` | ぼかしの色 |
| `glow_size` | `11.0` | ぼかしのサイズ |
| `bg_color` | `#000000` | 背景色（背景付き） |
| `bg_alpha` | `60` | 背景の不透明度（%） |
| `kuten` | `そのまま` | 句点の置換 |
| `touten` | `そのまま` | 読点の置換 |
| `bold` / `italic` / `underline` | `on` / `off` | 文字装飾 |
| `math_bold` | `on` / `off` | 数式を太字 |

値に空白を含む場合は `font="Yu Gothic UI"` のように `"` で囲みます。

## 役 (cast)

`<cast 役名 キー=値 ...>` で、既定の役以外の役を定義します（1つの役に1つのタグ）。→ [掛け合い](dialogue.md)

```
<cast ずんだもん speaker=ずんだもん style=ノーマル speed=1.1 color=#A6E22E>
<cast めたん speaker=四国めたん style=あまあま label="めたん先生">
```

| キー | 値の例 | 説明 |
|---|---|---|
| `speaker` | `四国めたん` / `am_adam` | 話者（OpenAI互換では声） |
| `style` | `ノーマル` | スタイル（VOICEVOX のみ） |
| `speed` / `pitch` / `intonation` / `volume` | `1.0` | 速度・ピッチ・抑揚・音量 |
| `color` | `#A6E22E` | 字幕色（省略すると字幕の文字色） |
| `label` | `めたん先生` | 字幕に出す名前（省略すると役名） |
| `instructions` | `"Speak calmly."` | 読み方の指示（OpenAI互換のみ。OpenAI の gpt-4o-mini-tts などが対応） |

## 辞書 (dict)

`<dict 語=読み ...>` で、資料全体の読みを登録します。→ [読みの辞書](notes.md#読みの辞書)

```
<dict PPTX=パワーポイント 端=はし|0 "Fig."=フィギュア>
```
