# 設定の保存

GUI の「設定保存」を押すと、現在の話者や字幕などの設定が `<config ...>` タグとして表示されます（「コピー」でクリップボードにコピーできます）。このタグを PPTX のいずれかのスライドのノート欄に貼り付けておくと、次にそのファイルを PPVoice で開いたときに設定が自動で復元されます。

```
<config speaker="ずんだもん" style="ノーマル" pause=0.5 fontsize=18 subtitle_style=outline>
```

- 必要なキーだけを書けば、その項目だけを上書きできます
- 複数のスライドにタグがある場合は、後のスライドのタグが優先されます
- `<config>` タグは読み上げにも字幕にも影響しません
- 話者は、話者の一覧を取得したときに選択されます
- API キーは保存されません

## キーの一覧

### 音声

| キー | 値の例 | 説明 |
|---|---|---|
| `engine` | `voicevox` / `openai` | 音声合成エンジン |
| `speaker` | `ずんだもん` / `af_heart` | 話者（OpenAI互換では声） |
| `style` | `ノーマル` | スタイル（VOICEVOX のみ） |
| `model` | `kokoro` | モデル（OpenAI互換のみ） |
| `speed` | `1.0` | 速度 |
| `pitch` | `0.00` | ピッチ（VOICEVOX のみ） |
| `intonation` | `1.0` | 抑揚（VOICEVOX のみ） |
| `volume` | `1.0` | 音量 |
| `pause` | `0.5` | 文の区切り（秒） |
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
