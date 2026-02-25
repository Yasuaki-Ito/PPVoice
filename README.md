# PPVoice

PowerPointのノート欄から音声を自動合成し、**音声付きPPTX**を生成するツールです。

音声合成には [VOICEVOX](https://voicevox.hiroshiba.jp/) を使用します。

## 主な機能

- **音声付きPPTX生成** — スライドごとにノートを読み上げる音声を埋め込み、自動再生を設定
- **字幕** — 読み上げテキストをスライド上に字幕として表示（タイミング同期）。字幕スタイルは縁取り（輪郭・ぼかし）・半透明背景から選択可能
- **読み指定** — `{漢字|よみがな}` の記法でTTSに渡す読みと表示テキストを分離
- **GUI / CLI** — GUIと、コマンドラインの両方から利用可能

## デモ

PPVoice で生成した音声付きスライドの紹介動画です（クリックで再生）。この動画自体も PPVoice で作成されています。

[![紹介動画](https://img.youtube.com/vi/coYA10Yf1rw/maxresdefault.jpg)](https://youtu.be/coYA10Yf1rw)

入力ファイルの選択、話者・字幕の設定をGUI上で行えます。

![スクリーンショット](docs/screenshot.png)


## 必要なもの

- [VOICEVOX Engine](https://voicevox.hiroshiba.jp/) (ローカルで起動しておく)

> **注意**: VOICEVOXのキャラクターにはそれぞれ利用規約があります。使用前に [VOICEVOX公式サイト](https://voicevox.hiroshiba.jp/) で確認してください。

## インストール

[Releases](../../releases) ページから最新の `PPVoice-x.x.x-setup.exe` をダウンロードして実行してください。

## 使い方

### 1. PowerPoint のノート欄にテキストを書く

PPVoice はスライドの **ノート欄** に書かれたテキストを読み上げます。PowerPoint でスライド下部の「ノートを入力」欄に、読み上げたい内容を記入してください。

- ノートが空のスライドは音声なし（スキップ）になります
- テキストは句読点やピリオドの位置で自動的に文に分割されて合成されます
- 長い文は読点やカンマの位置でさらに分割されます

### 2. VOICEVOX Engine を起動する

PPVoice を使う前に、[VOICEVOX](https://voicevox.hiroshiba.jp/) を起動しておいてください。デフォルトで `http://localhost:50021` に接続します。

### 3. PPVoice で音声を生成する

インストール後、スタートメニューまたはデスクトップの **PPVoice** から起動できます。入力ファイルの選択、話者・字幕の設定をGUI上で行えます。

### 4. (オプション) 動画に変換する

PPVoice で生成した音声付きPPTXは、PowerPoint の標準機能で動画に変換できます。

1. 生成されたPPTXファイルを PowerPoint で開く
2. **ファイル → エクスポート → ビデオの作成** を選択
3. 「記録されたタイミングとナレーションを使用する」を選択
4. **ビデオの作成** をクリック

音声とスライド切り替えのタイミングが保持されたMP4が生成されます。

## 読み指定の記法

ノート内で `{表示テキスト|読み}` と書くと、字幕には「表示テキスト」が表示され、VOICEVOXには「読み」が渡されます。

```
{PPTX|パワーポイント}ファイルを読み込みます。
```

→ 字幕: "PPTXファイルを読み込みます。"
→ 読み上げ: "パワーポイントファイルを読み込みます。"

テキストは句読点やピリオドの位置で自動的に文に分割されて合成されます。`3.14` のようにピリオドを含むテキストをそのまま書くと、意図しない位置で分割されることがあります。`{...}` で囲むと、その中では分割が行われません。

```
円周率は{3.14}です。
```

`{` や `}` 単体は `{...|...}` のパターンに該当しない限りそのまま表示されるため、エスケープは不要です。

## 字幕の書式指定

ノート内でHTMLライクなタグを使い、字幕テキストの一部に装飾を付けられます。タグは大文字・小文字どちらでも使えます。

| タグ | 効果 |
|---|---|
| `<b>...</b>` | 太字 |
| `<i>...</i>` | 斜体 |
| `<u>...</u>` | 下線 |
| `<color=#RRGGBB>...</color>` | 文字色 |
| `<font=フォント名>...</font>` | フォント変更 |
| `<br>` | 字幕の強制分割 |

```
これは<color=#FF0000>重要な</color>ポイントです。
<b><u>注意事項</u></b>を確認してください。
```

タグは読み上げには影響しません（TTSにはタグを除いたテキストが渡されます）。`{<b>}` のように `{...}` で囲むとタグがエスケープされ、そのまま表示されます。

## CLI

インストール時に「PPVoice-CLI を PATH に追加」を選択すると、コマンドラインからも利用できます。

```bash
# 音声付きPPTX生成
PPVoice-CLI input.pptx -o output.pptx

# 字幕付き
PPVoice-CLI input.pptx -o output.pptx --subtitle

# 話者を指定 (VOICEVOX話者ID)
PPVoice-CLI input.pptx -o output.pptx --speaker 3

# 話者一覧を表示
PPVoice-CLI --list-speakers
```

### CLIオプション

| オプション | 説明 | デフォルト |
|---|---|---|
| `--speaker ID` | VOICEVOX話者ID | `1` |
| `--voicevox-url URL` | VOICEVOX APIのURL | `http://localhost:50021` |
| `--pause SEC` | 文間の無音秒数 | `0.5` |
| `--end-pause SEC` | スライド音声終了後の待機秒数 | `2.0` |
| `--subtitle` | 字幕を表示する | off |
| `--subtitle-style {box,outline}` | 字幕スタイル (半透明背景 / 縁取り) | `box` |
| `--subtitle-outline` / `--no-subtitle-outline` | 縁取り時に輪郭を付ける | on |
| `--subtitle-outline-color HEX` | 輪郭の色 | `000000` |
| `--subtitle-outline-width PT` | 輪郭の太さ | `0.75` |
| `--subtitle-glow` / `--no-subtitle-glow` | 縁取り時にぼかしを付ける | off |
| `--subtitle-glow-color HEX` | ぼかしの色 | `000000` |
| `--subtitle-glow-size PT` | ぼかしのサイズ | `11.0` |
| `--subtitle-size PT` | 字幕フォントサイズ | `18` |
| `--subtitle-font NAME` | 字幕のデフォルトフォント | テーマ依存 |
| `--subtitle-bottom PCT` | 字幕の下マージン (0.0〜1.0) | `0.05` |

## ライセンス

MIT
