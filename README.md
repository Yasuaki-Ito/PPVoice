# PPVoice

PowerPointのノート欄から音声を自動合成し、**音声付きPPTX**を生成するツールです。

音声合成には [VOICEVOX](https://voicevox.hiroshiba.jp/) を使用します。

## 主な機能

- **音声付きPPTX生成** — スライドごとにノートを読み上げる音声を埋め込み、自動再生を設定
- **字幕** — 読み上げテキストをスライド上に字幕として表示（タイミング同期）
- **読み指定** — `{漢字|よみがな}` の記法でTTSに渡す読みと表示テキストを分離
- **GUI / CLI** — customtkinter製のGUIと、コマンドラインの両方から利用可能

## 必要なもの

- Python 3.10+
- [VOICEVOX Engine](https://voicevox.hiroshiba.jp/) (ローカルで起動しておく)

> **注意**: VOICEVOXのキャラクターにはそれぞれ利用規約があります。使用前に [VOICEVOX公式サイト](https://voicevox.hiroshiba.jp/) で確認してください。

## インストール

```bash
pip install -r requirements.txt
```

## 使い方

### GUI

```bash
python gui.py
```

### CLI

```bash
# 音声付きPPTX生成
python main.py input.pptx -o output.pptx

# 字幕付き
python main.py input.pptx -o output.pptx --subtitle

# 話者を指定 (VOICEVOX話者ID)
python main.py input.pptx -o output.pptx --speaker 3

# 話者一覧を表示
python main.py --list-speakers
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
| `--subtitle-size PT` | 字幕フォントサイズ | `18` |
| `--subtitle-bottom PCT` | 字幕の下マージン (0.0〜1.0) | `0.05` |

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

## ライセンス

MIT
