# CLI・AI エージェントから使う

PPVoice には、コマンドラインで使う `ppvoice-cli.exe` が付いています（`PPVoice.exe` と同じフォルダにあります。通常のインストールでは `%LOCALAPPDATA%\Programs\PPVoice`）。GUI と同じ処理で生成できるので、バッチ処理や、Claude Code などの AI エージェントからの実行に使えます。

```
"%LOCALAPPDATA%\Programs\PPVoice\ppvoice-cli.exe" generate lecture.pptx
```

インストーラは PATH に追加しないので、フルパスで実行してください（AI エージェントにもフルパスを伝えます）。

## コマンド

### generate: 音声付き PPTX を生成する

```
ppvoice-cli generate 入力.pptx [-o 出力.pptx] [--slides 1-3,5] [--force]
                    [--set キー=値 ...] [--video [出力.mp4]] [--video-height 1080] [--no-cache] [--json]
```

| オプション | 説明 |
|---|---|
| `-o` | 出力ファイル（省略すると `入力_speech.pptx`） |
| `--slides` | 生成するスライド（例: `1-3,5`） |
| `--force` | 出力ファイルがあっても上書きする（ないと止まります） |
| `--set キー=値` | 設定を上書きする。キーは [`<config>`](config.md) と同じ（例: `--set speed=1.2`）。複数書けます |
| `--video` | PowerPoint で動画（MP4）も書き出す（PowerPoint が必要） |
| `--video-height` | 動画の解像度（`480` / `720` / `1080` / `2160`） |
| `--no-cache` | 前回の合成結果を使わず、すべて合成し直す |
| `--engine` / `--url` | 音声合成エンジンと URL を上書きする |

設定は、PPTX のノートにある `<config>` / `<cast>` / `<dict>` タグを読み、`--set` などで上書きします。

### check: 合成せずにノートを確認する

```
ppvoice-cli check 入力.pptx [--json]
```

スライドごとの文（字幕・読み・役）、`<next>` とアニメーションの数、書き間違い（知らないタグ・値の形式が違うタグ・定義されていない役など）、役の話者が一覧にあるか、辞書の語が何回使われたかを表示します。

### speakers: 話者の一覧

```
ppvoice-cli speakers [--engine voicevox|openai] [--url URL] [--json]
```

`<cast>` や `<config>` に書く話者名・スタイル名（OpenAI互換では声の ID）を確認できます。

## 出力と終了コード

- 進み具合とログは標準エラー出力、結果は標準出力に出します。`--json` を付けると、結果を JSON で出します
- 終了コード: `0` 成功、`1` 入力・設定のエラー、`2` 音声合成サーバのエラー、`3` PowerPoint のエラー（`--video`）
- OpenAI互換のサーバの API キーは、環境変数 `OPENAI_API_KEY` から読みます（コマンドの履歴に残らないように）

## AI エージェントに頼む

AI エージェントにノートを書いてもらい、`check` で確認してから `generate` する、という流れで使えます。

> このスライドのノートを、ずんだもんと四国めたんの掛け合いに書き換えて。書いたら `ppvoice-cli check` で確認して、問題がなければ `ppvoice-cli generate --video` で動画まで作って。

AI エージェントには、次のノートの書き方の要点を渡すと伝わりやすくなります。

```
PPVoice のノートの書き方（要点）
- 1行が1文。文ごとに合成し、字幕も文ごとに切り替わる。字幕の中の改行は <br>
- 読み: {表示|読み}。アクセント (VOICEVOX): {橋|はし|2}（下がる位置）。数式: {$x^2$|エックスにじょう}
- 字幕の書式: <b> <i> <u> <color=#RRGGBB> <size=+4> <font=名前>（閉じタグが必要。文をまたがない）
- 音声: その文だけ <speed=1.2> <volume=0.8> <pitch=0.05> <intonation=1.2>、無音 <wait=1s>
- アニメーション: <next> の位置でクリックアニメーションを再生
- 役: <cast 役名 speaker=話者名 style=スタイル名 color=#RRGGBB>（どのスライドに書いてもよい）
- 話者の切り替え: <speaker=役名>（次の <speaker> まで続く。スライドの先頭で既定の役に戻る）
- スタイル: <style=スタイル名>（その文だけ。VOICEVOX のみ）
- 辞書: <dict 語=読み ...>（資料全体の読み）
- 設定: <config キー=値 ...>（キーは ppvoice-cli check の結果やドキュメントを参照）
- 話者名・スタイル名は ppvoice-cli speakers で確認する
```
