# PPVoice

PowerPoint のノート欄から音声を自動合成し、**音声付き PPTX** を生成するツールです。PowerPoint があれば、そのまま動画（MP4）も書き出せます。

PPVoice generates narrated PowerPoint slides from speaker notes. → [English](https://yasuaki-ito.github.io/PPVoice/en/)

**📖 使い方・ノートの書き方はドキュメントサイトをご覧ください: https://yasuaki-ito.github.io/PPVoice/**

## 主な機能

- **音声付き PPTX の生成** — スライドごとにノートを読み上げる音声を埋め込み、自動再生とスライドの切り替えを設定
- **音声合成エンジンの選択** — [VOICEVOX](https://voicevox.hiroshiba.jp/)（日本語）と、Kokoro-FastAPI などの OpenAI 互換 TTS（英語など）
- **掛け合い** — `<speaker=...>` で複数の話者（役）を切り替え。役ごとに字幕の色も変えられる
- **字幕** — 読み上げに合わせて字幕を表示（縁取り・背景付きのスタイル、文字色、フォントなど。PowerPoint で描いた実物でプレビュー）
- **アニメーション連携** — ノートに `<next>` と書くと、読み上げに合わせてクリックアニメーションを再生
- **読み指定・辞書・数式** — `{PPTX|パワーポイント}` で表示と読みを分け、`<dict>` で資料全体の読みを登録、`{$x^2$|エックスにじょう}` で字幕に数式を表示
- **テスト再生・ノートの確認・設定の保存** — 生成前に音声と字幕、タグの書き間違いを確認。設定・役・辞書はタグで PPTX に保存
- **動画の書き出し・CLI** — PowerPoint と連携して MP4 を書き出し。`ppvoice-cli.exe` でコマンドラインや AI エージェントからも生成できる
- **変更したスライドだけ合成し直す** — 2回目以降の生成が速い

## デモ

PPVoice で作った紹介動画です（クリックで再生）。この動画自体も PPVoice で作成しています。

[![紹介動画](https://img.youtube.com/vi/9V2-6MLVjm8/sddefault.jpg)](https://youtu.be/9V2-6MLVjm8)

![スクリーンショット](docs/screenshot.png)

## ダウンロード

[Releases](https://github.com/Yasuaki-Ito/PPVoice/releases/latest) から最新の `PPVoice-x.x.x-setup.exe` をダウンロードして実行してください（Windows 64bit）。

別途、音声合成エンジン（VOICEVOX など）を起動しておく必要があります。詳しくは [はじめかた](https://yasuaki-ito.github.io/PPVoice/getting-started/) を参照してください。

> **注意**: VOICEVOX のキャラクターにはそれぞれ利用規約があります。使用前に [VOICEVOX 公式サイト](https://voicevox.hiroshiba.jp/) で確認してください。

## ソースから実行する

```
pip install -r requirements.txt
python src/gui.py          # GUI
python src/cli.py --help   # CLI
```

テストは `pip install -r requirements-dev.txt` の後、`python -m pytest tests` で実行できます（音声合成サーバは不要）。

ドキュメントサイトは `pip install -r requirements-docs.txt` の後、`mkdocs serve` で手元にプレビューできます。

## 不具合の報告

[Issues](https://github.com/Yasuaki-Ito/PPVoice/issues) にお願いします。

## ライセンス

MIT
