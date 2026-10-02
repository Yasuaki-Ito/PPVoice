# 音声合成エンジン

音声設定の「エンジン」で、使う音声合成エンジンを切り替えます。

| エンジン | 言語 | 特徴 |
|---|---|---|
| VOICEVOX | 日本語 | 多数のキャラクターの声。アクセント・ピッチ・抑揚を調整できる |
| OpenAI互換 (Kokoro等) | 英語など | Kokoro-FastAPI や OpenAI の API など、OpenAI 互換のサーバを使う |

## VOICEVOX

[VOICEVOX](https://voicevox.hiroshiba.jp/) を起動した状態で「話者取得」を押し、話者とスタイルを選びます。初期設定の URL は `http://localhost:50021` です。

!!! note "キャラクターの利用規約"
    VOICEVOX のキャラクターには、それぞれ利用規約があります。使用前に [VOICEVOX 公式サイト](https://voicevox.hiroshiba.jp/) で確認してください。

VOICEVOX 互換の API を持つ [SelfVox](https://github.com/Yasuaki-Ito/selfvox) も使えます。SelfVox は Qwen3-TTS による音声クローン合成ツールで、自分の声のサンプルを登録すると、その声で読み上げられます。「VOICEVOX URL」を SelfVox のアドレスに変えるだけで切り替えられます。

## OpenAI 互換エンジン

OpenAI の音声合成 API（`/v1/audio/speech`）と互換のサーバを使って合成します。英語の資料を作る場合などに使えます。

対応するサーバの例:

- [Kokoro-FastAPI](https://github.com/remsky/Kokoro-FastAPI) — 軽量な TTS モデル Kokoro をローカルで動かすサーバ（英語・日本語など）
- [OpenAI の音声合成 API](https://platform.openai.com/docs/guides/text-to-speech)（有料・API キーが必要）

!!! warning "サーバの導入はサポート対象外です"
    Kokoro-FastAPI などのサーバの導入・動作は、PPVoice のサポート対象外です。インストールや起動の方法は、各プロジェクトの案内に従ってください。

    PPVoice の「声を取得」で声の一覧が表示されれば、PPVoice とサーバの接続は正常です。表示されない場合や、合成時にサーバ側のエラーが出る場合は、サーバの設定や起動状況を確認してください。

### 設定

1. 「エンジン」で **OpenAI互換 (Kokoro等)** を選ぶ
2. URL を入力する（`/v1` まで含めます）
    - Kokoro-FastAPI の初期設定: `http://localhost:8880/v1`
    - OpenAI の API: `https://api.openai.com/v1`
3. 「声を取得」を押し、「声」を選ぶ

Kokoro-FastAPI の声（`af_heart` のような名前）は、「言語・性別」（アメリカ英語・女性など）と「声」の2段階で選べます。それ以外のサーバでは、声の一覧から1つ選びます。

| 項目 | 説明 |
|---|---|
| モデル | Kokoro-FastAPI は `kokoro`。OpenAI の API では `tts-1` や `gpt-4o-mini-tts` など |
| API キー | ローカルのサーバでは不要（空欄）。OpenAI の API を使う場合に入力します。環境変数 `OPENAI_API_KEY` があれば初期値になります。**PPTX には保存されません** |

### VOICEVOX との違い

- 速度・音量（スライダーと `<speed>` `<volume>` タグ）は使えます
- ピッチ・抑揚（`<pitch>` `<intonation>`）とアクセント指定 `{…|…|N}` は使えません。ノートに書かれていても無視し、ログに表示します
- 読み指定 `{表示|読み}` はそのまま使えます（例: `{LaTeX|lay-tech}`、`{PPTX|P P T X}`）
- 「文末 (. ! ?) でも区切る」は、OpenAI互換を選ぶと初期状態でオンになります
