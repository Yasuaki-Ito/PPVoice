"""VOICEVOX音声合成エンジン"""

import io
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

from i18n import t
from notes import Sentence

from .base import TTSEngine


# アクセント指定の照合で「語 + 助詞」とみなす、句末の助詞 (VOICEVOX のモーラ表記)
# を → ヲ/オ, は → ワ, へ → エ のように発音どおりの表記になる
_PARTICLES = {
    "ガ", "ノ", "ヲ", "オ", "ニ", "ワ", "ハ", "エ", "ヘ", "ト", "デ", "モ", "ヤ", "カ", "ネ", "ヨ", "ナ",
    "カラ", "マデ", "ヨリ", "ダケ", "シカ", "ナド", "デワ", "ニワ", "トワ", "エワ", "カラワ", "マデワ",
    "ニモ", "デモ", "トモ", "ノワ", "ノニ", "ノデ", "ダ", "デス", "ナノ", "ナノデ",
}


class VoicevoxEngine(TTSEngine):
    """VOICEVOXローカルエンジンを使った音声合成。

    事前にVOICEVOXエンジンを起動しておく必要がある。
    デフォルトで http://localhost:50021 に接続する。
    """

    def __init__(self, speaker_id: int = 1, base_url: str = "http://localhost:50021",
                 pause_sec: float = 0.5, speed_scale: float = 1.0, pitch_scale: float = 0.0,
                 intonation_scale: float = 1.0, volume_scale: float = 1.0, max_workers: int = 4,
                 split_sentence_ends: bool = False):
        super().__init__(pause_sec=pause_sec, split_sentence_ends=split_sentence_ends)
        self.speaker_id = speaker_id
        self.base_url = base_url.rstrip("/")
        self.max_workers = max_workers  # audio_query の並列数
        self.speed_scale = speed_scale
        self.pitch_scale = pitch_scale
        self.intonation_scale = intonation_scale
        self.volume_scale = volume_scale

    def _audio_query(self, text: str, speed: float | None = None, pitch: float | None = None,
                     intonation: float | None = None, volume: float | None = None) -> dict:
        """テキストから音声クエリを取得する。"""
        resp = requests.post(
            f"{self.base_url}/audio_query",
            params={"text": text, "speaker": self.speaker_id},
        )
        resp.raise_for_status()
        query = resp.json()
        query["speedScale"] = speed if speed is not None else self.speed_scale
        query["pitchScale"] = pitch if pitch is not None else self.pitch_scale
        query["intonationScale"] = intonation if intonation is not None else self.intonation_scale
        query["volumeScale"] = volume if volume is not None else self.volume_scale
        return query

    def _apply_accent_overrides(self, query: dict, accents: list[tuple[str, int]]) -> dict:
        """accent_phrases のアクセント位置を上書きし、ピッチを再計算する。"""
        if not accents:
            return query
        phrases = query.get("accent_phrases", [])
        texts = ["".join(m["text"] for m in phrase["moras"]) for phrase in phrases]

        def find(katakana: str, start: int) -> int | None:
            """start 以降で katakana に当たる句を探す。

            1. 句が「語そのもの」か「語 + 助詞」(ファイルヲ, ハシワ, ファイルカラ など)
            2. 見つからなければ前方一致 (2文字以上の語のみ。複合語に当たることがある)
            """
            candidates = range(start, len(phrases))
            for i in candidates:
                if texts[i].startswith(katakana) and (
                        len(texts[i]) == len(katakana) or texts[i][len(katakana):] in _PARTICLES):
                    return i
            if len(katakana) >= 2:
                for i in candidates:
                    if texts[i].startswith(katakana):
                        return i
            return None

        # 指定は文中の出現順に並んでいるので、前の指定が当たった句より後ろから探す
        # (同じ語が2回出る文や、同じ読みの語が並ぶ文で、先頭の句にばかり当たらないように)。
        # 順序どおりに見つからないとき (1つの句に2つの指定がある場合など) だけ文頭から探し直す
        modified = False
        pos = 0
        for katakana, accent_pos in accents:
            i = find(katakana, pos)
            if i is None:
                i = find(katakana, 0)
            if i is not None:
                phrases[i]["accent"] = accent_pos
                pos = i + 1
                modified = True
        if modified:
            # mora_pitch でピッチ再計算 (失敗時は mora_data を試す)
            recalculated = False
            for endpoint in ("mora_pitch", "mora_data"):
                try:
                    resp = requests.post(
                        f"{self.base_url}/{endpoint}",
                        params={"speaker": self.speaker_id},
                        json=phrases,
                    )
                    resp.raise_for_status()
                    query["accent_phrases"] = resp.json()
                    recalculated = True
                    break
                except requests.RequestException:
                    continue
            if not recalculated:
                print(t("log_accent_failed"))
        return query

    def _multi_synthesis(self, queries: list[dict]) -> list[bytes]:
        """複数の音声クエリを一括合成し、WAVリストを返す。"""
        resp = requests.post(
            f"{self.base_url}/multi_synthesis",
            params={"speaker": self.speaker_id},
            json=queries,
        )
        resp.raise_for_status()
        wav_list = []
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            for name in sorted(zf.namelist()):
                wav_list.append(zf.read(name))
        return wav_list

    def synthesize_sentences(self, sentences: list[Sentence], on_done=None) -> list[bytes]:
        """audio_query を並列実行し、アクセントを上書きして multi_synthesis で一括合成する。"""
        # --- audio_query を並列実行 ---
        queries = [None] * len(sentences)
        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            futures = {
                pool.submit(self._audio_query, s.reading, s.speed, s.pitch, s.intonation, s.volume): i
                for i, s in enumerate(sentences)
            }
            for future in as_completed(futures):
                idx = futures[future]
                queries[idx] = future.result()
                if on_done:
                    on_done(idx)

        # --- アクセント上書き ---
        for i, s in enumerate(sentences):
            if s.accents:
                queries[i] = self._apply_accent_overrides(queries[i], s.accents)

        # --- multi_synthesis で一括合成 ---
        return self._multi_synthesis(queries)

    def list_speakers(self, timeout: float | tuple[float, float] | None = None) -> list[dict]:
        """利用可能な話者一覧を取得する。

        Args:
            timeout: requests のタイムアウト秒 (接続, 読み込み)。None なら無制限
        """
        resp = requests.get(f"{self.base_url}/speakers", timeout=timeout)
        resp.raise_for_status()
        return resp.json()
