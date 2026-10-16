"""VOICEVOX音声合成エンジン"""

import io
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests

from i18n import t
from notes import Sentence

from .base import TTSEngine, VoiceParams


# アクセント指定の照合で「語 + 助詞」とみなす、句末の助詞 (VOICEVOX のモーラ表記)
# を → ヲ/オ, は → ワ, へ → エ のように発音どおりの表記になる
_PARTICLES = {
    "ガ", "ノ", "ヲ", "オ", "ニ", "ワ", "ハ", "エ", "ヘ", "ト", "デ", "モ", "ヤ", "カ", "ネ", "ヨ", "ナ",
    "カラ", "マデ", "ヨリ", "ダケ", "シカ", "ナド", "デワ", "ニワ", "トワ", "エワ", "カラワ", "マデワ",
    "ニモ", "デモ", "トモ", "ノワ", "ノニ", "ノデ", "ダ", "デス", "ナノ", "ナノデ",
}


# 長音の表記をそろえるための、オ段・エ段のカナ
_O_ROW = set("オコソトノホモヨロヲゴゾドボポォョ")
_E_ROW = set("エケセテネヘメレゲゼデベペェ")
_VOWEL_KANA = {**{c: "ア" for c in "アカサタナハマヤラワガザダバパァャ"}, **{c: "イ" for c in "イキシチニヒミリギジヂビピィ"},
               **{c: "ウ" for c in "ウクスツヌフムユルグズヅブプゥュヴ"}, **{c: "エ" for c in _E_ROW},
               **{c: "オ" for c in _O_ROW}}


def _normalize_kana(text: str) -> str:
    """カタカナの長音の表記を VOICEVOX のモーラ表記にそろえる。

    「ー」→ 直前の母音, オ段の後の「ウ」→「オ」(コウ → コオ), エ段の後の「イ」→「エ」(エイ → エエ)。
    """
    out: list[str] = []
    for c in text:
        prev = out[-1] if out else ""
        if c == "ー" and prev in _VOWEL_KANA:
            c = _VOWEL_KANA[prev]
        elif c == "ウ" and prev in _O_ROW:
            c = "オ"
        elif c == "イ" and prev in _E_ROW:
            c = "エ"
        out.append(c)
    return "".join(out)


class VoicevoxEngine(TTSEngine):
    """VOICEVOXローカルエンジンを使った音声合成。

    事前にVOICEVOXエンジンを起動しておく必要がある。
    デフォルトで http://localhost:50021 に接続する。
    """

    def __init__(self, base_url: str = "http://localhost:50021", max_workers: int = 4):
        self.base_url = base_url.rstrip("/")
        self.max_workers = max_workers  # audio_query の並列数

    def _audio_query(self, text: str, p: VoiceParams) -> dict:
        """テキストから音声クエリを取得する (声は p.voice = スタイル ID)。"""
        resp = requests.post(
            f"{self.base_url}/audio_query",
            params={"text": text, "speaker": p.voice},
        )
        resp.raise_for_status()
        query = resp.json()
        query["speedScale"] = p.speed
        query["pitchScale"] = p.pitch
        query["intonationScale"] = p.intonation
        query["volumeScale"] = p.volume
        return query

    def _apply_accent_overrides(self, query: dict, accents: list[tuple[str, int]], speaker_id: int) -> dict:
        """accent_phrases のアクセント位置を上書きし、ピッチを再計算する。"""
        if not accents:
            return query
        phrases = query.get("accent_phrases", [])
        # 長音の表記をそろえて比べる (VOICEVOX は「こうぞう」を コオゾオ と表記する)
        texts = [_normalize_kana("".join(m["text"] for m in phrase["moras"])) for phrase in phrases]

        def find(katakana: str, start: int) -> tuple[int, int] | None:
            """start 以降で katakana に当たる句の範囲 (先頭の句, 末尾の句) を探す。

            1. 句が「語そのもの」か「語 + 助詞」(ファイルヲ, ハシワ, ファイルカラ など)
            2. 見つからなければ前方一致 (2文字以上の語のみ。複合語に当たることがある)
            3. 見つからなければ、連続する複数の句にまたがって一致する範囲
               (VOICEVOX が「リスト構造」を リスト + コオゾオ に分けた場合など)
            """
            candidates = range(start, len(phrases))
            for i in candidates:
                if texts[i].startswith(katakana) and (
                        len(texts[i]) == len(katakana) or texts[i][len(katakana):] in _PARTICLES):
                    return i, i
            if len(katakana) >= 2:
                for i in candidates:
                    if texts[i].startswith(katakana):
                        return i, i
            for i in candidates:
                joined = texts[i]
                for j in range(i + 1, len(phrases)):
                    if not katakana.startswith(joined):
                        break
                    joined += texts[j]
                    if joined.startswith(katakana):
                        return i, j
            return None

        # 指定は文中の出現順に並んでいるので、前の指定が当たった句より後ろから探す
        # (同じ語が2回出る文や、同じ読みの語が並ぶ文で、先頭の句にばかり当たらないように)。
        # 順序どおりに見つからないとき (1つの句に2つの指定がある場合など) だけ文頭から探し直す
        modified = False
        pos = 0
        for katakana, accent_pos in accents:
            katakana = _normalize_kana(katakana)
            found = find(katakana, pos) or find(katakana, 0)
            if found is None:
                continue
            i, j = found
            if j > i:
                # 複数の句を1つに結合する (VOICEVOX エディタの「アクセント句の結合」と同じ)
                merged = phrases[i]
                merged["moras"] = [m for p in phrases[i:j + 1] for m in p["moras"]]
                merged["pause_mora"] = phrases[j].get("pause_mora")
                merged["is_interrogative"] = phrases[j].get("is_interrogative", False)
                del phrases[i + 1:j + 1]
                texts[i:j + 1] = ["".join(texts[i:j + 1])]
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
                        params={"speaker": speaker_id},
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

    def _multi_synthesis(self, queries: list[dict], speaker_id: int) -> list[bytes]:
        """複数の音声クエリを一括合成し、WAVリストを返す。"""
        resp = requests.post(
            f"{self.base_url}/multi_synthesis",
            params={"speaker": speaker_id},
            json=queries,
        )
        resp.raise_for_status()
        wav_list = []
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            for name in sorted(zf.namelist()):
                wav_list.append(zf.read(name))
        return wav_list

    def synthesize_sentences(self, sentences: list[Sentence], params: list[VoiceParams],
                             on_done=None) -> list[bytes]:
        """audio_query を並列実行し、アクセントを上書きして multi_synthesis で一括合成する。"""
        # --- audio_query を並列実行 ---
        queries = [None] * len(sentences)
        with ThreadPoolExecutor(max_workers=self.max_workers) as pool:
            futures = {
                pool.submit(self._audio_query, s.reading, params[i]): i
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
                queries[i] = self._apply_accent_overrides(queries[i], s.accents, params[i].voice)

        # --- multi_synthesis で一括合成 (話者 ID ごとにまとめて呼ぶ) ---
        by_speaker: dict[int, list[int]] = {}
        for i, p in enumerate(params):
            by_speaker.setdefault(p.voice, []).append(i)
        wavs: list[bytes] = [b""] * len(sentences)
        for speaker_id, indices in by_speaker.items():
            for i, wav in zip(indices, self._multi_synthesis([queries[i] for i in indices], speaker_id)):
                wavs[i] = wav
        return wavs

    def cache_id(self) -> str:
        """URL とエンジンのバージョン (バージョンが変わると声が変わることがあるため)。"""
        if not hasattr(self, "_version"):
            try:
                resp = requests.get(f"{self.base_url}/version", timeout=5)
                resp.raise_for_status()
                self._version = resp.text.strip()
            except requests.RequestException:
                self._version = ""
        return f"voicevox {self.base_url} {self._version}"

    def engine_name(self) -> str | None:
        """エンジンの名前 (/engine_manifest の name。例: "VOICEVOX Engine")。取得できなければ None。"""
        try:
            resp = requests.get(f"{self.base_url}/engine_manifest", timeout=5)
            resp.raise_for_status()
            return resp.json().get("name")
        except (requests.RequestException, ValueError, AttributeError):
            return None

    def list_speakers(self, timeout: float | tuple[float, float] | None = None) -> list[dict]:
        """利用可能な話者一覧を取得する。

        Args:
            timeout: requests のタイムアウト秒 (接続, 読み込み)。None なら無制限
        """
        resp = requests.get(f"{self.base_url}/speakers", timeout=timeout)
        resp.raise_for_status()
        return resp.json()
