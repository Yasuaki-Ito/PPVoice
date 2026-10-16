"""PowerPoint との連携 (COM): 動画の書き出し・スライドの画像 (字幕のプレビュー・サムネイル)・スライドショー

PowerPoint がなくても PPVoice は動く (連携の機能だけが使えない)。安全に動かすための決まり:
- 利用者のファイルには触らない: 開くのは一時フォルダにコピーしたファイル (スライドショーは PPVoice の出力だけ)
- PowerPoint 本体は、連携を始めたときに起動していなかった場合だけ終了する。自分で開いた資料は必ず閉じる
- ウィンドウを出さない (WithWindow=False)。スライドショーは利用者に見せるのが目的なので出す
- PowerPoint を操作するのは専用のスレッド1つ (COM はスレッドごとに初期化が必要)。要求は順番に処理し、
  応答がなければタイムアウトで諦める
"""

import os
import queue
import shutil
import tempfile
import threading
import time
from concurrent.futures import Future
from typing import Callable

try:
    import pythoncom
    import win32com.client
    _HAS_PYWIN32 = True
except ImportError:
    _HAS_PYWIN32 = False

# CreateVideoStatus (PpMediaTaskStatus)
_STATUS_IN_PROGRESS, _STATUS_QUEUED, _STATUS_DONE, _STATUS_FAILED = 1, 2, 3, 4

OPEN_TIMEOUT = 120      # 開く・画像を書き出すときの待ち時間 (秒)


class PowerPointError(Exception):
    """PowerPoint が使えない・応答しない・処理に失敗した"""


def is_available() -> bool:
    """PowerPoint と連携できそうか (pywin32 があり、PowerPoint が登録されている)。PowerPoint は起動しない。"""
    if not _HAS_PYWIN32:
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"PowerPoint.Application\CLSID"):
            return True
    except OSError:
        return False


class _Worker:
    """PowerPoint を操作する専用のスレッド"""

    def __init__(self):
        self._jobs: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()

    def call(self, func: Callable, timeout: float | None):
        """func(app) を専用のスレッドで実行して結果を返す。timeout 秒で返らなければ PowerPointError。"""
        if not is_available():
            raise PowerPointError("PowerPoint is not available")
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(target=self._run, daemon=True)
                self._thread.start()
        future: Future = Future()
        self._jobs.put((func, future))
        try:
            return future.result(timeout)
        except TimeoutError:
            raise PowerPointError("PowerPoint is not responding")

    def _run(self):
        pythoncom.CoInitialize()
        app = None
        started_by_us = False
        try:
            while True:
                try:
                    func, future = self._jobs.get(timeout=30)
                except queue.Empty:
                    break  # しばらく要求がなければ後始末して終わる
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    if app is None:
                        started_by_us = not _is_running()
                        app = win32com.client.Dispatch("PowerPoint.Application")
                    future.set_result(func(app))
                except Exception as e:
                    future.set_exception(e if isinstance(e, PowerPointError) else PowerPointError(str(e)))
        finally:
            # 起動していなかった PowerPoint を起動したのなら、(利用者が資料を開いていなければ) 終了する
            if app is not None and started_by_us:
                try:
                    if app.Presentations.Count == 0:
                        app.Quit()
                except Exception:
                    pass
            app = None
            pythoncom.CoUninitialize()


def _is_running() -> bool:
    try:
        win32com.client.GetActiveObject("PowerPoint.Application")
        return True
    except Exception:
        return False


_worker = _Worker()


class _TempCopy:
    """利用者のファイルを一時フォルダにコピーし、終わったら消す"""

    def __init__(self, path: str):
        self.path = path

    def __enter__(self) -> str:
        self.dir = tempfile.mkdtemp(prefix="ppvoice_ppt_")
        # コピーは中身だけ (shutil.copyfile は「インターネットから取得」の印 (Zone.Identifier) を写さない
        # ため、保護ビューのダイアログで止まらない)
        dst = os.path.join(self.dir, "copy" + os.path.splitext(self.path)[1])
        shutil.copyfile(self.path, dst)
        return dst

    def __exit__(self, *exc):
        shutil.rmtree(self.dir, ignore_errors=True)


def _open(app, path: str, window: bool = False):
    return app.Presentations.Open(os.path.abspath(path), ReadOnly=True, Untitled=False, WithWindow=window)


# ---------------------------------------------------------------------------
# 機能
# ---------------------------------------------------------------------------

def export_slides(pptx_path: str, slide_numbers: list[int], out_dir: str, width: int = 1280,
                  height: int | None = None, timeout: float = OPEN_TIMEOUT) -> dict[int, str]:
    """スライドを PNG に書き出す。slide_numbers は 1 始まり。Returns {スライド番号: PNG のパス}"""
    os.makedirs(out_dir, exist_ok=True)

    def job(app):
        with _TempCopy(pptx_path) as copy:
            pres = _open(app, copy)
            try:
                w, h = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
                result = {}
                for n in slide_numbers:
                    if 1 <= n <= pres.Slides.Count:
                        path = os.path.normpath(os.path.join(os.path.abspath(out_dir), f"slide{n}.png"))
                        pres.Slides(n).Export(path, "PNG", width, height or round(width * h / w))
                        result[n] = path
                return result
            finally:
                pres.Close()

    return _worker.call(job, timeout)


def create_video(pptx_path: str, mp4_path: str, height: int = 1080, fps: int = 30, quality: int = 85,
                 default_duration: int = 5, on_progress: Callable[[float], None] | None = None,
                 should_cancel: Callable[[], bool] | None = None) -> str:
    """PPTX を MP4 に書き出す (ナレーションとタイミングを使う)。

    on_progress には経過秒数を渡す (PowerPoint は進み具合の割合を返さない)。
    should_cancel が True を返すと書き出しをやめる (PowerPointError)。
    """
    mp4_path = os.path.normpath(os.path.abspath(mp4_path))

    def job(app):
        with _TempCopy(pptx_path) as copy:
            pres = _open(app, copy)
            try:
                tmp_mp4 = os.path.join(os.path.dirname(copy), "video.mp4")
                pres.CreateVideo(tmp_mp4, True, default_duration, height, fps, quality)
                t0 = time.time()
                while True:
                    status = pres.CreateVideoStatus
                    if status not in (_STATUS_IN_PROGRESS, _STATUS_QUEUED):
                        break
                    if should_cancel and should_cancel():
                        raise PowerPointError("cancelled")
                    if on_progress:
                        on_progress(time.time() - t0)
                    time.sleep(0.5)
                if status != _STATUS_DONE or not os.path.exists(tmp_mp4):
                    raise PowerPointError(f"video export failed (status {status})")
            finally:
                pres.Close()
            # 書き出しが終わってから出力先に移す (途中のファイルを残さない)
            shutil.move(tmp_mp4, mp4_path)
        return mp4_path

    return _worker.call(job, None)


def start_slideshow(pptx_path: str, slide: int = 1, timeout: float = OPEN_TIMEOUT):
    """PPTX (PPVoice の出力) を PowerPoint で開き、slide からスライドショーを始める。資料は開いたままにする。"""
    def job(app):
        pres = _open(app, pptx_path, window=True)
        settings = pres.SlideShowSettings
        settings.StartingSlide = max(1, min(slide, pres.Slides.Count))
        settings.EndingSlide = pres.Slides.Count
        settings.Run()

    _worker.call(job, timeout)


def close_file(path: str, timeout: float = 10):
    """path を PowerPoint で開いていれば、保存せずに閉じる (上書きする前に)。PowerPoint が起動していなければ何もしない。"""
    if not _HAS_PYWIN32 or not _is_running_quick():
        return
    target = os.path.normcase(os.path.abspath(path))

    def job(app):
        for i in range(app.Presentations.Count, 0, -1):
            pres = app.Presentations(i)
            if os.path.normcase(pres.FullName) == target:
                pres.Saved = True
                pres.Close()

    try:
        _worker.call(job, timeout)
    except PowerPointError:
        pass


def _is_running_quick() -> bool:
    """PowerPoint が起動しているか (このスレッドで COM を初期化して調べる)"""
    pythoncom.CoInitialize()
    try:
        return _is_running()
    finally:
        pythoncom.CoUninitialize()
