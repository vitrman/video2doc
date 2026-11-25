"""Video2Doc — видео в Word: речь (GigaAM, локально) + кадры через N секунд с отметками времени.

Распознаёт локально. Модель GigaAM v3 (RNNT «точнее» или CTC «быстрее») ищется в папке models рядом
с exe и в кэше, а если её нет — один раз скачивается (CDN Сбера, запасной источник — релиз
models-v1 на GitHub). Итогов и выводов не делает: время, кадр, текст.
Движок (asr_engine, model_fetch, postprocess, paths) и оформление взяты из Voice2Text.

Запуск без сборки:  python video2doc_app.py [файлы/папки ...] [--auto]
--auto — сразу начать обработку переданных путей и закрыть окно по завершении.
"""
from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import wave
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox
from xml.sax.saxutils import escape

FROZEN = bool(getattr(sys, "frozen", False))
APP_DIR = Path(sys.executable).resolve().parent if FROZEN else Path(__file__).resolve().parent

VIDEO_EXT = {".mp4", ".mkv", ".avi", ".mov", ".webm", ".wmv", ".m4v", ".flv", ".mpg", ".mpeg", ".ts"}
SR = 16_000
NO_WINDOW = 0x08000000 if os.name == "nt" else 0
LOG_FILE = Path(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()) / "Video2Doc" / "video2doc.log"


def ffmpeg_exe() -> str:
    local = APP_DIR / "bin" / "ffmpeg.exe"
    return str(local) if local.exists() else "ffmpeg"


def ts(sec: int) -> str:
    return f"{sec // 3600:02d}:{sec % 3600 // 60:02d}:{sec % 60:02d}"


# ---------- чистка текста ----------
# короткие заполнители, которые не ловит voice2text («э», «э-э», «ээ», «эм», «мм», «хм»)
_SHORT_FILLERS = re.compile(r"(?<!\w)(?:э(?:-э)*|э{2,}|эм+|м{2,}|хм+|([аоуыи])(?:-\1)+)(?!\w)", re.IGNORECASE)


def clean(text: str) -> str:
    from postprocess import drop_fillers
    text = drop_fillers(text)
    text = _SHORT_FILLERS.sub("", text)
    text = re.sub(r"\s+([,.;:!?])", r"\1", text)
    text = re.sub(r"(?:\s*,\s*){2,}", ", ", text)
    text = re.sub(r"([.!?])\s*,\s*", r"\1 ", text)
    text = re.sub(r"^[\s,;:—–-]+", "", text)
    text = re.sub(r"\s{2,}", " ", text).strip()
    text = re.sub(r"([.!?]\s+)([а-яё])", lambda m: m.group(1) + m.group(2).upper(), text)
    return text[:1].upper() + text[1:] if text else text


# ---------- Word ----------
def _para(text: str, bold: bool = False) -> str:
    b = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return f'<w:p><w:r>{b}<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>'


def _image(rid: str, n: int, w: int, h: int) -> str:
    cx = 5_400_000  # ~15 см
    cy = cx * h // max(w, 1)
    return (
        '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{n}" name="k{n}"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{n}" name="k{n}.jpg"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )


def _jpg_size(data: bytes) -> tuple[int, int]:
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        m = data[i + 1]
        if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            return int.from_bytes(data[i + 7:i + 9], "big"), int.from_bytes(data[i + 5:i + 7], "big")
        i += 2 + int.from_bytes(data[i + 2:i + 4], "big")
    return 16, 9


def write_docx(path: Path, title: str, items: list[tuple[int, Path | None, str]], meta: str = "") -> None:
    body, rels, media = [_para(title, True)], [], []
    if meta:
        body.append(_para(meta))
    for n, (sec, frame, text) in enumerate(items, 1):
        body.append(_para(f"[{ts(sec)}]", True))
        if frame and frame.exists():
            data = frame.read_bytes()
            w, h = _jpg_size(data)
            rid = f"rId{n + 10}"
            body.append(_image(rid, n, w, h))
            rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/'
                        f'2006/relationships/image" Target="media/k{n}.jpg"/>')
            media.append((f"word/media/k{n}.jpg", data))
        if text:
            body.append(_para(text))
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
           + "".join(body) + "</w:body></w:document>")
    tmp = path.with_name(path.name + ".part")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
                   'wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
                   'officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + "".join(rels) + "</Relationships>")
        z.writestr("word/document.xml", doc)
        for name, data in media:
            z.writestr(name, data)
    tmp.replace(path)


# ---------- обработка ----------
class Cancelled(Exception):
    pass


def collect_videos(sources: list[str], recursive: bool) -> list[Path]:
    found: list[Path] = []
    for src in sources:
        p = Path(src)
        if p.is_dir():
            files = p.rglob("*") if recursive else p.glob("*")
            found += sorted(f for f in files if f.is_file() and f.suffix.lower() in VIDEO_EXT)
        elif p.is_file() and p.suffix.lower() in VIDEO_EXT:
            found.append(p)
    unique, seen = [], set()
    for f in found:
        key = str(f.resolve()).lower()
        if key not in seen:
            seen.add(key)
            unique.append(f)
    return unique


def run_ffmpeg(args: list, stop: threading.Event) -> None:
    proc = subprocess.Popen([ffmpeg_exe(), "-y", "-v", "error", *map(str, args)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, creationflags=NO_WINDOW)
    while proc.poll() is None:
        if stop.is_set():
            proc.kill()
            raise Cancelled
        time.sleep(0.1)
    if proc.returncode:
        err = proc.stderr.read().decode("utf-8", "replace").strip().splitlines()
        raise RuntimeError("ffmpeg: " + (err[-1] if err else f"код {proc.returncode}"))


VERSION = "1.4"

# модели GigaAM: имя -> (подпись в списке, подсказка, имя в docx)
MODELS = {
    "v3_e2e_rnnt": ("GigaAM v3 — точнее",
                    "RNNT: лучше расставляет слова и знаки, чуть медленнее. Подходит по умолчанию.",
                    "GigaAM v3 RNNT"),
    "v3_e2e_ctc": ("GigaAM v3 CTC — быстрее",
                   "CTC: быстрее на слабом процессоре, иногда больше ошибок в словах.",
                   "GigaAM v3 CTC"),
}
DEFAULT_MODEL = "v3_e2e_rnnt"
CONFIG_FILE = LOG_FILE.parent / "config.json"
ERROR_LOG = LOG_FILE.parent / "errors.log"
# out_dir: "" — сохранять рядом с видео, иначе — в эту папку
DEFAULTS = {"model": DEFAULT_MODEL, "step": 20, "recursive": False, "overwrite": False, "fillers": True,
            "frames": True, "quality": "full",
            "out_dir": ""}
# качество кадров: подпись, подсказка, аргументы ffmpeg, примерный размер кадра в МБ
QUALITY = {
    "full": ("Как в видео", "Исходное разрешение — мелкий текст на записи экрана читается. Файл тяжелее.",
             ["-q:v", 2], 0.3),
    "960": ("Сжатые, 960 px", "Уменьшены до 960 px по ширине — документ легче, мелкий текст может размыться.",
            ["-vf", "scale='min(960,iw)':-2", "-q:v", 4], 0.08),
}


def load_config() -> dict:
    import json
    cfg = dict(DEFAULTS)
    try:
        cfg.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    if cfg.get("quality") not in QUALITY:
        cfg["quality"] = DEFAULTS["quality"]
    if cfg.get("model") not in MODELS:
        cfg["model"] = DEFAULT_MODEL
    return cfg


def save_config(cfg: dict) -> None:
    import json
    try:
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


# ---------- диагностика: проверки до старта, понятные ошибки, отчёт ----------
class PreflightError(Exception):
    """Проблема, найденная до обработки: текст уже понятен человеку."""


def out_path(video: Path, opts: dict) -> Path:
    folder = Path(opts["out_dir"]) if opts.get("out_dir") else video.parent
    return folder / (video.stem + ".docx")


def check_writable(folder: Path) -> str:
    """Пусто — можно писать; иначе — понятная причина."""
    if not folder.exists():
        return f"Папка не найдена: {folder}"
    probe = folder / f".video2doc_probe_{os.getpid()}.tmp"
    try:
        probe.write_bytes(b"ok")
        probe.unlink()
        return ""
    except PermissionError:
        return (f"Нет доступа на запись в папку: {folder}\n"
                "    Возможные причины:\n"
                "    • антивирус блокирует запись неизвестной программе (Kaspersky — «Контроль программ» / "
                "«Защита от шифровальщиков», Защитник Windows — «Контролируемый доступ к папкам»): "
                "добавьте Video2Doc.exe в доверенные программы;\n"
                "    • у пользователя нет прав на эту папку или она только для чтения (сетевой диск, флешка, CD);\n"
                "    • выход: в «Настройках» выберите другую папку для результатов.")
    except OSError as exc:
        return f"Не удаётся записать в папку {folder}: {exc}"


def temp_dir() -> Path:
    """Временная папка; tempfile.gettempdir() падает, если писать некуда, — тогда берём TEMP как есть."""
    try:
        return Path(tempfile.gettempdir())
    except OSError:
        return Path(os.environ.get("TEMP") or os.environ.get("TMP") or ".")


def free_mb(path: Path) -> float:
    import shutil
    try:
        return shutil.disk_usage(path).free / 1e6
    except OSError:
        return float("inf")


def video_duration(video: Path) -> int:
    """Длительность в секундах по выводу ffmpeg -i (0, если не удалось)."""
    try:
        proc = subprocess.run([ffmpeg_exe(), "-hide_banner", "-i", str(video)], capture_output=True,
                              creationflags=NO_WINDOW, timeout=60)
        m = re.search(rb"Duration: (\d+):(\d+):(\d+)", proc.stderr)
        return int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3)) if m else 0
    except (OSError, subprocess.SubprocessError):
        return 0


def model_location(model: str) -> Path | None:
    """Где лежит модель: рядом с exe, в своём кэше или в кэше Voice2Text. None — надо скачать."""
    from asr_engine import MODEL_DIR
    from model_fetch import search_dirs
    for folder in search_dirs(MODEL_DIR):
        if (folder / f"{model}.ckpt").is_file():
            return folder
    return None


def preflight(videos: list[Path], opts: dict, log) -> list[Path]:
    """Проверяет окружение до долгой обработки. Возвращает видео, которые можно обрабатывать."""
    try:
        subprocess.run([ffmpeg_exe(), "-version"], capture_output=True, check=True, creationflags=NO_WINDOW,
                       timeout=30)
    except (OSError, subprocess.SubprocessError):
        raise PreflightError(f"Не найден или не запускается ffmpeg: {ffmpeg_exe()}\n"
                             "    Скопируйте программу целиком, вместе с папкой bin.")
    model = opts["model"]
    if model_location(model) is None:
        from asr_engine import MODEL_DIR
        log(f"Модели {MODELS[model][0]} нет на компьютере — скачаю один раз (~450 МБ) в {MODEL_DIR}")
        if problem := check_writable(MODEL_DIR):
            raise PreflightError("Некуда скачать модель. " + problem)
        if free_mb(MODEL_DIR) < 1000:
            raise PreflightError(f"Мало места для модели в {MODEL_DIR}: свободно {free_mb(MODEL_DIR):.0f} МБ, "
                                 "нужно около 1000 МБ.")

    temp = temp_dir()
    if problem := check_writable(temp):
        raise PreflightError("Временная папка недоступна. " + problem)

    folders: dict[Path, list[Path]] = {}
    for v in videos:
        folders.setdefault(out_path(v, opts).parent, []).append(v)
    ok: list[Path] = []
    longest = 0
    for folder, items in folders.items():
        if problem := check_writable(folder):
            log(f"✖ {problem}\n    Пропущено видео в этой папке: {len(items)}")
            continue
        need = 0.0
        per_shot = QUALITY[opts["quality"]][3] if opts["frames"] else 0.0   # размер кадра в docx
        for v in items:
            d = video_duration(v)
            longest = max(longest, d)
            need += d / max(opts["step"], 1) * per_shot + 1
        if free_mb(folder) < need + 50:
            log(f"✖ Мало места на диске для {folder}: свободно {free_mb(folder):.0f} МБ, "
                f"нужно примерно {need + 50:.0f} МБ. Пропущено видео: {len(items)}")
            continue
        ok += items
    temp_need = longest * 0.032 + longest / max(opts["step"], 1) * (QUALITY[opts["quality"]][3] if opts["frames"] else 0) + 50  # wav + кадры
    if ok and free_mb(temp) < temp_need:
        raise PreflightError(f"Мало места во временной папке {temp}: свободно {free_mb(temp):.0f} МБ, "
                             f"нужно примерно {temp_need:.0f} МБ. Освободите место на диске C:.")
    return ok


def explain(exc: BaseException, out: Path) -> str:
    """Короткое объяснение частых ошибок обработки."""
    import errno
    if isinstance(exc, PermissionError):
        if out.exists():
            return (f"Не удалось перезаписать {out.name}: файл, вероятно, открыт в Word. "
                    "Закройте его и запустите снова.")
        return check_writable(out.parent) or f"Нет доступа: {exc}"
    if isinstance(exc, OSError) and exc.errno == errno.ENOSPC:
        return f"Закончилось место на диске ({out.drive or 'диск'})."
    if isinstance(exc, MemoryError):
        return "Не хватило оперативной памяти. Закройте другие программы или выберите модель «быстрее»."
    text = str(exc)
    if text.startswith("ffmpeg:"):
        return f"Видео не читается ({text}). Файл повреждён или формат не поддерживается."
    return text or exc.__class__.__name__


def write_error(video: Path | None, exc: BaseException, opts: dict) -> None:
    """Подробная запись в errors.log: этот файл — для разработчика."""
    import traceback
    try:
        ERROR_LOG.parent.mkdir(parents=True, exist_ok=True)
        with ERROR_LOG.open("a", encoding="utf-8") as f:
            f.write(f"===== {time.strftime('%Y-%m-%d %H:%M:%S')} · Video2Doc {VERSION}\n"
                    f"видео: {video}\nнастройки: {opts}\n"
                    + "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)) + "\n")
    except OSError:
        pass


def environment_report(cfg: dict) -> str:
    import platform
    lines = [f"Video2Doc {VERSION}", f"Дата: {time.strftime('%Y-%m-%d %H:%M:%S')}",
             f"Windows: {platform.platform()}", f"Процессор: {platform.processor()} · ядер: {os.cpu_count()}"]
    try:
        import ctypes

        class MEM(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                        ("a", ctypes.c_ulonglong), ("b", ctypes.c_ulonglong), ("c", ctypes.c_ulonglong),
                        ("d", ctypes.c_ulonglong), ("e", ctypes.c_ulonglong)]
        m = MEM()
        m.dwLength = ctypes.sizeof(MEM)
        ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
        lines.append(f"Память: всего {m.ullTotalPhys / 1e9:.1f} ГБ, свободно {m.ullAvailPhys / 1e9:.1f} ГБ")
    except Exception:
        pass
    lines += [f"Папка программы: {APP_DIR}",
              f"ffmpeg: {ffmpeg_exe()} ({'есть' if Path(ffmpeg_exe()).exists() else 'из PATH'})",
              "Модели: " + ", ".join(f"{m}: {model_location(m) or 'не скачана'}" for m in MODELS),
              f"Временная папка: {temp_dir()} · свободно {free_mb(temp_dir()):.0f} МБ",
              f"Настройки: {cfg}"]
    return "\n".join(lines)


class Worker:
    def __init__(self, log, progress):
        self.log, self.progress = log, progress
        self.engine = None
        self.engine_model = ""
        self.stop = threading.Event()
        self.current = (1, 1)   # номер файла и всего — для полосы прогресса при скачивании модели

    # сообщения движка, которые нужны человеку: скачивание модели и его сбои
    _ENGINE_LOG = ("скачиваю", "источник недоступен", "скачивание сорвалось", "контрольная сумма",
                   "модель не скачалась", "токенизатор")

    def _engine_log(self, text: str) -> None:
        if text.startswith(self._ENGINE_LOG):
            self.log("    " + text)

    def _download_progress(self, percent: int) -> None:
        n, total = self.current
        self.progress(n, total, percent / 100, "скачивание модели")

    def get_engine(self, model: str):
        if self.engine is not None and self.engine_model != model:
            self.engine.dispose()
            self.engine = None
        if self.engine is None:
            self.log(f"Загружаю модель {MODELS[model][0]}…")
            started = time.perf_counter()
            from asr_engine import GigaAMEngine
            engine = GigaAMEngine(model_name=model, log=self._engine_log, on_progress=self._download_progress)
            engine.load()
            self.engine, self.engine_model = engine, model
            self.log(f"Модель готова: {MODELS[model][0]} ({engine.name}), {time.perf_counter() - started:.0f} с")
        return self.engine

    def process(self, video: Path, opts: dict, n: int, total_files: int) -> Path:
        import numpy as np
        from asr_engine import split_audio
        step, model = opts["step"], opts["model"]
        self.current = (n, total_files)
        out = out_path(video, opts)
        with tempfile.TemporaryDirectory(prefix="video2doc_") as tmp:
            tmp = Path(tmp)
            wav = tmp / "a.wav"
            self.progress(n, total_files, 0, "звук")
            run_ffmpeg(["-i", video, "-vn", "-ac", 1, "-ar", SR, "-sample_fmt", "s16", wav], self.stop)
            with wave.open(str(wav), "rb") as wf:
                audio = np.frombuffer(wf.readframes(wf.getnframes()), dtype=np.int16)
            duration = len(audio) // SR
            engine = self.get_engine(model)
            chunks = split_audio(audio, SR, 20.0)
            speech, offset = [], 0
            for i, chunk in enumerate(chunks, 1):
                if self.stop.is_set():
                    raise Cancelled
                text = engine._transcribe_chunk(chunk, SR)
                if text:
                    speech.append((offset // SR, text))
                offset += len(chunk)
                share = 0.85 if opts["frames"] else 1.0
                self.progress(n, total_files, i / max(len(chunks), 1) * share, "речь")
            items = []
            points = list(range(0, max(duration, 1), step))
            for i, t in enumerate(points, 1):
                frame = None
                if opts["frames"]:      # без кадров ffmpeg по видео не ходит вовсе — только текст
                    frame = tmp / f"f{t}.jpg"
                    try:
                        run_ffmpeg(["-ss", t, "-i", video, "-frames:v", 1, *QUALITY[opts["quality"]][2], frame],
                                   self.stop)
                    except RuntimeError:
                        frame = None  # видео без картинки или битый кадр — просто без скрина
                    self.progress(n, total_files, 0.85 + i / len(points) * 0.15, "кадры")
                text = " ".join(s for sec, s in speech if t <= sec < t + step)
                items.append((t, frame, clean(text) if opts["fillers"] else text))
            shots = (f"кадр каждые {step} с, {QUALITY[opts['quality']][0].lower()}") if opts["frames"] else f"без кадров, отметки каждые {step} с"
            meta = f"Расшифровка: {MODELS[model][2]}, {shots}, {time.strftime('%d.%m.%Y')}"
            write_docx(out, video.stem, items, meta)
        return out


# ---------- окно (оформление как в voice2text) ----------
BG, CARD, CARD_HOVER, BORDER = "#121316", "#1b1d21", "#23262b", "#2e3238"
ACCENT, ACCENT_HOVER = "#4f8cff", "#3d78e6"
TEXT, MUTED, ERROR, OK = "#f2f3f5", "#8b9199", "#e0554f", "#4caf68"


class App:
    def __init__(self, paths: list[str], auto: bool):
        import customtkinter as ctk
        self.ctk = ctk
        ctk.set_appearance_mode("dark")
        self.root = ctk.CTk()
        self.root.title(f"Video2Doc {VERSION}")
        self.root.geometry("1100x680")
        self.root.minsize(900, 520)
        self.root.configure(fg_color=BG)
        ico = APP_DIR / "video2doc.ico"
        if ico.exists():
            try:
                self.root.iconbitmap(str(ico))
            except tk.TclError:
                pass

        self.cfg = load_config()
        self.msgs: queue.Queue = queue.Queue()
        self.worker = Worker(self.log_message, lambda *a: self.msgs.put(("progress", a)))
        self.thread: threading.Thread | None = None
        self.auto = auto
        self.sources: list[str] = []
        self.font = lambda size=13, weight="normal": ctk.CTkFont("Segoe UI", size, weight)

        main = ctk.CTkFrame(self.root, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=(16, 18))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(0, weight=1)
        left = ctk.CTkFrame(main, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        self._build_left(left)
        panel = ctk.CTkScrollableFrame(main, width=340, corner_radius=16, fg_color=CARD)
        panel.grid(row=0, column=1, sticky="ns")
        self._build_settings(panel)

        for p in paths:
            self.add_source(p)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.poll)
        if auto and paths:
            self.root.after(300, self.start)

    # ----- разметка -----
    def _button(self, parent, text, command, primary=False, width=150, height=34):
        ctk = self.ctk
        if primary:
            return ctk.CTkButton(parent, text=text, command=command, width=width, height=height,
                                 corner_radius=10, fg_color=ACCENT, hover_color=ACCENT_HOVER,
                                 font=self.font(13, "bold" if height > 40 else "normal"))
        return ctk.CTkButton(parent, text=text, command=command, width=width, height=height, corner_radius=10,
                             fg_color="transparent", border_width=1, border_color=BORDER, text_color=TEXT,
                             hover_color=CARD_HOVER, font=self.font(13))

    def _build_left(self, left):
        ctk = self.ctk
        header = ctk.CTkFrame(left, fg_color="transparent")
        header.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(header, text="Видео в Word", font=self.font(20, "bold")).pack(side="left")
        ctk.CTkLabel(header, text=f"GigaAM · {VERSION}", text_color=MUTED,
                     font=self.font(12)).pack(side="left", padx=(10, 0), pady=(6, 0))

        card = ctk.CTkFrame(left, corner_radius=16, fg_color=CARD)
        card.pack(fill="both", expand=True, pady=(0, 10))
        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=(14, 6))
        ctk.CTkLabel(top, text="Источники: файлы и папки", font=self.font(13)).pack(side="left")
        self._button(top, "Очистить", self.clear_sources, width=100).pack(side="right")
        self._button(top, "Добавить папку…", self.add_folder).pack(side="right", padx=(0, 8))
        self._button(top, "Добавить файлы…", self.add_files, primary=True).pack(side="right", padx=(0, 8))
        self.src_frame = ctk.CTkScrollableFrame(card, fg_color=BG, corner_radius=12, height=150)
        self.src_frame.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self.src_empty = ctk.CTkLabel(self.src_frame, text="Добавьте видео или папку — или перетащите их на Video2Doc.exe",
                                      text_color=MUTED, font=self.font(12))
        self.src_empty.pack(pady=20)

        run = ctk.CTkFrame(left, corner_radius=16, fg_color=CARD)
        run.pack(fill="x", pady=(0, 10))
        self.start_btn = self._button(run, "▶  Старт", self.start, primary=True, width=150, height=46)
        self.start_btn.pack(side="left", padx=(16, 8), pady=14)
        self.stop_btn = self._button(run, "Стоп", self.cancel, width=90, height=46)
        self.stop_btn.configure(state="disabled")
        self.stop_btn.pack(side="left", pady=14)
        self.bar = ctk.CTkProgressBar(run, height=10, corner_radius=5, progress_color=ACCENT, fg_color=CARD_HOVER)
        self.bar.set(0)
        self.bar.pack(side="left", fill="x", expand=True, padx=16)

        self.status = tk.StringVar(value=f"Готово к работе · модель: {MODELS[self.cfg['model']][0]}")
        ctk.CTkLabel(left, textvariable=self.status, text_color=MUTED, anchor="w",
                     font=self.font(12)).pack(fill="x", pady=(0, 8))
        self.logbox = ctk.CTkTextbox(left, wrap="word", corner_radius=14, fg_color=CARD, border_width=1,
                                     border_color=BORDER, text_color=TEXT, font=self.font(12), height=160)
        self.logbox.pack(fill="both", expand=True)
        # журнал только для чтения, но выделяется и копируется (в disabled-режиме копирование не работало)
        self.logbox.bind("<Key>", self._log_key)
        self.logbox.bind("<Button-3>", self._log_menu)
        self.logbox.tag_config("err", foreground=ERROR)
        logbar = ctk.CTkFrame(left, fg_color="transparent")
        logbar.pack(fill="x", pady=(8, 0))
        self._button(logbar, "Копировать журнал", self.copy_log, width=170, height=30).pack(side="left")
        self._button(logbar, "Отчёт для разработчика…", self.save_report, width=210, height=30).pack(
            side="left", padx=(8, 0))
        self._button(logbar, "Открыть папку журналов", self.open_log, width=200, height=30).pack(
            side="left", padx=(8, 0))

    def _log_key(self, event):
        if event.state & 0x4:                       # Ctrl; keycode не зависит от раскладки
            if event.keycode == 65:                   # Ctrl+A / Ctrl+Ф
                self.logbox.tag_add("sel", "1.0", "end")
                return "break"
            if event.keycode in (67, 45):             # Ctrl+C / Ctrl+С, Ctrl+Insert
                self.logbox.event_generate("<<Copy>>")
                return "break"
        if event.keysym in ("Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next"):
            return None
        return "break"

    def _log_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Копировать выделенное", command=lambda: self.logbox.event_generate("<<Copy>>"))
        menu.add_command(label="Копировать весь журнал", command=self.copy_log)
        menu.tk_popup(event.x_root, event.y_root)

    def copy_log(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.logbox.get("1.0", "end").strip())
        self.status.set("Журнал скопирован в буфер обмена")

    def save_report(self):
        """Один текстовый файл для разработчика: окружение, ошибки, хвост журнала."""
        def tail(path: Path, n: int) -> str:
            try:
                return "\n".join(path.read_text(encoding="utf-8", errors="replace").splitlines()[-n:])
            except OSError:
                return "(нет)"
        report = (environment_report(self.cfg)
                  + "\n\n===== Журнал в окне (текущий сеанс) =====\n" + self.logbox.get("1.0", "end").strip()
                  + "\n\n===== Ошибки (errors.log) =====\n" + tail(ERROR_LOG, 400)
                  + "\n\n===== Журнал (последние 300 строк) =====\n" + tail(LOG_FILE, 300) + "\n")
        default = f"Video2Doc_отчёт_{time.strftime('%Y%m%d_%H%M')}.txt"
        path = filedialog.asksaveasfilename(title="Сохранить отчёт для разработчика", initialfile=default,
                                            defaultextension=".txt", filetypes=[("Текст", "*.txt")],
                                            initialdir=str(Path.home() / "Desktop"))
        if not path:
            return
        try:
            Path(path).write_text(report, encoding="utf-8")
        except OSError as exc:
            self.root.clipboard_clear()
            self.root.clipboard_append(report)
            messagebox.showwarning("Video2Doc", f"Не удалось сохранить файл ({exc}).\n"
                                                "Отчёт скопирован в буфер обмена — вставьте его в письмо.")
            return
        os.startfile(Path(path).parent)
        self.status.set(f"Отчёт сохранён: {path}")

    def _label(self, panel, text, top=12):
        self.ctk.CTkLabel(panel, text=text, anchor="w", text_color=TEXT,
                          font=self.font(12)).pack(fill="x", padx=16, pady=(top, 2))

    def _hint(self, panel, text):
        w = self.ctk.CTkLabel(panel, text=text, anchor="w", justify="left", wraplength=300, text_color=MUTED,
                              font=self.font(11))
        w.pack(fill="x", padx=16, pady=(0, 6))
        return w

    def _switch(self, panel, text, key, on_change=None):
        var = tk.BooleanVar(value=bool(self.cfg[key]))

        def changed():
            self.set_cfg(key, var.get())
            if on_change:
                on_change()
        self.ctk.CTkSwitch(panel, text=text, variable=var, width=44, font=self.font(12), progress_color=ACCENT,
                           command=changed).pack(fill="x", padx=16, pady=(10, 2))
        return var

    def _refresh_step_texts(self):
        self.quality_box.configure(state="readonly" if self.cfg["frames"] else "disabled")
        if self.cfg["frames"]:
            self.frames_hint.configure(text="Кадр вставляется в начало каждого отрезка.")
            self.step_label.configure(text="Кадр каждые, секунд")
            self.step_hint.configure(text="Можно вписать своё число. Меньше шаг — больше кадров и тяжелее файл.")
        else:
            self.frames_hint.configure(text="Выключено — только текст: кадры не вырезаются, это быстрее, "
                                            "файл маленький.")
            self.step_label.configure(text="Отметка времени каждые, секунд")
            self.step_hint.configure(text="Текст разбивается на отрезки с отметками [чч:мм:сс].")

    def _build_settings(self, panel):
        ctk = self.ctk
        ctk.CTkLabel(panel, text="Настройки", anchor="w", font=self.font(16, "bold")).pack(fill="x", padx=16,
                                                                                            pady=(16, 4))
        self._label(panel, "Модель распознавания")
        labels = [v[0] for v in MODELS.values()]
        self.var_model = tk.StringVar(value=MODELS[self.cfg["model"]][0])
        self.model_box = ctk.CTkComboBox(panel, values=labels, variable=self.var_model, state="readonly",
                                         corner_radius=10, fg_color=BG, border_color=BORDER, button_color=BORDER,
                                         button_hover_color=CARD_HOVER, dropdown_fg_color=BG,
                                         font=self.font(12), command=self.on_model)
        self.model_box.pack(fill="x", padx=16, pady=(0, 4))
        self.model_hint = self._hint(panel, MODELS[self.cfg["model"]][1])

        self.var_frames = self._switch(panel, "Сохранять кадры (скриншоты)", "frames",
                                       on_change=self._refresh_step_texts)
        self.frames_hint = self._hint(panel, "")
        self._label(panel, "Качество кадров", top=6)
        self.var_quality = tk.StringVar(value=QUALITY[self.cfg["quality"]][0])
        self.quality_box = ctk.CTkComboBox(panel, values=[v[0] for v in QUALITY.values()],
                                           variable=self.var_quality, state="readonly", corner_radius=10,
                                           fg_color=BG, border_color=BORDER, button_color=BORDER,
                                           button_hover_color=CARD_HOVER, dropdown_fg_color=BG,
                                           font=self.font(12), command=self.on_quality)
        self.quality_box.pack(fill="x", padx=16, pady=(0, 4))
        self.quality_hint = self._hint(panel, QUALITY[self.cfg["quality"]][1])
        self.step_label = self.ctk.CTkLabel(panel, text="", anchor="w", text_color=TEXT, font=self.font(12))
        self.step_label.pack(fill="x", padx=16, pady=(12, 2))
        self.var_step = tk.StringVar(value=str(self.cfg["step"]))
        self.step_box = ctk.CTkComboBox(panel, values=["10", "15", "20", "30", "60", "120"],
                                        variable=self.var_step, corner_radius=10, fg_color=BG,
                                        border_color=BORDER, button_color=BORDER, button_hover_color=CARD_HOVER,
                                        dropdown_fg_color=BG, font=self.font(12),
                                        command=lambda _v: self.on_step())
        self.step_box.pack(fill="x", padx=16, pady=(0, 4))
        self.step_box._entry.bind("<FocusOut>", lambda _e: self.on_step())
        self.step_hint = self._hint(panel, "")
        self._refresh_step_texts()

        self.var_recursive = self._switch(panel, "Искать в подпапках", "recursive")
        self.var_overwrite = self._switch(panel, "Перезаписывать готовые .docx", "overwrite")
        self._hint(panel, "Выключено — видео, у которых уже есть .docx, пропускаются.")
        self.var_fillers = self._switch(panel, "Убирать «э-э», «ммм», «а-а»", "fillers")
        self._hint(panel, "Слова «ну», «то есть» не трогаются: они бывают частью смысла.")

        self._label(panel, "Куда сохранять .docx", top=16)
        self.out_label = self._hint(panel, "")
        row = ctk.CTkFrame(panel, fg_color="transparent")
        row.pack(fill="x", padx=16, pady=(0, 4))
        self._button(row, "Выбрать папку…", self.choose_out_dir, width=145, height=30).pack(side="left")
        self._button(row, "Рядом с видео", lambda: self.set_out_dir(""), width=145, height=30).pack(
            side="left", padx=(8, 0))
        self._refresh_out_label()
        self._hint(panel, "Имя — как у видео. Внутри: время, кадр, текст речи, без итогов и выводов.")

    # ----- настройки -----
    def on_quality(self, label):
        key = next(k for k, v in QUALITY.items() if v[0] == label)
        self.set_cfg("quality", key)
        self.quality_hint.configure(text=QUALITY[key][1])

    def set_cfg(self, key, value):
        self.cfg[key] = value
        save_config(self.cfg)

    def _refresh_out_label(self):
        self.out_label.configure(text=f"В папку: {self.cfg['out_dir']}" if self.cfg["out_dir"]
                                 else "Рядом с каждым видео")

    def choose_out_dir(self):
        d = filedialog.askdirectory(title="Папка для результатов")
        if d:
            self.set_out_dir(str(Path(d)))

    def set_out_dir(self, path: str):
        if path and (problem := check_writable(Path(path))):
            messagebox.showwarning("Video2Doc", problem)
            return
        self.set_cfg("out_dir", path)
        self._refresh_out_label()

    def on_model(self, label):
        name = next(k for k, v in MODELS.items() if v[0] == label)
        self.set_cfg("model", name)
        self.model_hint.configure(text=MODELS[name][1])
        if not (self.thread and self.thread.is_alive()):
            self.status.set(f"Готово к работе · модель: {label}")

    def on_step(self):
        try:
            step = max(1, min(3600, int(self.var_step.get())))
        except ValueError:
            step = self.cfg["step"]
        self.var_step.set(str(step))
        self.set_cfg("step", step)

    # ----- источники -----
    def add_source(self, path: str):
        if path in self.sources:
            return
        self.sources.append(path)
        self._render_sources()

    def _render_sources(self):
        for w in self.src_frame.winfo_children():
            if w is not self.src_empty:
                w.destroy()
        if not self.sources:
            self.src_empty.pack(pady=20)
            return
        self.src_empty.pack_forget()
        for path in self.sources:
            row = self.ctk.CTkFrame(self.src_frame, fg_color="transparent")
            row.pack(fill="x", pady=1)
            kind = "📁" if Path(path).is_dir() else "🎬"
            self.ctk.CTkLabel(row, text=f"{kind}  {path}", anchor="w", font=self.font(12)).pack(
                side="left", fill="x", expand=True, padx=(6, 0))
            self.ctk.CTkButton(row, text="✕", width=28, height=24, corner_radius=8, fg_color="transparent",
                               hover_color=CARD_HOVER, text_color=MUTED,
                               command=lambda p=path: self.remove_source(p)).pack(side="right", padx=4)

    def remove_source(self, path):
        self.sources.remove(path)
        self._render_sources()

    def clear_sources(self):
        self.sources.clear()
        self._render_sources()

    def add_files(self):
        pattern = " ".join(f"*{e}" for e in sorted(VIDEO_EXT))
        for f in filedialog.askopenfilenames(title="Видео", filetypes=[("Видео", pattern), ("Все файлы", "*.*")]):
            self.add_source(str(Path(f)))

    def add_folder(self):
        d = filedialog.askdirectory(title="Папка с видео")
        if d:
            self.add_source(str(Path(d)))

    # ----- журнал -----
    def log_message(self, text: str):
        """В окно и в файл журнала %LOCALAPPDATA%\\Video2Doc\\video2doc.log."""
        self.msgs.put(("log", text))
        try:
            LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
            with LOG_FILE.open("a", encoding="utf-8") as f:
                f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + text + "\n")
        except OSError:
            pass

    def write_log(self, text: str):
        tag = ("err",) if text.lstrip().startswith(("✖", "ошибка")) else ()
        self.logbox.insert("end", time.strftime("%H:%M:%S  ") + text + "\n", tag)
        self.logbox.see("end")

    def open_log(self):
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        os.startfile(LOG_FILE.parent)

    # ----- обработка -----
    def start(self):
        self.on_step()
        videos = collect_videos(self.sources, self.cfg["recursive"])
        if not videos:
            messagebox.showinfo("Video2Doc", "Видео не найдено. Добавьте файлы или папку.")
            return
        self.worker.stop.clear()
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.model_box.configure(state="disabled")
        self.thread = threading.Thread(target=self.run_all, args=(videos, dict(self.cfg)), daemon=True)
        self.thread.start()

    def run_all(self, videos: list[Path], opts: dict):
        log = self.worker.log
        done = skipped = failed = 0
        log(f"Найдено видео: {len(videos)} · модель: {MODELS[opts['model']][0]} · шаг {opts['step']} с · "
            + ("с кадрами" if opts["frames"] else "только текст"))
        try:
            found = len(videos)
            log("Проверяю программу, доступ к папкам и место на диске…")
            videos = preflight(videos, opts, log)
            failed += found - len(videos)
            log(f"Проверка пройдена, к обработке: {len(videos)}" if videos else "✖ Обрабатывать нечего.")
        except Exception as exc:
            log(f"✖ {exc}" if isinstance(exc, PreflightError) else f"✖ Проверка перед стартом не удалась: {exc}")
            write_error(None, exc, opts)
            videos, failed = [], len(videos)
        for n, video in enumerate(videos, 1):
            if self.worker.stop.is_set():
                break
            out = out_path(video, opts)
            if out.exists() and not opts["overwrite"]:
                skipped += 1
                log(f"[{n}/{len(videos)}] пропуск, уже есть: {out.name}")
                continue
            started = time.perf_counter()
            log(f"[{n}/{len(videos)}] {video}")
            try:
                self.worker.process(video, opts, n, len(videos))
                done += 1
                log(f"    готово за {time.perf_counter() - started:.0f} с → {out}")
            except Cancelled:
                log("    остановлено")
                break
            except Exception as exc:
                failed += 1
                write_error(video, exc, opts)
                log(f"    ошибка: {explain(exc, out)}")
                log(f"    (технически: {exc.__class__.__name__}: {exc})")
        log(f"Итог: сделано {done}, пропущено {skipped}, ошибок {failed}"
            + (" · подробности — «Отчёт для разработчика»" if failed else ""))
        self.msgs.put(("finished", (done, skipped, failed)))

    def cancel(self):
        self.worker.stop.set()
        self.status.set("Останавливаю…")

    def poll(self):
        try:
            while True:
                kind, data = self.msgs.get_nowait()
                if kind == "log":
                    self.write_log(data)
                elif kind == "progress":
                    n, total, frac, stage = data
                    self.bar.set(((n - 1) + frac) / total)
                    self.status.set(f"Файл {n} из {total}: {stage} {frac * 100:.0f}% · "
                                    f"модель: {MODELS[self.cfg['model']][0]}")
                elif kind == "finished":
                    done, skipped, failed = data
                    self.start_btn.configure(state="normal")
                    self.stop_btn.configure(state="disabled")
                    self.model_box.configure(state="readonly")
                    self.bar.set(1)
                    self.status.set(f"Завершено: сделано {done}, пропущено {skipped}, ошибок {failed}")
                    if self.auto:
                        self.root.after(500, self.root.destroy)
        except queue.Empty:
            pass
        self.root.after(150, self.poll)

    def on_close(self):
        if self.thread and self.thread.is_alive():
            if not messagebox.askyesno("Video2Doc", "Идёт обработка. Остановить и выйти?"):
                return
            self.worker.stop.set()
        self.root.destroy()

    def mainloop(self):
        self.root.mainloop()


def main():
    args = [a for a in sys.argv[1:] if a != "--auto"]
    App(args, "--auto" in sys.argv[1:]).mainloop()


if __name__ == "__main__":
    main()
