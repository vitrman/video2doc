"""Обновление Video2Doc из GitHub Releases (по образцу update.py из Voice2Text).

Установщика у Video2Doc нет — программа лежит папкой. Поэтому:
1. check() читает последний релиз через публичный API (без токена) и сравнивает версии;
2. download() качает Video2Doc-<версия>-win64.zip в %TEMP%\\video2doc-update, сверяет sha256;
3. apply() распаковывает архив и запускает из него новую Video2Doc.exe --finish-update: она ждёт
   выхода программы, заменяет её файлы (старые переименовываются в *.old и при ошибке
   возвращаются обратно), записывает итог в update-result.txt и снова запускает Video2Doc.exe.
Папка models и любые другие файлы, которых нет в архиве, не трогаются; настройки, журналы
и скачанные модели лежат в %LOCALAPPDATA%\\Video2Doc и тоже не затрагиваются.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

REPO = "vitrman/video2doc"
API = "https://api.github.com/repos/{repo}/releases/latest"
RELEASES = "https://api.github.com/repos/{repo}/releases"
ASSET_RE = re.compile(r"^Video2Doc-(\d+(?:\.\d+)+)-win64\.zip$", re.IGNORECASE)
EXE_NAME = "Video2Doc.exe"
USER_AGENT = "Video2Doc-updater"
TIMEOUT = 20
CHUNK = 1 << 20
CREATE_NO_WINDOW = 0x08000000
DETACHED_PROCESS = 0x00000008


class UpdateError(RuntimeError):
    """Понятная человеку ошибка обновления."""


class UpdateCancelled(UpdateError):
    """Пользователь отменил скачивание."""


def version_tuple(text: str) -> tuple[int, ...]:
    """`1.4` -> (1, 4, 0, 0); лишнее в конце игнорируется."""
    numbers = [int(part) for part in re.findall(r"\d+", str(text))][:4]
    return tuple(numbers + [0] * (4 - len(numbers)))


def _request(url: str):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT,
                                                   "Accept": "application/vnd.github+json"})
    try:
        return urllib.request.urlopen(request, timeout=TIMEOUT)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise UpdateError("на GitHub пока нет опубликованных версий") from exc
        if exc.code == 403:
            raise UpdateError("GitHub временно ограничил число запросов, попробуйте позже") from exc
        raise UpdateError(f"GitHub ответил ошибкой {exc.code}") from exc
    except (urllib.error.URLError, OSError) as exc:
        raise UpdateError(f"нет связи с GitHub ({getattr(exc, 'reason', exc)})") from exc


def _sha_from_release(data: dict, asset: dict) -> str:
    """sha256 архива: поле digest у файла или соседний `<архив>.sha256`."""
    digest = str(asset.get("digest") or "")
    if digest.lower().startswith("sha256:"):
        return digest.split(":", 1)[1].strip().lower()
    want = str(asset.get("name") or "").lower() + ".sha256"
    for item in data.get("assets") or []:
        if str(item.get("name") or "").lower() != want:
            continue
        try:
            request = urllib.request.Request(item.get("browser_download_url") or "",
                                             headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
                parts = response.read(4096).decode("utf-8", "replace").split()
        except Exception:
            return ""
        if parts and re.fullmatch(r"[0-9a-fA-F]{64}", parts[0]):
            return parts[0].lower()
    return ""


def _parse_release(data: dict) -> dict | None:
    if data.get("draft") or data.get("prerelease"):
        return None
    for asset in data.get("assets") or []:
        match = ASSET_RE.match(str(asset.get("name") or ""))
        if match:
            return {"version": match.group(1), "name": asset["name"], "url": asset["browser_download_url"],
                    "size": int(asset.get("size") or 0), "sha256": _sha_from_release(data, asset),
                    "page": data.get("html_url") or "", "notes": str(data.get("body") or "")}
    return None


def fetch_latest(repo: str = REPO) -> dict:
    """Последний релиз с архивом программы."""
    with _request(API.format(repo=repo)) as response:
        found = _parse_release(json.loads(response.read().decode("utf-8")))
    if found:
        return found
    with _request(RELEASES.format(repo=repo)) as response:
        for data in json.loads(response.read().decode("utf-8")):
            found = _parse_release(data) if isinstance(data, dict) else None
            if found:
                return found
    raise UpdateError("в релизах нет архива Video2Doc-<версия>-win64.zip")


def check(current: str, repo: str = REPO) -> dict | None:
    """Описание обновления, если на GitHub версия новее текущей, иначе None."""
    latest = fetch_latest(repo)
    return latest if version_tuple(latest["version"]) > version_tuple(current) else None


def cache_dir() -> Path:
    folder = Path(tempfile.gettempdir()) / "video2doc-update"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for block in iter(lambda: fh.read(CHUNK), b""):
            digest.update(block)
    return digest.hexdigest()


def download(info: dict, on_progress=None, cancel=None) -> Path:
    """Качает архив релиза; уже скачанный и проверенный повторно не качает."""
    target = cache_dir() / Path(info["name"]).name
    if target.is_file() and (not info["size"] or target.stat().st_size == info["size"]) \
            and (not info["sha256"] or _sha256(target) == info["sha256"]):
        return target
    part = target.with_name(target.name + ".part")
    digest = hashlib.sha256()
    request = urllib.request.Request(info["url"], headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response, open(part, "wb") as out:
            total = info["size"] or int(response.headers.get("Content-Length") or 0)
            done = 0
            while True:
                if cancel is not None and cancel.is_set():
                    raise UpdateCancelled("скачивание обновления отменено")
                block = response.read(CHUNK)
                if not block:
                    break
                out.write(block)
                digest.update(block)
                done += len(block)
                if on_progress is not None:
                    on_progress(done, total)
    except UpdateCancelled:
        part.unlink(missing_ok=True)
        raise
    except (urllib.error.URLError, OSError) as exc:
        part.unlink(missing_ok=True)
        raise UpdateError(f"скачивание прервалось ({getattr(exc, 'reason', exc)})") from exc
    if info["size"] and part.stat().st_size != info["size"]:
        part.unlink(missing_ok=True)
        raise UpdateError("архив скачался не целиком, попробуйте ещё раз")
    if info["sha256"] and digest.hexdigest() != info["sha256"]:
        part.unlink(missing_ok=True)
        raise UpdateError("архив обновления повреждён (не совпала контрольная сумма), попробуйте ещё раз")
    part.replace(target)
    return target


def unpack(archive: Path) -> Path:
    """Распаковывает архив и возвращает папку, в которой лежит Video2Doc.exe."""
    folder = cache_dir() / f"unpacked-{archive.stem}"
    shutil.rmtree(folder, ignore_errors=True)
    try:
        with zipfile.ZipFile(archive) as bundle:
            if bundle.testzip():
                raise UpdateError("архив обновления повреждён")
            bundle.extractall(folder)
    except zipfile.BadZipFile as exc:
        archive.unlink(missing_ok=True)
        raise UpdateError(f"архив обновления повреждён: {exc}") from exc
    for exe in folder.rglob(EXE_NAME):
        return exe.parent
    raise UpdateError(f"в архиве обновления нет {EXE_NAME}")


def apply(archive: Path, target: Path, result_file: Path, version: str, pid: int | None = None) -> None:
    """Распаковывает архив и запускает из него новую Video2Doc.exe в режиме --finish-update.

    Она ждёт выхода этой программы (pid), заменяет файлы в target, пишет итог в result_file
    и запускает обновлённую программу. PowerShell и bat-файлы не нужны: аргументы процесса —
    Unicode, кириллица в путях не ломается, а антивирусы не путают это со скриптом.
    """
    source = unpack(archive)
    result_file.parent.mkdir(parents=True, exist_ok=True)
    result_file.unlink(missing_ok=True)
    subprocess.Popen([str(source / EXE_NAME), "--finish-update", str(target), str(int(pid or os.getpid())),
                      str(result_file), version],
                     cwd=str(source), creationflags=DETACHED_PROCESS, close_fds=True)


def _wait_pid(pid: int, timeout: float = 180) -> None:
    """Ждёт завершения процесса (Windows API, без сторонних библиотек)."""
    import ctypes
    kernel = ctypes.windll.kernel32
    handle = kernel.OpenProcess(0x00100000, False, pid)  # SYNCHRONIZE
    if handle:
        kernel.WaitForSingleObject(handle, int(timeout * 1000))
        kernel.CloseHandle(handle)


def _retry(action, attempts: int = 60) -> None:
    """Антивирус может ещё держать файлы только что закрытой программы — ждём до минуты."""
    for attempt in range(1, attempts + 1):
        try:
            action()
            return
        except OSError:
            if attempt == attempts:
                raise
            time.sleep(1)


def _remove(path: Path) -> None:
    if path.is_dir():
        shutil.rmtree(path)
    elif path.exists():
        path.unlink()


def finish_update(source: Path, target: Path, pid: int, result_file: Path, version: str) -> None:
    """Вторая половина обновления: выполняется новой версией, запущенной из распакованного архива."""
    _wait_pid(pid)
    time.sleep(0.7)
    moved: list[str] = []
    try:
        for item in sorted(source.iterdir()):
            dest = target / item.name
            if dest.exists():
                old = target / (item.name + ".old")
                if old.exists():
                    _retry(lambda: _remove(old))
                _retry(lambda: os.replace(dest, old))
                moved.append(item.name)
            if item.is_dir():
                shutil.copytree(item, dest)
            else:
                shutil.copy2(item, dest)
        for name in moved:          # *.old, которые не удалились сейчас, уберёт cleanup() при запуске
            try:
                _remove(target / (name + ".old"))
            except OSError:
                pass
        result = f"ok {version}"
    except Exception as exc:  # откат: возвращаем старые файлы на место
        for name in moved:
            old, dest = target / (name + ".old"), target / name
            try:
                _remove(dest)
                os.replace(old, dest)
            except OSError:
                pass
        result = f"error {version} {exc}"
    try:
        result_file.write_text(result, encoding="utf-8")
    except OSError:
        pass
    try:
        subprocess.Popen([str(target / EXE_NAME)], cwd=str(target), creationflags=DETACHED_PROCESS, close_fds=True)
    except OSError:
        pass  # не запустилась — человек откроет сам, итог обновления уже записан


def read_result(result_file: Path) -> str:
    """Итог прошлого обновления ('ok 1.5' / 'error 1.5 …') — один раз, файл удаляется."""
    try:
        text = result_file.read_text(encoding="utf-8-sig").strip()
        result_file.unlink(missing_ok=True)
        return text
    except OSError:
        return ""


def cleanup(target: Path, current_version: str) -> None:
    """Убирает *.old в папке программы и архивы не новее текущей версии в кэше."""
    for old in target.glob("*.old"):
        try:
            shutil.rmtree(old) if old.is_dir() else old.unlink()
        except OSError:
            pass
    folder = Path(tempfile.gettempdir()) / "video2doc-update"
    if not folder.is_dir():
        return
    now = time.time()
    for entry in folder.iterdir():
        match = re.search(r"(\d+(?:\.\d+)+)", entry.name)
        stale = now - entry.stat().st_mtime > 7 * 86400
        if stale or entry.name.endswith(".part") or (
                match and version_tuple(match.group(1)) <= version_tuple(current_version)):
            shutil.rmtree(entry, ignore_errors=True) if entry.is_dir() else entry.unlink(missing_ok=True)
