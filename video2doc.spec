# -*- mode: python ; coding: utf-8 -*-
"""Сборка: build.ps1 (или pyinstaller video2doc.spec --noconfirm из venv с requirements.txt).

Результат — папка dist\\Video2Doc: Video2Doc.exe + _internal (Python, torch, GigaAM);
bin\\ffmpeg.exe кладёт build.ps1. Модели в сборку не входят: скачиваются при первом запуске.
"""
import os

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# torch требует Visual C++ Runtime; на чистой Windows его может не быть — кладём рядом
_VC_DLLS = ("msvcp140.dll", "msvcp140_1.dll", "msvcp140_2.dll", "msvcp140_atomic_wait.dll",
            "msvcp140_codecvt_ids.dll", "vcruntime140_1.dll", "vcruntime140_threads.dll", "concrt140.dll")
_system32 = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32")
binaries = [(os.path.join(_system32, n), ".") for n in _VC_DLLS if os.path.isfile(os.path.join(_system32, n))]

hiddenimports = ["asr_engine", "model_fetch", "postprocess", "paths", "update"]
for package in ("gigaam", "hydra", "omegaconf", "sentencepiece", "customtkinter", "darkdetect"):
    try:
        hiddenimports += collect_submodules(package)
    except Exception:
        pass

datas = []
for package in ("gigaam", "customtkinter"):
    try:
        datas += collect_data_files(package)
    except Exception:
        pass

a = Analysis(
    ["video2doc_app.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=["matplotlib", "pandas", "IPython", "notebook", "pytest", "tests", "torchvision",
              "tensorboard", "tkinter.test", "pystray", "sounddevice", "soundcard"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="Video2Doc", console=False,
          upx=False, icon="video2doc.ico")
coll = COLLECT(exe, a.binaries, a.datas, upx=False, name="Video2Doc")
