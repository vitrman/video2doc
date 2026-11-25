# Сторонние компоненты

Код проекта распространяется по [MIT](LICENSE). Ниже — что попадает в сборку или
скачивается при первом запуске.

## Распознавание речи

| Компонент | Лицензия | Где используется |
|---|---|---|
| [GigaAM](https://github.com/salute-developers/GigaAM) v3 (код и веса `v3_e2e_rnnt`, `v3_e2e_ctc`) | MIT, © SberDevices | распознавание речи |
| [Voice2Text](https://github.com/vitrman/voice2text): `asr_engine.py`, `model_fetch.py`, `postprocess.py`, `paths.py` | MIT | загрузка модели, нарезка звука, чистка «э-э» |
| PyTorch, torchaudio | BSD | исполнение модели |
| sentencepiece | Apache-2.0 | токенизатор модели |
| omegaconf | BSD-3-Clause | конфиг модели |
| hydra-core | MIT | конфиг модели |
| NumPy, SciPy | BSD-3-Clause | обработка сигнала |

Веса моделей (~450 МБ каждая) не входят ни в репозиторий, ни в архив релиза: программа
скачивает выбранную модель при первом запуске с CDN Сбера, а если он недоступен — из релиза
[`models-v1`](https://github.com/vitrman/voice2text/releases/tag/models-v1) репозитория Voice2Text.

## Видео

| Компонент | Лицензия | Где используется |
|---|---|---|
| [FFmpeg](https://ffmpeg.org) (`bin\ffmpeg.exe`, сборка [gyan.dev](https://www.gyan.dev/ffmpeg/builds/)) | GPL-3.0 | звук и кадры из видео |

`ffmpeg.exe` поставляется отдельным исполняемым файлом без изменений и вызывается как внешняя
программа. Исходный код FFmpeg: <https://ffmpeg.org/download.html>, конфигурация сборки —
`bin\ffmpeg.exe -version`. Текст лицензии: <https://www.gnu.org/licenses/gpl-3.0.html>.

## Окно и сборка

| Компонент | Лицензия | Где используется |
|---|---|---|
| CustomTkinter | CC0-1.0 | оформление окна |
| darkdetect | BSD-3-Clause | тема CustomTkinter |
| Pillow | MIT-CMU (HPND) | рисование иконки (`make_icon.py`) |
| Microsoft Visual C++ Runtime (`msvcp140*.dll`, `vcruntime140_1.dll`, `concrt140.dll`) | распространяемые компоненты Microsoft | нужны torch на чистой Windows |
| PyInstaller (в поставку не входит) | GPL-2.0 с исключением для собранных приложений | сборка exe |
