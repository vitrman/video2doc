# Сборка Video2Doc: exe + ffmpeg в одну переносимую папку dist\Video2Doc.
# Модели GigaAM в сборку не входят: программа скачивает их при первом запуске.
#   .\build.ps1                         — собрать папку dist\Video2Doc
#   .\build.ps1 -Release                — плюс архив dist\Video2Doc-<версия>-win64.zip и .sha256 для релиза
#   .\build.ps1 -WithModels <папка>     — положить модели из папки в dist\Video2Doc\models (сборка без интернета)
#   -Python <python.exe>  интерпретатор с пакетами из requirements.txt (по умолчанию .venv\Scripts\python.exe)
#   -Stage <папка>        где PyInstaller собирает промежуточные файлы (по умолчанию build\)
#   -Ffmpeg <ffmpeg.exe>  какой ffmpeg положить в bin\ (по умолчанию из PATH)
param(
    [string]$Python = ".venv\Scripts\python.exe",
    [string]$Stage = "build",
    [string]$Ffmpeg = "",
    [string]$WithModels = "",
    [switch]$Release
)
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
if (-not [IO.Path]::IsPathRooted($Stage)) { $Stage = Join-Path $PSScriptRoot $Stage }
New-Item -ItemType Directory -Force "$Stage\tmp" | Out-Null
$env:TEMP = "$Stage\tmp"; $env:TMP = $env:TEMP
& $Python -m PyInstaller video2doc.spec --noconfirm --distpath "$Stage\dist" --workpath "$Stage\work"
if ($LASTEXITCODE) { throw "PyInstaller завершился с кодом $LASTEXITCODE" }

$out = "dist\Video2Doc"
if (Test-Path $out) {
    # файл старой сборки может держать антивирус — тогда отодвигаем папку в сторону
    try { Remove-Item -Recurse -Force $out } catch { Rename-Item $out ("Video2Doc_old_" + (Get-Date -Format yyyyMMdd_HHmmss)) }
}
New-Item -ItemType Directory -Force dist | Out-Null
Move-Item "$Stage\dist\Video2Doc" $out
Remove-Item -Recurse -Force "$Stage\dist", "$Stage\work", "$Stage\tmp" -ErrorAction SilentlyContinue

New-Item -ItemType Directory -Force "$out\bin" | Out-Null
if (-not $Ffmpeg) { $Ffmpeg = (Get-Command ffmpeg -ErrorAction Stop).Source }
Copy-Item $Ffmpeg "$out\bin\ffmpeg.exe" -Force
if ($WithModels) {
    New-Item -ItemType Directory -Force "$out\models" | Out-Null
    foreach ($m in "v3_e2e_rnnt", "v3_e2e_ctc") {
        Copy-Item "$WithModels\$m.ckpt", "$WithModels\${m}_tokenizer.model" "$out\models\" -Force
    }
}
Copy-Item video2doc.ico, README.txt, LICENSE, THIRD_PARTY.md "$out\" -Force

# подпись локальным сертификатом, если он есть (только для своей машины: антивирус спокойнее к подписанным exe)
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object { $_.Subject -like "*CN=Voice2Text*" } | Select-Object -First 1
if ($cert) { "Подпись: " + (Set-AuthenticodeSignature "$out\Video2Doc.exe" -Certificate $cert).Status } else { "Подпись: сертификата нет, exe не подписан" }

if ($Release) {
    if ($WithModels) { throw "Архив для релиза собирается без моделей: уберите -WithModels" }
    $version = (Select-String -Path video2doc_app.py -Pattern '^VERSION = "(.+)"').Matches[0].Groups[1].Value
    $zip = "dist\Video2Doc-$version-win64.zip"
    if (Test-Path $zip) { Remove-Item $zip }
    Compress-Archive -Path $out -DestinationPath $zip -CompressionLevel Optimal
    $hash = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
    "$hash  $(Split-Path $zip -Leaf)" | Set-Content "$zip.sha256" -Encoding ascii
    "Архив: $PSScriptRoot\$zip ({0:N0} МБ)" -f ((Get-Item $zip).Length / 1MB)
}
"Готово: $PSScriptRoot\$out"
