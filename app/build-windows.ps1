$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$sidecarDir = Join-Path $PSScriptRoot 'src-tauri\binaries'
$cargoBin = Join-Path $env:USERPROFILE '.cargo\bin'
if (Test-Path -LiteralPath $cargoBin) { $env:PATH = "$cargoBin;$env:PATH" }
if (-not (Get-Command link.exe -ErrorAction SilentlyContinue)) {
    $vswhere = 'C:\Program Files (x86)\Microsoft Visual Studio\Installer\vswhere.exe'
    if (-not (Test-Path -LiteralPath $vswhere)) { throw 'Не найдены Visual Studio C++ Build Tools.' }
    $installation = & $vswhere -latest -products '*' -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
    if (-not $installation) { throw 'Не установлен компонент Visual Studio C++ Build Tools.' }
    $devCommand = Join-Path $installation 'Common7\Tools\VsDevCmd.bat'
    $environmentLines = & cmd.exe /d /c "`"$devCommand`" -arch=amd64 >nul && set"
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось подготовить окружение MSVC.' }
    foreach ($line in $environmentLines) {
        if ($line -match '^([^=]+)=(.*)$') {
            [Environment]::SetEnvironmentVariable($matches[1], $matches[2], 'Process')
        }
    }
    if (-not (Get-Command link.exe -ErrorAction SilentlyContinue)) { throw 'Не найден MSVC link.exe после настройки окружения.' }
}

foreach ($required in @((Join-Path $projectRoot 'excel_loading.ipynb'), (Join-Path $projectRoot 'resources\nephrolife_logo.jpg'))) {
    if (-not (Test-Path -LiteralPath $required)) { throw "Не найден ресурс сборки: $required" }
}

Push-Location -LiteralPath $projectRoot
try {
    python -m PyInstaller --onefile --noconfirm --clean --name papagun-engine `
        --distpath $sidecarDir --workpath (Join-Path $PSScriptRoot '.pyinstaller-build') `
        --specpath $PSScriptRoot --hidden-import pandas --hidden-import pptx `
        --add-data "$(Join-Path $projectRoot 'excel_loading.ipynb');." `
        --add-data "$(Join-Path $projectRoot 'resources\nephrolife_logo.jpg');." `
        (Join-Path $PSScriptRoot 'engine.py')
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать локальный обработчик.' }

    Copy-Item -LiteralPath (Join-Path $sidecarDir 'papagun-engine.exe') `
        -Destination (Join-Path $sidecarDir 'papagun-engine-x86_64-pc-windows-msvc.exe') -Force
    Push-Location -LiteralPath $PSScriptRoot
    try {
        npm.cmd ci
        if ($LASTEXITCODE -ne 0) { throw 'Не удалось установить зависимости интерфейса.' }
        npm.cmd exec tauri -- build
        if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать Tauri; проверьте Rust и Visual Studio C++ Build Tools.' }
    } finally { Pop-Location }
} finally { Pop-Location }
