$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$mainApp = Join-Path $projectRoot 'app'
$stage = Join-Path $PSScriptRoot '.build\app'
$sidecarDir = Join-Path $stage 'src-tauri\binaries'
$releaseDir = Join-Path $stage 'src-tauri\target\release'
$destination = Join-Path $PSScriptRoot 'dist'

if (Test-Path -LiteralPath $stage) {
    $resolvedStage = [IO.Path]::GetFullPath($stage)
    $allowedRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '.build'))
    if (-not $resolvedStage.StartsWith($allowedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Небезопасный путь каталога demo-сборки.'
    }
    Remove-Item -LiteralPath $stage -Recurse -Force
}
New-Item -ItemType Directory -Path $stage, (Join-Path $stage 'src-tauri'), $sidecarDir, $destination -Force | Out-Null

foreach ($name in @('src', 'index.html', 'vite.config.ts', 'tsconfig.json', 'package.json', 'package-lock.json', 'engine.py')) {
    Copy-Item -LiteralPath (Join-Path $mainApp $name) -Destination $stage -Recurse -Force
}
foreach ($name in @('src', 'icons', 'capabilities', 'Cargo.toml', 'Cargo.lock', 'build.rs', 'tauri.conf.json')) {
    Copy-Item -LiteralPath (Join-Path (Join-Path $mainApp 'src-tauri') $name) -Destination (Join-Path $stage 'src-tauri') -Recurse -Force
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'watermark.py') -Destination (Join-Path $stage 'watermark.py')

$enginePath = Join-Path $stage 'engine.py'
$engine = [IO.File]::ReadAllText($enginePath)
$needle = '    result = ns["combine_created_slides"](payloads)'
if ($engine.Split(@($needle), [StringSplitOptions]::None).Length -ne 2) { throw 'Точка добавления водяного знака в engine.py изменилась.' }
$engine = $engine.Replace($needle, "$needle`n    from watermark import apply_watermark`n    result = apply_watermark(result)")
[IO.File]::WriteAllText($enginePath, $engine, [Text.UTF8Encoding]::new($false))

$rustPath = Join-Path $stage 'src-tauri\src\main.rs'
$rust = [IO.File]::ReadAllText($rustPath).Replace('papagun-engine.exe', 'papagun-demo-engine.exe')
[IO.File]::WriteAllText($rustPath, $rust, [Text.UTF8Encoding]::new($false))
$cargoPath = Join-Path $stage 'src-tauri\Cargo.toml'
$cargo = [IO.File]::ReadAllText($cargoPath).Replace('name = "papagun"', 'name = "papagun-demo"')
[IO.File]::WriteAllText($cargoPath, $cargo, [Text.UTF8Encoding]::new($false))
$configPath = Join-Path $stage 'src-tauri\tauri.conf.json'
$config = [IO.File]::ReadAllText($configPath).Replace('"Papagun"', '"Papagun Demo"').Replace('kz.papagun.desktop', 'kz.papagun.demo').Replace('binaries/papagun-engine', 'binaries/papagun-demo-engine')
[IO.File]::WriteAllText($configPath, $config, [Text.UTF8Encoding]::new($false))
$uiPath = Join-Path $stage 'src\main.tsx'
$ui = [IO.File]::ReadAllText($uiPath).Replace('<strong>Papagun</strong>', '<strong>Papagun Demo</strong>').Replace("Papagun_отчёт.pptx", "Papagun_Demo_отчёт.pptx")
[IO.File]::WriteAllText($uiPath, $ui, [Text.UTF8Encoding]::new($false))

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
        if ($line -match '^([^=]+)=(.*)$') { [Environment]::SetEnvironmentVariable($matches[1], $matches[2], 'Process') }
    }
}

Push-Location -LiteralPath $projectRoot
try {
    python -m PyInstaller --onefile --noconfirm --clean --name papagun-demo-engine `
        --distpath $sidecarDir --workpath (Join-Path $stage '.pyinstaller-build') `
        --specpath $stage --hidden-import pandas --hidden-import pptx `
        --add-data "$(Join-Path $projectRoot 'excel_loading.ipynb');." `
        --add-data "$(Join-Path $projectRoot 'resources\nephrolife_logo.jpg');." `
        $enginePath
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать demo-обработчик.' }
    Copy-Item -LiteralPath (Join-Path $sidecarDir 'papagun-demo-engine.exe') `
        -Destination (Join-Path $sidecarDir 'papagun-demo-engine-x86_64-pc-windows-msvc.exe') -Force
} finally { Pop-Location }

Push-Location -LiteralPath $stage
try {
    npm.cmd ci
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось установить зависимости интерфейса demo.' }
    npm.cmd exec tauri -- build
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось собрать demo-приложение.' }
} finally { Pop-Location }

$installer = Get-ChildItem -LiteralPath (Join-Path $releaseDir 'bundle\nsis') -Filter '*-setup.exe' | Select-Object -First 1
if (-not $installer) { throw 'Demo-установщик не найден.' }
Copy-Item -LiteralPath $installer.FullName -Destination (Join-Path $destination 'Papagun-Demo-Setup.exe') -Force

$portableFiles = @('papagun-demo.exe', 'papagun-demo-engine.exe')
$portableReady = $true
foreach ($name in $portableFiles) {
    $path = Join-Path $destination $name
    if (-not (Test-Path -LiteralPath $path)) { continue }
    try {
        $handle = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $handle.Dispose()
    } catch {
        $portableReady = $false
        break
    }
}
if ($portableReady) {
    Copy-Item -LiteralPath (Join-Path $releaseDir 'papagun-demo.exe') -Destination (Join-Path $destination 'papagun-demo.exe') -Force
    Copy-Item -LiteralPath (Join-Path $sidecarDir 'papagun-demo-engine.exe') -Destination (Join-Path $destination 'papagun-demo-engine.exe') -Force
} else {
    $portableDir = Join-Path $destination ('portable-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $portableDir -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $releaseDir 'papagun-demo.exe') -Destination (Join-Path $portableDir 'papagun-demo.exe')
    Copy-Item -LiteralPath (Join-Path $sidecarDir 'papagun-demo-engine.exe') -Destination (Join-Path $portableDir 'papagun-demo-engine.exe')
    Write-Warning "Запущенная demo-версия не позволяет обновить переносимые EXE в dist. Новая пара: $portableDir"
}
Write-Output "Demo: $destination"
