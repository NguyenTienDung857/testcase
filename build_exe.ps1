param(
    [string]$PythonExe = "d:/project/TESTCASE/JG1#7/.venv/Scripts/python.exe",
    [string]$EntryScript = "fit_images_column_k.py",
    [string]$IconPath = "YuRa - Copy.ico",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

if ($Clean) {
    Remove-Item -Recurse -Force "build" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "dist" -ErrorAction SilentlyContinue
    Remove-Item -Force "FitImagesColumnK.spec" -ErrorAction SilentlyContinue
}

if (-not (Test-Path $PythonExe)) {
    throw "Khong tim thay python tai: $PythonExe"
}

if (-not (Test-Path $EntryScript)) {
    throw "Khong tim thay file script: $EntryScript"
}

if (-not (Test-Path $IconPath)) {
    throw "Khong tim thay icon: $IconPath"
}

& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install -r requirements-build.txt

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name "FitImagesColumnK" `
    --icon "$IconPath" `
    --add-data "YuRa - Copy.ico;." `
    --add-data "templates;templates" `
    --add-data "static;static" `
    "$EntryScript"

Write-Host "Build xong: dist/FitImagesColumnK.exe" -ForegroundColor Green
