param(
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA "MiniWord")
)

$ErrorActionPreference = "Stop"
$source = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "MiniWord.exe"
if (-not (Test-Path -LiteralPath $source)) {
    throw "No se encontró MiniWord.exe junto a este instalador."
}

New-Item -ItemType Directory -Force -Path $InstallDirectory | Out-Null
$target = Join-Path $InstallDirectory "MiniWord.exe"
Copy-Item -LiteralPath $source -Destination $target -Force

$programs = [Environment]::GetFolderPath("Programs")
$shortcutPath = Join-Path $programs "Mini Word.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $target
$shortcut.WorkingDirectory = $InstallDirectory
$shortcut.Description = "Mini Word"
$shortcut.Save()

Set-Content -LiteralPath (Join-Path $InstallDirectory "VERSION.txt") `
    -Value "Mini Word 0.0.16 Beta (2026-11-25)" -Encoding UTF8

Write-Host "Mini Word instalado en $InstallDirectory"
Write-Host "Acceso directo: $shortcutPath"
Write-Host "Este instalador portátil no está firmado digitalmente."
