#requires -version 2
<#
  Compila Mini Word sin SDK ni paquetes NuGet.
  Requiere el compilador de C# incluido en .NET Framework 4.x (Windows 7+).
#>
[CmdletBinding()]
param(
    [switch]$RunSelfTest
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Definition
$cscCandidates = @(
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'),
    (Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe')
)
$csc = $cscCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $csc) {
    throw 'No se encontró csc.exe de .NET Framework 4.x. Instale .NET Framework 4.0 o posterior.'
}

$dist = Join-Path $root 'dist'
$exe = Join-Path $dist 'MiniWord.exe'
$icon = Join-Path $root 'assets\MiniWord.ico'
$manifest = Join-Path $root 'assets\app.manifest'
New-Item -ItemType Directory -Force -Path $dist | Out-Null

$sources = @(Get-ChildItem -Path (Join-Path $root 'src') -Recurse -Filter '*.cs' |
    Sort-Object FullName | ForEach-Object { $_.FullName })
if ($sources.Count -eq 0) { throw 'No se encontraron archivos .cs en src\.' }

$references = @(
    '/r:System.dll',
    '/r:System.Core.dll',
    '/r:System.Drawing.dll',
    '/r:System.Windows.Forms.dll',
    '/r:System.Xml.dll',
    '/r:System.Xml.Linq.dll'
)

$arguments = @(
    '/nologo',
    '/target:winexe',
    '/optimize+',
    '/platform:anycpu',
    '/codepage:65001',
    '/utf8output',
    ('/out:"' + $exe + '"'),
    ('/win32icon:"' + $icon + '"'),
    ('/win32manifest:"' + $manifest + '"')
) + $references + $sources

Write-Host 'Compilando Mini Word...' -ForegroundColor Cyan
& $csc $arguments
if ($LASTEXITCODE -ne 0) { throw "La compilación falló con el código $LASTEXITCODE." }

# La autoprueba usa un .doc real como fixture. Se copia junto al ejecutable
# para que `MiniWord.exe --selftest` funcione también desde dist\.
$fixtures = Join-Path $root 'test-assets'
if (Test-Path $fixtures) {
    $target = Join-Path $dist 'test-assets'
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    Copy-Item (Join-Path $fixtures 'sample.doc') $target -Force
    if (Test-Path (Join-Path $fixtures 'ascii.doc')) {
        Copy-Item (Join-Path $fixtures 'ascii.doc') $target -Force
    }
}

$size = [Math]::Round((Get-Item $exe).Length / 1KB, 0)
Write-Host ("Listo: {0} ({1} KB)" -f $exe, $size) -ForegroundColor Green

if ($RunSelfTest) {
    $testDir = Join-Path $root 'test-output'
    if (Test-Path $testDir) { Remove-Item -LiteralPath $testDir -Recurse -Force }
    $log = Join-Path $testDir 'selftest.log'
    $process = Start-Process -FilePath $exe -ArgumentList @('--selftest', $testDir) -Wait -PassThru
    if (Test-Path $log) { Get-Content -LiteralPath $log }
    if ($process.ExitCode -ne 0) { throw "La autoprueba falló con el código $($process.ExitCode)." }
    Write-Host 'Autoprueba correcta.' -ForegroundColor Green
}
