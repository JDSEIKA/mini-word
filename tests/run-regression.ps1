$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$tmp = Join-Path $root "tmp-regression"

Push-Location $root
try {
    & cmd /c build.cmd
    if ($LASTEXITCODE -ne 0) { throw "build failed" }

    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force (Join-Path $tmp "test-assets") | Out-Null
    Copy-Item ".\test-assets\sample.doc", ".\test-assets\ascii.doc" (Join-Path $tmp "test-assets") -Force

    Push-Location $tmp
    try {
        & (Join-Path $root "dist\MiniWord.exe") --selftest (Join-Path $tmp "self")
        if ($LASTEXITCODE -ne 0) { throw "selftest failed" }
    }
    finally { Pop-Location }

    $source = Get-ChildItem (Join-Path $root "src") -Recurse -Filter *.cs |
        Where-Object { $_.Name -ne "Program.cs" } |
        ForEach-Object { $_.FullName }
    $test = Join-Path $PSScriptRoot "CorpusSmoke.cs"
    $exe = Join-Path $tmp "CorpusSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$exe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $test
    if ($LASTEXITCODE -ne 0) { throw "smoke compile failed" }

    Push-Location $root
    try { & (Join-Path $tmp "CorpusSmoke.exe") }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "corpus smoke failed" }

    $uiTest = Join-Path $PSScriptRoot "UiSmoke.cs"
    $uiExe = Join-Path $tmp "UiSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$uiExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $uiTest
    if ($LASTEXITCODE -ne 0) { throw "UI smoke compile failed" }
    Push-Location $root
    try { & $uiExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "UI smoke failed" }

    $clipTest = Join-Path $PSScriptRoot "ClipboardSmoke.cs"
    $clipExe = Join-Path $tmp "ClipboardSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$clipExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $clipTest
    if ($LASTEXITCODE -ne 0) { throw "clipboard smoke compile failed" }
    Push-Location $root
    try { & $clipExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "clipboard smoke failed" }

    $featureTest = Join-Path $PSScriptRoot "FeatureAudit.cs"
    $featureExe = Join-Path $tmp "FeatureAudit.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$featureExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $featureTest
    if ($LASTEXITCODE -ne 0) { throw "feature audit compile failed" }
    Push-Location $root
    try { & $featureExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "feature audit failed" }

    Write-Output "REGRESSION=PASS"
}
finally {
    Pop-Location
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
