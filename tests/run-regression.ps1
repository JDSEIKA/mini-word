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

    $caretTest = Join-Path $PSScriptRoot "CaretSmoke.cs"
    $caretExe = Join-Path $tmp "CaretSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$caretExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $caretTest
    if ($LASTEXITCODE -ne 0) { throw "caret smoke compile failed" }
    Push-Location $root
    try { & $caretExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "caret smoke failed" }
    $autoSaveTest = Join-Path $PSScriptRoot "AutoSaveAudit.cs"
    $autoSaveExe = Join-Path $tmp "AutoSaveAudit.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$autoSaveExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $autoSaveTest
    if ($LASTEXITCODE -ne 0) { throw "autosave audit compile failed" }
    Push-Location $root
    try { & $autoSaveExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "autosave audit failed" }
    $shortcutTest = Join-Path $PSScriptRoot "ShortcutSmoke.cs"
    $shortcutExe = Join-Path $tmp "ShortcutSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$shortcutExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $shortcutTest
    if ($LASTEXITCODE -ne 0) { throw "shortcut smoke compile failed" }
    Push-Location $root
    try { & $shortcutExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "shortcut smoke failed" }
    $cellFocusTest = Join-Path $PSScriptRoot "CellFocusSmoke.cs"
    $cellFocusExe = Join-Path $tmp "CellFocusSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$cellFocusExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $cellFocusTest
    if ($LASTEXITCODE -ne 0) { throw "cell focus smoke compile failed" }
    Push-Location $root
    try { & $cellFocusExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "cell focus smoke failed" }
    $spellTableTest = Join-Path $PSScriptRoot "SpellTableSmoke.cs"
    $spellTableExe = Join-Path $tmp "SpellTableSmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$spellTableExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $spellTableTest
    if ($LASTEXITCODE -ne 0) { throw "spell table smoke compile failed" }
    Push-Location $root
    try { & $spellTableExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "spell table smoke failed" }
    $securityTest = Join-Path $PSScriptRoot "SecuritySmoke.cs"
    $securityExe = Join-Path $tmp "SecuritySmoke.exe"
    & $csc /nologo /warn:4 /codepage:65001 /target:exe /main:P `
        "/out:$securityExe" `
        /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.Xml.dll /r:System.Xml.Linq.dll `
        $source $securityTest
    if ($LASTEXITCODE -ne 0) { throw "security smoke compile failed" }
    Push-Location $root
    try { & $securityExe }
    finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw "security smoke failed" }

    Write-Output "REGRESSION=PASS"
}
finally {
    Pop-Location
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
