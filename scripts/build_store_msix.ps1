<#
.SYNOPSIS
  Build a full-trust Microsoft Store MSIX package bundling ClickShift.exe + ClickShiftShell.dll.
.PARAMETER Publisher
  Exact Subject DN from Partner Center (e.g. "CN=XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX").
.PARAMETER IdentityName
  Exact Package Identity Name assigned by Partner Center.
.PARAMETER Version
  4-part MSIX version (default matches pyproject.toml + ".0").
#>
param(
    [string]$IdentityName = "ClickShift.Converter",
    [string]$Publisher = "CN=ClickShift",
    [string]$Version = "1.0.20.0",
    [string]$ExePath = ""
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root "assets"
$dist = Join-Path $root "dist"

# 1. Compile ClickShiftShell.dll cleanly if g++ is present
$gpp = Get-Command "g++.exe" -ErrorAction SilentlyContinue
if (-not $gpp) {
    foreach ($cand in @("C:\mingw64\bin\g++.exe", "C:\msys64\mingw64\bin\g++.exe", "C:\msys64\ucrt64\bin\g++.exe")) {
        if (Test-Path $cand) { $gpp = Get-Item $cand; break }
    }
}
if ($gpp) {
    $env:PATH = "$(Split-Path -Parent $gpp.Source);$env:PATH"
    & $gpp.Source -shared -O2 -std=c++17 "$root\src\shell_ext\ClickShiftShell.cpp" "$root\src\shell_ext\ClickShiftShell.def" -o "$assets\ClickShiftShell.dll" -lole32 -luuid -lshlwapi -lshell32 -ladvapi32 -static
    if ($LASTEXITCODE -ne 0) { throw "Failed compiling ClickShiftShell.dll" }
}

# 2. Ensure ClickShift.exe exists
if (-not $ExePath) {
    $candExe = Get-ChildItem "$dist\ClickShift-v*.exe" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($candExe) { $ExePath = $candExe.FullName }
}
if (-not $ExePath -or -not (Test-Path $ExePath)) {
    Push-Location $root
    python build.py
    $POP = $LASTEXITCODE
    Pop-Location
    if ($POP -ne 0) { throw "build.py failed" }
    $candExe = Get-ChildItem "$dist\ClickShift-v*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $ExePath = $candExe.FullName
}

# 3. Locate Windows SDK makeappx.exe
$sdkRoot = "C:\Program Files (x86)\Windows Kits\10\bin"
$makeappx = $null
if (Test-Path $sdkRoot) {
    $makeappx = Get-ChildItem $sdkRoot -Recurse -Filter "makeappx.exe" -ErrorAction SilentlyContinue | Where-Object { $_.FullName -match "x64" } | Select-Object -Last 1
}
if (-not $makeappx) {
    throw "Windows SDK makeappx.exe not found under $sdkRoot. Install Windows 10/11 SDK to build Store MSIX."
}

# 4. Stage Store MSIX layout + resized Store PNG assets
$stage = Join-Path $dist "store_msix_stage"
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force -Path (Join-Path $stage "Assets") | Out-Null

Copy-Item $ExePath (Join-Path $stage "ClickShift.exe") -Force
Copy-Item "$assets\ClickShiftShell.dll" (Join-Path $stage "ClickShiftShell.dll") -Force
if (Test-Path "$assets\clickshift.ico") {
    Copy-Item "$assets\clickshift.ico" (Join-Path $stage "clickshift.ico") -Force
}

$manifestRaw = Get-Content "$root\src\shell_ext\StoreAppxManifest.xml" -Raw -Encoding UTF8
$manifestPatched = $manifestRaw `
    -replace 'Name="ClickShift\.Converter"', "Name=""$IdentityName""" `
    -replace 'Publisher="CN=ClickShift"', "Publisher=""$Publisher""" `
    -replace 'Version="[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+"', "Version=""$Version"""
[System.IO.File]::WriteAllText((Join-Path $stage "AppxManifest.xml"), $manifestPatched, [System.Text.UTF8Encoding]::new($false))

python -c "from pathlib import Path; from PIL import Image; im = Image.open(r'$assets\clickshift.png').convert('RGBA'); out = Path(r'$stage\Assets'); [(im.resize(sz, Image.LANCZOS).save(out / fn)) for fn, sz in [('StoreLogo.png',(50,50)),('Square44x44Logo.png',(44,44)),('Square150x150Logo.png',(150,150))]] ; bg = Image.new('RGBA',(310,150),(20,10,11,255)); lg = im.resize((120,120),Image.LANCZOS); bg.paste(lg,(95,15),lg); bg.save(out / 'Wide310x150Logo.png')"
if ($LASTEXITCODE -ne 0) { throw "Failed generating Store PNG tile sizes" }

New-Item -ItemType Directory -Force -Path $dist | Out-Null
$outMsix = Join-Path $dist "ClickShift-Store-v$($Version -replace '\.0$','').msix"
if (Test-Path $outMsix) { Remove-Item $outMsix -Force }
& $makeappx.FullName pack /d $stage /p $outMsix /o
if ($LASTEXITCODE -ne 0) { throw "makeappx pack failed" }

Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Built Store MSIX: $outMsix"
