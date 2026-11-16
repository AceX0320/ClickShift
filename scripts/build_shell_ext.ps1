$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root "assets"
$srcExt = Join-Path $root "src\shell_ext"

New-Item -ItemType Directory -Force -Path $assets | Out-Null

# 1. Compile ClickShiftShell.dll and ClickShiftStub.exe if g++ is available
$gpp = Get-Command "g++.exe" -ErrorAction SilentlyContinue
if (-not $gpp -and (Test-Path "C:\mingw64\bin\g++.exe")) {
    $gpp = Get-Item "C:\mingw64\bin\g++.exe"
}
if ($gpp) {
    Write-Host "Compiling ClickShiftShell.dll and ClickShiftStub.exe with $($gpp.Source)..."
    & $gpp.Source -shared -O2 -std=c++17 "$srcExt\ClickShiftShell.cpp" "$srcExt\ClickShiftShell.def" -o "$assets\ClickShiftShell.dll" -lole32 -luuid -lshlwapi -lshell32 -ladvapi32 -static
    & $gpp.Source -O2 -municode -mwindows "$srcExt\ClickShiftStub.cpp" -o "$assets\ClickShiftStub.exe" -static
}

Copy-Item "$srcExt\AppxManifest.xml" "$assets\AppxManifest.xml" -Force

# 2. Ensure clickshift.png exists
if (-not (Test-Path "$assets\clickshift.png")) {
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap(150, 150)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 37, 99, 235))
    $g.FillRectangle($brush, 10, 10, 130, 130)
    $wBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $g.FillRectangle($wBrush, 42, 30, 50, 90)
    $yBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 214, 10))
    $g.FillRectangle($yBrush, 92, 58, 36, 12)
    $bmp.Save("$assets\clickshift.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose()
    $bmp.Dispose()
}

# 3. Pack and sign ClickShiftShell.msix if Windows SDK tools (makeappx & signtool) are available
$sdkRoot = "C:\Program Files (x86)\Windows Kits\10\bin"
$makeappx = $null
$signtool = $null
if (Test-Path $sdkRoot) {
    $makeappx = Get-ChildItem $sdkRoot -Recurse -Filter "makeappx.exe" -ErrorAction SilentlyContinue | Where-Object { $_.FullName -match "x64" } | Select-Object -Last 1
    $signtool = Get-ChildItem $sdkRoot -Recurse -Filter "signtool.exe" -ErrorAction SilentlyContinue | Where-Object { $_.FullName -match "x64" } | Select-Object -Last 1
}

if ($makeappx -and $signtool) {
    Write-Host "Building and signing ClickShiftShell.msix using Windows SDK..."
    $stageDir = Join-Path $env:TEMP "clickshift_msix_stage"
    if (Test-Path $stageDir) { Remove-Item $stageDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path $stageDir | Out-Null

    Copy-Item "$assets\AppxManifest.xml" "$stageDir\AppxManifest.xml" -Force
    Copy-Item "$assets\ClickShiftShell.dll"  "$stageDir\ClickShiftShell.dll"  -Force
    Copy-Item "$assets\ClickShiftStub.exe"   "$stageDir\ClickShiftStub.exe"   -Force
    Copy-Item "$assets\clickshift.png"       "$stageDir\clickshift.png"       -Force
    if (Test-Path "$assets\clickshift.ico") {
        Copy-Item "$assets\clickshift.ico"   "$stageDir\clickshift.ico"       -Force
    }

    $cert = New-SelfSignedCertificate -Type Custom -Subject "CN=ClickShift" -KeyUsage DigitalSignature -FriendlyName "ClickShift Shell Extension" -CertStoreLocation "Cert:\CurrentUser\My" -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3", "2.5.29.19={text}")
    $cerPath = Join-Path $assets "ClickShiftShell.cer"
    $pfxPath = Join-Path $env:TEMP "ClickShiftShell.pfx"
    $pwd = ConvertTo-SecureString -String "clickshift-local" -Force -AsPlainText
    Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
    Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pwd -Force | Out-Null

    & $signtool.FullName sign /fd SHA256 /f $pfxPath /p "clickshift-local" "$stageDir\ClickShiftShell.dll"
    & $signtool.FullName sign /fd SHA256 /f $pfxPath /p "clickshift-local" "$stageDir\ClickShiftStub.exe"

    $msixPath = Join-Path $assets "ClickShiftShell.msix"
    if (Test-Path $msixPath) { Remove-Item $msixPath -Force }
    & $makeappx.FullName pack /d $stageDir /p $msixPath /o
    & $signtool.FullName sign /fd SHA256 /f $pfxPath /p "clickshift-local" $msixPath

    Remove-Item $pfxPath -Force -ErrorAction SilentlyContinue
    Remove-Item $stageDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Built and signed $msixPath"
} else {
    Write-Host "Windows SDK makeappx/signtool not found locally; using unpacked AppxManifest + ClickShiftShell.dll"
}
