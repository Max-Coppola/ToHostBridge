$ErrorActionPreference = "Stop"

$makeappx = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\makeappx.exe"
$signtool = "C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\signtool.exe"

$sourceExe = "build\Release\ToHostBridge.exe"
$manifest = "msix\Package.appxmanifest"
$assetsDir = "msix\Assets"
$outputMsix = "releases\ToHostBridge_v1.2.0_x64.msix"
$layoutDir = "msix\layout"

if (-not (Test-Path $sourceExe)) {
    Write-Error "Source binary not found at $sourceExe. Please build the Release configuration first."
}

Write-Host "--- Creating MSIX Layout ---"
if (Test-Path $layoutDir) { Remove-Item $layoutDir -Recurse -Force }
New-Item -ItemType Directory -Path $layoutDir
New-Item -ItemType Directory -Path "$layoutDir\Assets"

Copy-Item $sourceExe -Destination "$layoutDir\ToHostBridge.exe"
Copy-Item $manifest -Destination "$layoutDir\AppxManifest.xml"
Copy-Item "$assetsDir\*" -Destination "$layoutDir\Assets"

Write-Host "--- Packing MSIX ---"
if (Test-Path $outputMsix) { Remove-Item $outputMsix -Force }
& $makeappx pack /d $layoutDir /p $outputMsix

Write-Host "--- Signing MSIX (for Local Testing) ---"
$pfx = "msix\ToHostBridgeDev.pfx"
$password = "Pass1234"
& $signtool sign /fd SHA256 /a /f $pfx /p $password $outputMsix

Write-Host "Build Complete: $outputMsix"
Write-Host "NOTE: To install this on another machine, you MUST install the .pfx (or .cer) into that machine's 'Trusted People' certificate store first."
