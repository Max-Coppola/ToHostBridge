$ErrorActionPreference = "Stop"

$nugetUrl = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
$nugetFile = "nuget.exe"

if (-not (Test-Path $nugetFile)) {
    Write-Host "Downloading nuget.exe..."
    Invoke-WebRequest -Uri $nugetUrl -OutFile $nugetFile
}

$midiNupkgUrl = "https://github.com/microsoft/MIDI/releases/download/rc-3/Microsoft.Windows.Devices.Midi2.1.0.16-rc.3.7.nupkg"
$midiNupkgFile = "Microsoft.Windows.Devices.Midi2.1.0.16-rc.3.7.nupkg"

if (-not (Test-Path $midiNupkgFile)) {
    Write-Host "Downloading Windows MIDI Services NuGet package..."
    Invoke-WebRequest -Uri $midiNupkgUrl -OutFile $midiNupkgFile
}

Write-Host "Installing Microsoft.Windows.CppWinRT..."
.\nuget.exe install Microsoft.Windows.CppWinRT -Version 2.0.240405.15 -OutputDirectory packages

Write-Host "Extracting Microsoft.Windows.Devices.Midi2..."
.\nuget.exe install Microsoft.Windows.Devices.Midi2 -Source "$PWD" -OutputDirectory packages -Prerelease

Write-Host "Setup complete!"
