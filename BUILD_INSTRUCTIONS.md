# Build Instructions: ToHost Bridge v1.1

To ensure the Windows MIDI Services SDK and the virtual port manifestation work correctly, this project **must** be built as an **x64** application using the **Visual Studio 18 2026** generator.

## Prerequisites
1.  **Windows MIDI Services SDK Runtime**: [Download here](https://aka.ms/MidiServicesLatestSdkRuntimeInstaller).
2.  **Visual Studio 2022/2026**: With "Desktop development with C++" workload.
3.  **CMake 3.20+**.

## Compilation Steps

### 1. Clean Previous Config (Optional but Recommended)
Delete the `build/` folder if switching architectures:
```powershell
Remove-Item -Path build -Recurse -Force
```

### 2. Configure for x64
Explicitly specify the 64-bit architecture:
```powershell
cmake -S . -B build -G "Visual Studio 18 2026" -A x64
```

### 3. Build Release
```powershell
cmake --build build --config Release
```

The resulting binary will be located at:
`build/Release/ToHostBridge.exe`

## Releases & Packaging
When you are ready to create a new official release for GitHub:
1.  **Preparation**: Ensure the `1.1.0` version and binary are stable in `build/Release/ToHostBridge.exe`.
2.  **Asset Sync**: (Optional) Copy the compiled binary `ToHostBridge.exe` from `build/Release/` to the root `releases/` directory for an easy-access copy.
3.  **Generating the MSI**: To generate the Windows Installer (`.msi`), use the local WiX Toolset located in the `packages/` folder.
    -   **Candle (Compile)**: 
        ```powershell
        .\packages\WiX.Toolset.2015.3.10.0.1503\tools\wix\candle.exe ToHostBridge.wxs -ext WixUIExtension -ext WixUtilExtension -o ToHostBridge.wixobj
        ```
    -   **Light (Link)**: 
        ```powershell
        .\packages\WiX.Toolset.2015.3.10.0.1503\tools\wix\light.exe ToHostBridge.wixobj -ext WixUIExtension -ext WixUtilExtension -o ToHostBridge.msi -b .
        ```
4.  **Finalize**: Move the newly generated `ToHostBridge.msi` to the `releases/` folder before committing to Git.

## Critical Developer Notes
- **COM Apartment**: `winrt::init_apartment()` MUST be called at the very top of `wWinMain`.
- **Threading**: Virtual ports must be created synchronously on the UI thread or carefully managed within a WinRT-compatible apartment.
- **Library Linking**: Ensure `windowsapp.lib` and `winmm.lib` are linked in `CMakeLists.txt`.
