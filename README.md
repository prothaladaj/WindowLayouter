# WindowLayouter

Native Windows utility for discovering, filtering, and rearranging top-level windows across monitors.

## Repository layout

- `native/` - the actual Win32 application, assets, resource script, WSL build script, and Visual Studio files
- `installers/` - installer definitions for packaging the native app

## Build from WSL

Requires `mingw-w64`.

```bash
./build.sh
```

Or directly:

```bash
bash native/build-wsl.sh
```

Output:

- `native/bin/Release/WindowLayouter.Native.exe`

## Build on Windows

Open:

- `native/WindowLayouter.Native.sln`

Then build `Release | x64`.

## Installer

Use Inno Setup 6 on Windows and build:

- `installers/WindowLayouter.Native.iss`

This produces:

- `artifacts/installer/WindowLayouter.Native-Setup-x64.exe`

## Notes

- The app is `PerMonitorV2` DPI-aware.
- The main app lives in the tray and can be shown from the tray icon.
- Presets, hotkeys, layout gap, and window placement are stored in:
  - `%LOCALAPPDATA%\\WindowLayouter.Native\\settings.ini`
