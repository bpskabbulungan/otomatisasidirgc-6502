# Packaging (Windows)

## Build EXE (PyInstaller)

Run:

```powershell
.\packaging\build_exe.ps1
```

Output will be created in `dist/filename.exe`.
This build script also downloads and embeds Chromium so end users don't need to install Playwright.

## Build Installer (Inno Setup)

1) Install Inno Setup.
2) Open `packaging/installer.iss`.
3) Build to generate the installer EXE in `packaging/dist-installer/`.

If you change the app name or version, update the constants at the top of the `.iss` file.
