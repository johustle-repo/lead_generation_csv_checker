# Installable Build

This project can be packaged as a Windows desktop app.

Current packaged version target: `1.2.1`

## Option 1: Build an `.exe`

1. Open PowerShell in this folder.
2. Run:

```powershell
.\build_exe.ps1
```

After the build finishes, the app will be here:

```text
dist\ElmarsLeadGenerationQualityStudio\ElmarsLeadGenerationQualityStudio.exe
```

This executable includes:

- the premium desktop UI
- reviewed error export
- highlighted Excel error export (`.xlsx`) for files with issues

## Option 2: Build a Windows installer

1. Install Inno Setup 6 on the PC used for packaging.
2. Run:

```powershell
.\build_installer.ps1
```

After the build finishes, the installer will be here:

```text
installer-dist\Setup.exe
```

## Notes

- `build_exe.ps1` creates a local `.venv` if needed.
- The installer script depends on the executable from PyInstaller.
- If Windows warns about an unknown publisher, that is expected unless the installer is code-signed.
- The highlighted error export needs `openpyxl`, which is already included in `requirements.txt`.
