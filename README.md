# Elmar's Lead Generation Quality Studio

Windows desktop app for checking lead-generation CSV files before upload to Reply.io.

It helps you:

- upload a CSV file
- detect null or blank values
- catch missing required columns
- validate email fields
- spot duplicate emails
- export a clean CSV when the data passes
- export a highlighted Excel review file when issues are found

## Installation

### Option 1: Install the app on Windows

If you already have the installer file:

1. Download the latest versioned installer, for example `ElmarsLeadGenerationQualityStudio-1.2.2-Setup.exe`.
2. Double-click the installer.
3. Follow the setup steps.
4. Launch `Elmar's Lead Generation Quality Studio` from the Start menu or desktop shortcut.

If Windows shows an unknown publisher warning, that is expected unless the installer is code-signed.

### Option 2: Use the standalone app folder

If you already have the standalone app build:

1. Download the `ElmarsLeadGenerationQualityStudio` folder.
2. Keep all files inside that folder together.
3. Run `ElmarsLeadGenerationQualityStudio.exe` from inside the folder.

This option does not require a full installer.

## Build From Source

If you are installing from this GitHub repository directly, build the app locally.

### Requirements

- Windows 10 or 11
- Python 3.11+ installed and available in `PATH`
- PowerShell

### Build the executable

1. Open PowerShell in the project folder.
2. Run:

```powershell
.\build_exe.ps1
```

After the build completes, the app will be created at:

```text
dist\ElmarsLeadGenerationQualityStudio\ElmarsLeadGenerationQualityStudio.exe
```

### Build the installer

1. Install Inno Setup 6.
2. Open PowerShell in the project folder.
3. Run:

```powershell
.\build_installer.ps1
```

After the build completes, the installer will be created at:

```text
installer-dist\ElmarsLeadGenerationQualityStudio-1.2.2-Setup.exe
```

## Run From Source

If you want to start the app without packaging it:

1. Install dependencies:

```powershell
pip install -r requirements.txt
```

2. Run the app:

```powershell
python check_nulls.py
```

## How To Use

1. Open the app.
2. Click `Upload File`.
3. Select your CSV file.
4. Click `Analyze`.
5. If the file is clean, save the clean export.
6. If the file has issues, review the popup and export the highlighted Excel file if needed.

## Project Files

- `check_nulls.py`: main desktop application
- `build_exe.ps1`: builds the standalone `.exe`
- `build_installer.ps1`: builds the Windows installer
- `installer.iss`: Inno Setup installer configuration
- `requirements.txt`: Python dependencies
- `INSTALL.md`: additional packaging notes

## Notes

- CSV files cannot store cell fill colors.
- Highlighted error cells are exported in Excel `.xlsx` format.
- This repository can also include the packaged Windows app files when they are intentionally committed for download.
