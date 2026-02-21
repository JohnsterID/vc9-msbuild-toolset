# VC9 SP1 Installer

Installs Visual C++ 2008 SP1 (`PlatformToolset=v90`) for use with Visual Studio 2010-2022.

## Quick Start

```powershell
# Run as Administrator
.\install-vc9.ps1
```

## What It Does

1. Downloads **Windows SDK 7.0** and installs VC9 SP1 compiler (15.0.30729.1)
2. Downloads **VS2010 Express** and extracts the 2 MSBuild toolset files (8KB)
3. Installs toolset to legacy MSBuild path (works with all VS versions)

## Requirements

- Windows 10/11
- Visual Studio 2010 or later (any edition)
- PowerShell 5.1+
- [7-Zip](https://www.7-zip.org/) in PATH
- ~1.5 GB free disk space for downloads (can be cleaned up after)

## Pre-downloaded ISOs

If you already have the ISOs, place them next to the script to skip downloading:

```
install-vc9.ps1
GRMSDK_EN_DVD.iso      (567 MB - SDK 7.0)
VS2010Express1.iso     (694 MB - VS2010 Express)
```

The script checks for local files first, then cached downloads, then downloads fresh.

## What Gets Installed

| Component | Version | Source | Installed Size |
|-----------|---------|--------|----------------|
| VC9 x86 compiler | 15.0.30729.1 | SDK 7.0 | ~50 MB |
| VC9 x64 cross-compiler | 15.0.30729.1 | SDK 7.0 | ~50 MB |
| MSBuild v90 toolset | - | VS2010 | 8 KB |

## Installation Paths

```
C:\Program Files (x86)\Microsoft Visual Studio 9.0\
├── Common7\Tools\
│   └── vsvars32.bat         # Environment setup script
└── VC\
    ├── bin\cl.exe           # Compiler
    ├── include\             # Headers (with TR1)
    └── lib\                 # Libraries

C:\Program Files (x86)\MSBuild\Microsoft.Cpp\v4.0\Platforms\Win32\PlatformToolsets\v90\
├── Microsoft.Cpp.Win32.v90.props
└── Microsoft.Cpp.Win32.v90.targets
```

The legacy MSBuild path (`v4.0\Platforms`) is checked by all Visual Studio versions (2010-2022).

## Environment Variables

The script sets:

| Variable | Value |
|----------|-------|
| `VS90COMNTOOLS` | `C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\Tools\` |

Registry keys for MSBuild:
- `HKLM\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\9.0\Setup\VC\ProductDir`
- `HKLM\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\9.0\Setup\VS\ProductDir`

## Sources

Microsoft's original download links are no longer available. Files are downloaded from archive sources:

| File | Size | Source |
|------|------|--------|
| SDK 7.0 | 567 MB | [archive.org/grmsdkx-en-dvd](https://archive.org/details/grmsdkx-en-dvd) |
| VS2010 Express | 694 MB | [web.archive.org](https://web.archive.org/web/20140424044344/http://download.microsoft.com/download/1/E/5/1E5F1C0A-0D5B-426A-A603-1798B951DDAE/VS2010Express1.iso) |
| VS2008 SP1 | 900 MB | [web.archive.org](https://web.archive.org/web/20250618154620/https://download.microsoft.com/download/E/8/E/E8EEB394-7F42-4963-A2D8-29559B738298/VS2008ExpressWithSP1ENUX1504728.iso) |

All files are original Microsoft releases preserved in archives.

## Optional: MFC Feature Pack

```powershell
.\install-vc9.ps1 -IncludeMFC
```

Adds MFC Feature Pack (Office 2007-style ribbon UI, docking panes, visual managers).  
Source: VS2008 SP1 (~900 MB download, ~100 MB installed)

## Troubleshooting

### "Platform Toolset v90 cannot be found"
Re-run `install-vc9.ps1` as Administrator.

### "Cannot open include file: 'array'"
VC9 not installed correctly. Check that `cl.exe` exists:
```cmd
dir "C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\bin\cl.exe"
```

### 7-Zip not found
Install 7-Zip and ensure it's in your PATH:
```powershell
winget install 7zip.7zip
```

## See Also

- [Community-Patch-DLL](https://github.com/JohnsterID/Community-Patch-DLL) - Example project using v90
- [archaic-msvc/msvc900](https://github.com/archaic-msvc/msvc900) - Pre-extracted VC9 toolchain
