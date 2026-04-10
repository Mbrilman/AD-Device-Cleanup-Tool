# AD Device Inventory, Cleanup & Action Tool

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-3.2.0-informational)

A PowerShell script that queries Active Directory for all computer objects, evaluates each against configurable inactivity thresholds, and produces a colour-coded Excel report — with an optional interactive cleanup menu to disable or delete stale devices directly from AD.

No Microsoft Office, NuGet, or internet connection required. Excel output is built with pure PowerShell via `System.IO.Compression`.

---

## Features

- Auto-detects the domain from the machine running the script
- Auto-installs the ActiveDirectory module (RSAT) if missing, when running as Administrator
- Non-administrator execution detected and handled gracefully — export-only mode with clear warnings
- Colour-coded Excel workbook: Summary sheet + Device sheet + Action Log sheet
- Three-tier configuration: CLI parameters > embedded config block > interactive prompts
- Threshold validation — warns if Delete threshold is not greater than Disable threshold
- AD Recycle Bin check and auto-enable (requires Enterprise Admin)
- Interactive action menu: disable, delete, or combined operations with WhatIf support
- PowerShell version and execution context shown in header, summary, and Excel report

---

## Requirements

| Requirement | Details |
|---|---|
| PowerShell | 5.1 or higher |
| ActiveDirectory module | Auto-installed if missing (admin required) |
| Permissions (export only) | Domain User |
| Permissions (disable/delete) | Domain Admin |
| Permissions (Recycle Bin enable) | Enterprise Admin |

---

## Usage

### Fully interactive — prompts for all settings at runtime
```powershell
.\Export-ADDevices-v3.2.0.ps1
```

### Pass thresholds via parameters — skips prompts for those values
```powershell
.\Export-ADDevices-v3.2.0.ps1 -DisableThreshold 60 -DeleteThreshold 180
```

### Enable the interactive AD cleanup menu
```powershell
.\Export-ADDevices-v3.2.0.ps1 -Interactive
```

### Dry-run — no changes written to Active Directory
```powershell
.\Export-ADDevices-v3.2.0.ps1 -Interactive -WhatIf
```

### Custom output path and report author
```powershell
.\Export-ADDevices-v3.2.0.ps1 -OutputPath "C:\Reports\ADCleanup.xlsx" -Author "John Doe / IT"
```

---

## Configuration Block

Open the script and locate the `CONFIGURATION BLOCK` near the top. Set values here to avoid being prompted on every run. Leave a value as `$null` to be prompted at runtime.

```powershell
$Config = @{
    DisableThreshold        = 60      # days inactive before DISABLE recommendation
    DeleteThreshold         = 180     # days inactive before DELETE recommendation
    DisabledDeleteThreshold = 365     # days since last AD change before deleting already-disabled devices
    OutputPath              = $null   # $null = auto-generate in script directory
    Author                  = $null   # $null = current Windows user (DOMAIN\user)
}
```

**Configuration priority (highest to lowest):**
1. CLI parameter (`-DisableThreshold 60`)
2. Config block value
3. Interactive prompt at runtime

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-DisableThreshold` | int | *(prompted)* | Days inactive before DISABLE recommendation |
| `-DeleteThreshold` | int | *(prompted)* | Days inactive before DELETE recommendation |
| `-DisabledDeleteThreshold` | int | *(prompted)* | Days since last AD change for already-disabled devices |
| `-OutputPath` | string | *(auto)* | Full path for the output .xlsx file |
| `-Author` | string | `DOMAIN\user` | Name written into the Excel report header |
| `-Interactive` | switch | off | Enables the AD action menu after export |
| `-WhatIf` | switch | off | Simulates actions without writing to AD |

---

## Excel Output

The workbook contains up to three sheets:

| Sheet | Contents |
|---|---|
| **Summary** | Domain, thresholds, Recycle Bin status, totals, legend |
| **\<domain\>** | All AD computer objects with colour-coded recommendations |
| **Action Log** | Timestamped log of disable/delete operations (Interactive mode only) |

### Colour legend

| Colour | Recommendation |
|---|---|
| Green | ACTIVE — no action required |
| Yellow | DISABLE — inactive between Disable and Delete threshold |
| Red/salmon | DELETE — inactive beyond Delete threshold |
| Orange | DISABLED — already disabled in AD, pending review |
| Light blue | MONITOR — recently created, no logon yet |

---

## Recommendation Logic

| Condition | Recommendation |
|---|---|
| Last activity < DisableThreshold | ACTIVE |
| Last activity between Disable and Delete threshold | DISABLE (inactive >Xd) |
| Last activity > DeleteThreshold | DELETE (inactive >Xd) |
| No logon, created < DisableThreshold days ago | MONITOR |
| No logon, created > DisableThreshold days ago | DISABLE (no logon, >Xd) |
| No logon, created > DeleteThreshold days ago | DELETE (no logon, >Xd) |
| Already disabled, AD changed < DisabledDeleteThreshold days ago | DISABLED - pending review |
| Already disabled, AD changed > DisabledDeleteThreshold days ago | DELETE (disabled >Xd) |

---

## Running Without Administrator Rights

The script handles non-admin execution gracefully:

- Displays a clear warning listing which features are unavailable
- Prompts to continue in limited export-only mode
- Automatically disables `-Interactive` (AD writes would fail)
- Skips Recycle Bin enable attempt with an explanatory message

Export and reporting work fully as a standard domain user, provided the ActiveDirectory module is already installed.

---

## Installing RSAT Manually

If the auto-install fails or you are running without admin rights:

**Windows Server:**
```powershell
Add-WindowsFeature -Name RSAT-AD-PowerShell
```

**Windows 10 / 11:**
```powershell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

Or: Settings > Optional Features > Add a feature > search "RSAT: Active Directory"

---

## Credits

| Role | Name |
|---|---|
| Author | [Macel Brilman](https://github.com/Mbrilman) — Senior Hybrid IT Specialist @ Enstall |
| AI Assistant | [Claude Sonnet](https://claude.ai) — Anthropic |

---

## License

MIT License — free to use, modify, and distribute. Attribution appreciated.
