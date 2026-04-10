<#
.SYNOPSIS
    AD Device Inventory, Cleanup & Action Tool
    Version : 3.2.0

.DESCRIPTION
    Queries Active Directory for all computer objects in the current domain,
    evaluates each against configurable inactivity thresholds, and produces a
    colour-coded Excel (.xlsx) report with an optional interactive cleanup menu.

    Threshold values can be set in three ways (evaluated in this order):
      1. Command-line parameters  (-DisableThreshold, -DeleteThreshold, ...)
      2. The CONFIGURATION block  in this script file (set once, reuse always)
      3. Interactive prompts      at runtime (fallback when neither above is set)

    Features:
      - Domain auto-detected from the running machine
      - ActiveDirectory module auto-installed if missing (requires admin)
      - Non-administrator execution detected and handled gracefully
      - No Microsoft Office, NuGet, or internet connection required
      - Excel output built with pure PowerShell via System.IO.Compression

.NOTES
    Author       : Macel Brilman
    AI Assistant : Claude (Anthropic) - claude.ai
    AI Model     : Claude Sonnet 4.6
    License      : MIT
    Version      : 3.2.0
    Date         : 2026-04-10
    Requires     : PowerShell 5.1+
                   ActiveDirectory module (auto-installed when running as admin)
    Run As       : Domain user for export only
                   Domain Admin for disable/delete actions

.EXAMPLE
    # Fully interactive - prompts for all settings
    .\Export-ADDevices-v3.2.0.ps1

    # Parameters override config block and skip prompts
    .\Export-ADDevices-v3.2.0.ps1 -DisableThreshold 60 -DeleteThreshold 180

    # Interactive cleanup menu
    .\Export-ADDevices-v3.2.0.ps1 -Interactive

    # Dry-run, no AD writes
    .\Export-ADDevices-v3.2.0.ps1 -Interactive -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [int]    $DisableThreshold        = -1,
    [int]    $DeleteThreshold         = -1,
    [int]    $DisabledDeleteThreshold = -1,
    [string] $OutputPath              = "",
    [string] $Author                  = "",
    [switch] $Interactive
)

$ErrorActionPreference = "Stop"

$ScriptVersion = "3.2.0"
$ScriptDate    = "2026-04-10"
$PSVer         = "$($PSVersionTable.PSVersion.Major).$($PSVersionTable.PSVersion.Minor)"

# =============================================================================
#  CONFIGURATION BLOCK
#  Set values here to avoid being prompted every run.
#  A value of $null means "ask the user at runtime".
#  Command-line parameters always take precedence over these settings.
# =============================================================================
$Config = @{
    # How many days inactive before a device is recommended for DISABLE
    # Suggested: 60
    DisableThreshold        = $null

    # How many days inactive before a device is recommended for DELETE
    # Suggested: 180
    DeleteThreshold         = $null

    # How many days since last AD change before an already-DISABLED device
    # is recommended for DELETE
    # Suggested: 365
    DisabledDeleteThreshold = $null

    # Full path for the output .xlsx file
    # Leave $null to auto-generate in the script directory
    # Example: "C:\Reports\AD_Export.xlsx"
    OutputPath              = $null

    # Name written into the Excel report header
    # Leave $null to use the current Windows user (DOMAIN\user)
    # Example: "John Doe / IT Department"
    Author                  = $null
}
# =============================================================================
#  END CONFIGURATION BLOCK
# =============================================================================

# =============================================================================
#  ELEVATION CHECK
# =============================================================================
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
              [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Host ""
    Write-Host "  +-------------------------------------------------------------+" -ForegroundColor Yellow
    Write-Host "  |  WARNING: Running WITHOUT administrator privileges           |" -ForegroundColor Yellow
    Write-Host "  |                                                              |" -ForegroundColor Yellow
    Write-Host "  |  The following features will NOT be available:               |" -ForegroundColor Yellow
    Write-Host "  |    - ActiveDirectory module auto-install                     |" -ForegroundColor Yellow
    Write-Host "  |    - AD Recycle Bin enable                                   |" -ForegroundColor Yellow
    Write-Host "  |    - Device disable / delete actions                         |" -ForegroundColor Yellow
    Write-Host "  |                                                              |" -ForegroundColor Yellow
    Write-Host "  |  Export-only mode will work if the AD module is installed.   |" -ForegroundColor Yellow
    Write-Host "  |                                                              |" -ForegroundColor Yellow
    Write-Host "  |  Recommendation: Right-click PowerShell > Run as Admin       |" -ForegroundColor Yellow
    Write-Host "  +-------------------------------------------------------------+" -ForegroundColor Yellow
    Write-Host ""
    $Continue = Read-Host "  Continue in limited export-only mode? (Y/N)"
    if ($Continue -notmatch "^[Yy]") {
        Write-Host "  Aborted. Restart as Administrator for full functionality." -ForegroundColor DarkGray
        exit 0
    }
    Write-Host ""
    if ($Interactive) {
        Write-Host "  [INFO] -Interactive disabled. AD changes require Administrator." -ForegroundColor DarkYellow
        Write-Host ""
        $Interactive = $false
    }
}

# =============================================================================
#  ACTIVEDIRECTORY MODULE BOOTSTRAP
# =============================================================================
function Install-ADModule {
    if (-not $IsAdmin) {
        Write-Host "  [SKIP] Cannot auto-install module without Administrator rights." -ForegroundColor DarkYellow
        Write-Host "         Restart as Administrator and re-run, or install manually:" -ForegroundColor DarkYellow
        Write-Host "           Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor DarkYellow
        exit 1
    }
    $OS = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
    if ($OS -match "Server") {
        Write-Host "  Detected: Windows Server. Installing via Add-WindowsFeature..." -ForegroundColor Yellow
        try {
            Import-Module ServerManager -ErrorAction Stop
            $r = Add-WindowsFeature -Name RSAT-AD-PowerShell -ErrorAction Stop
            if (-not $r.Success) { throw "Add-WindowsFeature reported failure." }
            Write-Host "  RSAT-AD-PowerShell installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "  AUTO-INSTALL FAILED. Run manually:" -ForegroundColor Red
            Write-Host "    Add-WindowsFeature -Name RSAT-AD-PowerShell" -ForegroundColor Yellow
            exit 1
        }
    } else {
        Write-Host "  Detected: Windows Client. Installing via Add-WindowsCapability..." -ForegroundColor Yellow
        try {
            Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop | Out-Null
            Write-Host "  RSAT ActiveDirectory tools installed successfully." -ForegroundColor Green
        } catch {
            Write-Host "  AUTO-INSTALL FAILED. Run manually as Administrator:" -ForegroundColor Red
            Write-Host "    Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor Yellow
            Write-Host "  Or: Settings > Optional Features > Add a feature > RSAT: Active Directory" -ForegroundColor Yellow
            exit 1
        }
    }
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host ""
    Write-Host "  [WARN] ActiveDirectory module not found. Attempting install..." -ForegroundColor Yellow
    Install-ADModule
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    Write-Host "  [ERROR] Failed to import ActiveDirectory module after install attempt." -ForegroundColor Red
    Write-Host "  Restart PowerShell and re-run." -ForegroundColor Yellow
    exit 1
}

# =============================================================================
#  AUTO-DETECT DOMAIN
# =============================================================================
try {
    $Domain = (Get-ADDomain -ErrorAction Stop).DNSRoot
} catch {
    Write-Error "Cannot detect domain. Ensure this machine is domain-joined."
    exit 1
}

Clear-Host
Write-Host ""
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "|       AD Device Inventory, Cleanup & Action Tool            |" -ForegroundColor Cyan
Write-Host ("|  Domain  : {0,-49}|" -f $Domain)                              -ForegroundColor Cyan
Write-Host ("|  Version : {0,-49}|" -f "$ScriptVersion  |  $ScriptDate")     -ForegroundColor Cyan
Write-Host ("|  Runtime : {0,-49}|" -f ("PowerShell $PSVer  |  " + $(if ($IsAdmin) { "Administrator" } else { "Standard User (limited)" }))) -ForegroundColor $(if ($IsAdmin) { "Cyan" } else { "Yellow" })
Write-Host ("|  Author  : {0,-49}|" -f "Macel Brilman  |  Claude Sonnet (Anthropic)")  -ForegroundColor Cyan
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
if ($WhatIfPreference) {
    Write-Host "  *** WHATIF MODE - no changes written to Active Directory ***" -ForegroundColor Magenta
    Write-Host ""
}

# =============================================================================
#  RESOLVE CONFIGURATION
#  Priority: CLI parameter > Config block > Interactive prompt
# =============================================================================

function Read-IntPrompt {
    param([string]$Label, [string]$Suggestion)
    while ($true) {
        $raw = Read-Host ("  {0} (suggested: {1})" -f $Label, $Suggestion)
        $v   = 0
        if ([int]::TryParse($raw.Trim(), [ref]$v) -and $v -gt 0) { return $v }
        Write-Host "  Enter a positive integer." -ForegroundColor DarkYellow
    }
}

function Read-StringPrompt {
    param([string]$Label, [string]$Default)
    $raw = Read-Host ("  {0} [Enter to use: {1}]" -f $Label, $Default)
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    return $raw.Trim()
}

Write-Host "[0/4] Configuration" -ForegroundColor Yellow
Write-Host ""

# DisableThreshold
if ($DisableThreshold -gt 0) {
    $FinalDisable = $DisableThreshold
    Write-Host ("  DisableThreshold        : {0} days  [parameter]" -f $FinalDisable) -ForegroundColor DarkCyan
} elseif ($null -ne $Config.DisableThreshold) {
    $FinalDisable = [int]$Config.DisableThreshold
    Write-Host ("  DisableThreshold        : {0} days  [config block]" -f $FinalDisable) -ForegroundColor DarkCyan
} else {
    Write-Host "  DisableThreshold - Days inactive before a device is recommended for DISABLE." -ForegroundColor Gray
    $FinalDisable = Read-IntPrompt "  Disable threshold (days)" "60"
}

# DeleteThreshold
if ($DeleteThreshold -gt 0) {
    $FinalDelete = $DeleteThreshold
    Write-Host ("  DeleteThreshold         : {0} days  [parameter]" -f $FinalDelete) -ForegroundColor DarkCyan
} elseif ($null -ne $Config.DeleteThreshold) {
    $FinalDelete = [int]$Config.DeleteThreshold
    Write-Host ("  DeleteThreshold         : {0} days  [config block]" -f $FinalDelete) -ForegroundColor DarkCyan
} else {
    Write-Host "  DeleteThreshold  - Days inactive before a device is recommended for DELETE." -ForegroundColor Gray
    $FinalDelete = Read-IntPrompt "  Delete threshold (days)" "180"
}

# DisabledDeleteThreshold
if ($DisabledDeleteThreshold -gt 0) {
    $FinalDisDel = $DisabledDeleteThreshold
    Write-Host ("  DisabledDeleteThreshold : {0} days  [parameter]" -f $FinalDisDel) -ForegroundColor DarkCyan
} elseif ($null -ne $Config.DisabledDeleteThreshold) {
    $FinalDisDel = [int]$Config.DisabledDeleteThreshold
    Write-Host ("  DisabledDeleteThreshold : {0} days  [config block]" -f $FinalDisDel) -ForegroundColor DarkCyan
} else {
    Write-Host "  DisabledDeleteThreshold  - Days since last AD change before an already-disabled device is deleted." -ForegroundColor Gray
    $FinalDisDel = Read-IntPrompt "  Disabled-delete threshold (days)" "365"
}

# Author
if (-not [string]::IsNullOrWhiteSpace($Author)) {
    $FinalAuthor = $Author
    Write-Host ("  Author                  : {0}  [parameter]" -f $FinalAuthor) -ForegroundColor DarkCyan
} elseif (-not [string]::IsNullOrWhiteSpace($Config.Author)) {
    $FinalAuthor = $Config.Author
    Write-Host ("  Author                  : {0}  [config block]" -f $FinalAuthor) -ForegroundColor DarkCyan
} else {
    $DefaultAuthor = "$($env:USERDOMAIN)\$($env:USERNAME)"
    $FinalAuthor   = Read-StringPrompt "Author name for report" $DefaultAuthor
}

# OutputPath
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $FinalOutput = $OutputPath
    Write-Host ("  OutputPath              : {0}  [parameter]" -f $FinalOutput) -ForegroundColor DarkCyan
} elseif (-not [string]::IsNullOrWhiteSpace($Config.OutputPath)) {
    $FinalOutput = $Config.OutputPath
    Write-Host ("  OutputPath              : {0}  [config block]" -f $FinalOutput) -ForegroundColor DarkCyan
} else {
    $TS          = Get-Date -Format "yyyyMMdd_HHmmss"
    $SafeDom     = $Domain.Replace(".", "-")
    $FinalOutput = Join-Path $PSScriptRoot ("AD_Device_Export_v{0}_{1}_{2}.xlsx" -f $ScriptVersion, $SafeDom, $TS)
    Write-Host ("  OutputPath              : {0}  [auto]" -f $FinalOutput) -ForegroundColor DarkCyan
}

Write-Host ""
Write-Host ("  Thresholds applied: DISABLE > {0}d  |  DELETE > {1}d  |  DISABLED-DELETE > {2}d" -f $FinalDisable, $FinalDelete, $FinalDisDel) -ForegroundColor White
Write-Host ""

# Validate: Delete must be greater than Disable
if ($FinalDelete -le $FinalDisable) {
    Write-Host "  [WARN] DeleteThreshold ($FinalDelete) is not greater than DisableThreshold ($FinalDisable)." -ForegroundColor DarkYellow
    Write-Host "         This means devices will jump straight to DELETE without a DISABLE phase." -ForegroundColor DarkYellow
    $Confirm = Read-Host "  Continue anyway? (Y/N)"
    if ($Confirm -notmatch "^[Yy]") { Write-Host "  Aborted." -ForegroundColor DarkGray; exit 0 }
    Write-Host ""
}

# =============================================================================
#  XLSX WRITER - Pure PowerShell, no Office/NuGet required
# =============================================================================
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$STYLE_NORMAL     = 0
$STYLE_HEADER     = 1
$STYLE_ACTIVE     = 2
$STYLE_DELETE     = 3
$STYLE_DISABLE    = 4
$STYLE_DISABLED   = 5
$STYLE_MONITOR    = 6
$STYLE_STATUS_DIS = 7
$STYLE_BOLD       = 8
$STYLE_BOLD_GREEN = 9
$STYLE_BOLD_RED   = 10

function New-Workbook {
    return @{
        SharedStr    = [System.Collections.Generic.List[string]]::new()
        SharedStrIdx = [System.Collections.Generic.Dictionary[string,int]]::new()
    }
}

function Get-SSIndex {
    param($WB, [string]$Val)
    $idx = 0
    if ($WB.SharedStrIdx.TryGetValue($Val, [ref]$idx)) { return $idx }
    $idx = $WB.SharedStr.Count
    $WB.SharedStr.Add($Val)
    $WB.SharedStrIdx[$Val] = $idx
    return $idx
}

function ConvertTo-XmlSafe {
    param([string]$s)
    $s = $s -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    $s = $s -replace '"', '&quot;'
    return $s
}

function Get-ColLetter {
    param([int]$n)
    $r = ""
    while ($n -gt 0) { $n--; $r = [char](65 + ($n % 26)) + $r; $n = [Math]::Floor($n / 26) }
    return $r
}

function Build-StylesXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="5">
  <font><sz val="10"/><name val="Calibri"/></font>
  <font><sz val="10"/><b/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  <font><sz val="10"/><b/><name val="Calibri"/></font>
  <font><sz val="10"/><b/><color rgb="FF8B0000"/><name val="Calibri"/></font>
  <font><sz val="10"/><b/><color rgb="FF006400"/><name val="Calibri"/></font>
</fonts>
<fills count="9">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1E2761"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF90EE90"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFA8072"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFF0"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFA500"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFADD8E6"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFAFAFA"/></patternFill></fill>
</fills>
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color auto="1"/></left>
    <right style="thin"><color auto="1"/></right>
    <top style="thin"><color auto="1"/></top>
    <bottom style="thin"><color auto="1"/></bottom>
    <diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1"/>
  <xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1"/>
  <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1"/>
  <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1"/>
  <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1"/>
  <xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1"/>
  <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1"/>
</cellXfs>
</styleSheet>
'@
}

function Build-SharedStrings {
    param($WB)
    $n   = $WB.SharedStr.Count
    $xml = [System.Text.StringBuilder]::new()
    [void]$xml.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$xml.AppendLine('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + $n + '" uniqueCount="' + $n + '">')
    foreach ($s in $WB.SharedStr) {
        [void]$xml.AppendLine('<si><t xml:space="preserve">' + (ConvertTo-XmlSafe $s) + '</t></si>')
    }
    [void]$xml.AppendLine('</sst>')
    return $xml.ToString()
}

function Build-DeviceSheet {
    param($WB, [object[]]$Rows, [string[]]$Cols, [int]$RecoCol, [int]$StatusCol)
    $sb = [System.Text.StringBuilder]::new()
    $nC = $Cols.Count
    $cw = @{}
    foreach ($c in $Cols) { $cw[$c] = [Math]::Max($c.Length + 2, 10) }
    foreach ($row in $Rows) {
        foreach ($c in $Cols) {
            $v = [string]$row.$c
            if (($v.Length + 2) -gt $cw[$c]) { $cw[$c] = [Math]::Min($v.Length + 2, 55) }
        }
    }
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.AppendLine('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.AppendLine('<cols>')
    for ($c = 1; $c -le $nC; $c++) {
        [void]$sb.AppendLine('<col min="' + $c + '" max="' + $c + '" width="' + $cw[$Cols[$c-1]] + '" customWidth="1"/>')
    }
    [void]$sb.AppendLine('</cols>')
    [void]$sb.AppendLine('<sheetData>')
    [void]$sb.AppendLine('<row r="1">')
    for ($c = 1; $c -le $nC; $c++) {
        $a = (Get-ColLetter $c) + '1'
        $i = Get-SSIndex $WB $Cols[$c - 1]
        [void]$sb.AppendLine('<c r="' + $a + '" t="s" s="' + $STYLE_HEADER + '"><v>' + $i + '</v></c>')
    }
    [void]$sb.AppendLine('</row>')
    $r = 2
    foreach ($row in $Rows) {
        [void]$sb.AppendLine('<row r="' + $r + '">')
        $rv = if ($RecoCol   -gt 0) { [string]$row.($Cols[$RecoCol   - 1]) } else { "" }
        $sv = if ($StatusCol -gt 0) { [string]$row.($Cols[$StatusCol - 1]) } else { "" }
        for ($c = 1; $c -le $nC; $c++) {
            $a   = (Get-ColLetter $c) + $r
            $val = [string]$row.($Cols[$c - 1])
            $s   = $STYLE_NORMAL
            if ($c -eq $RecoCol) {
                $s = if     ($rv -eq "ACTIVE")      { $STYLE_ACTIVE   }
                     elseif ($rv -like "DELETE*")    { $STYLE_DELETE   }
                     elseif ($rv -like "DISABLE (*") { $STYLE_DISABLE  }
                     elseif ($rv -like "DISABLED*")  { $STYLE_DISABLED }
                     elseif ($rv -like "MONITOR*")   { $STYLE_MONITOR  }
                     else                            { $STYLE_NORMAL   }
            } elseif ($c -eq $StatusCol -and $sv -eq "Disabled") {
                $s = $STYLE_STATUS_DIS
            }
            $num = 0.0
            if ([double]::TryParse($val, [System.Globalization.NumberStyles]::Any,
                                   [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
                [void]$sb.AppendLine('<c r="' + $a + '" s="' + $s + '"><v>' + $num + '</v></c>')
            } else {
                $i = Get-SSIndex $WB $val
                [void]$sb.AppendLine('<c r="' + $a + '" t="s" s="' + $s + '"><v>' + $i + '</v></c>')
            }
        }
        [void]$sb.AppendLine('</row>')
        $r++
    }
    [void]$sb.AppendLine('</sheetData>')
    [void]$sb.AppendLine('<autoFilter ref="A1:' + (Get-ColLetter $nC) + '1"/>')
    [void]$sb.AppendLine('</worksheet>')
    return $sb.ToString()
}

function Build-SummarySheet {
    param($WB, [object[]]$Rows)
    $sb   = [System.Text.StringBuilder]::new()
    $bold = @("AD Device","Domain","Generated","Report Date","Script Version",
              "POLICY","AD RECYCLE","TOTALS","LEGEND")
    [void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.AppendLine('<cols><col min="1" max="1" width="48" customWidth="1"/><col min="2" max="2" width="55" customWidth="1"/></cols>')
    [void]$sb.AppendLine('<sheetData>')
    $r = 1
    foreach ($row in $Rows) {
        $k = [string]$row.Item
        $v = [string]$row.Value
        [void]$sb.AppendLine('<row r="' + $r + '">')
        $ks = $STYLE_NORMAL
        foreach ($b in $bold) { if ($k -like "*$b*") { $ks = $STYLE_BOLD; break } }
        $vs = if ($v -eq "ENABLED") { $STYLE_BOLD_GREEN } elseif ($v -like "*RISK*") { $STYLE_BOLD_RED } else { $STYLE_NORMAL }
        $ki = Get-SSIndex $WB $k
        $vi = Get-SSIndex $WB $v
        [void]$sb.AppendLine('<c r="A' + $r + '" t="s" s="' + $ks + '"><v>' + $ki + '</v></c>')
        [void]$sb.AppendLine('<c r="B' + $r + '" t="s" s="' + $vs + '"><v>' + $vi + '</v></c>')
        [void]$sb.AppendLine('</row>')
        $r++
    }
    [void]$sb.AppendLine('</sheetData></worksheet>')
    return $sb.ToString()
}

function Build-ActionSheet {
    param($WB, [object[]]$Rows)
    $cols = @($Rows[0].PSObject.Properties.Name)
    return Build-DeviceSheet $WB $Rows $cols 0 0
}

function Save-Workbook {
    param(
        $WB,
        [string]   $Path,
        [object[]] $DeviceData,
        [string[]] $DisplayCols,
        [int]      $RecoCol,
        [int]      $StatusCol,
        [string]   $DevSheetName,
        [object[]] $SummaryRows,
        [object[]] $ActionLog
    )

    $hasAction = ($null -ne $ActionLog -and $ActionLog.Count -gt 0)

    foreach ($row in $DeviceData) {
        foreach ($c in $DisplayCols) { Get-SSIndex $WB ([string]$row.$c) | Out-Null }
    }
    foreach ($row in $SummaryRows) {
        Get-SSIndex $WB ([string]$row.Item)  | Out-Null
        Get-SSIndex $WB ([string]$row.Value) | Out-Null
    }
    if ($hasAction) {
        foreach ($row in $ActionLog) {
            foreach ($c in @($row.PSObject.Properties.Name)) { Get-SSIndex $WB ([string]$row.$c) | Out-Null }
        }
    }

    $sheetDevXml = Build-DeviceSheet  $WB $DeviceData $DisplayCols $RecoCol $StatusCol
    $sheetSumXml = Build-SummarySheet $WB $SummaryRows
    $stylesXml   = Build-StylesXml
    $sharedXml   = Build-SharedStrings $WB
    $sheetActXml = ""
    if ($hasAction) { $sheetActXml = Build-ActionSheet $WB $ActionLog }

    $safeName = ConvertTo-XmlSafe $DevSheetName

    $sheetEntries = '<sheet name="Summary" sheetId="1" r:id="rId1"/>' + "`n" +
                    '<sheet name="' + $safeName + '" sheetId="2" r:id="rId2"/>'
    $relEntries   = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' + "`n" +
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    $ctEntries    = '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' + "`n" +
                    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'

    if ($hasAction) {
        $sheetEntries += "`n" + '<sheet name="Action Log" sheetId="3" r:id="rId3"/>'
        $relEntries   += "`n" + '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
        $ctEntries    += "`n" + '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    }

    $workbookXml  = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + "`n" +
                    '<sheets>' + $sheetEntries + '</sheets></workbook>'

    $wbRels       = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "`n" +
                    $relEntries + "`n" +
                    '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' + "`n" +
                    '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' + "`n" +
                    '</Relationships>'

    $rootRels     = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "`n" +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' + "`n" +
                    '</Relationships>'

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + "`n" +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + "`n" +
                    '<Default Extension="xml"  ContentType="application/xml"/>' + "`n" +
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + "`n" +
                    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' + "`n" +
                    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + "`n" +
                    $ctEntries + "`n" +
                    '</Types>'

    if (Test-Path $Path) { Remove-Item $Path -Force }

    $enc = [System.Text.Encoding]::UTF8
    $ms  = [System.IO.MemoryStream]::new()
    $zip = [System.IO.Compression.ZipArchive]::new($ms, [System.IO.Compression.ZipArchiveMode]::Create, $true)

    $files = [ordered]@{
        "[Content_Types].xml"        = $contentTypes
        "_rels/.rels"                = $rootRels
        "xl/workbook.xml"            = $workbookXml
        "xl/_rels/workbook.xml.rels" = $wbRels
        "xl/styles.xml"              = $stylesXml
        "xl/sharedStrings.xml"       = $sharedXml
        "xl/worksheets/sheet1.xml"   = $sheetSumXml
        "xl/worksheets/sheet2.xml"   = $sheetDevXml
    }
    if ($hasAction) { $files["xl/worksheets/sheet3.xml"] = $sheetActXml }

    foreach ($kv in $files.GetEnumerator()) {
        $entry  = $zip.CreateEntry($kv.Key)
        $bytes  = $enc.GetBytes($kv.Value)
        $stream = $entry.Open()
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Close()
    }
    $zip.Dispose()
    [System.IO.File]::WriteAllBytes($Path, $ms.ToArray())
    $ms.Dispose()
}

# =============================================================================
#  AD HELPERS
# =============================================================================
function Test-RecycleBin {
    param([string]$Dom)
    try {
        $f = Get-ADOptionalFeature -Filter { Name -eq "Recycle Bin Feature" } -Server $Dom -ErrorAction Stop
        return ($null -ne $f -and $f.EnabledScopes.Count -gt 0)
    } catch { return $false }
}

function Enable-RecycleBin {
    param([string]$Dom)
    try {
        $Forest = (Get-ADDomain -Server $Dom -ErrorAction Stop).Forest
        Enable-ADOptionalFeature -Identity "Recycle Bin Feature" -Scope ForestOrConfigurationSet `
            -Target $Forest -Server $Dom -Confirm:$false -ErrorAction Stop
        return $true
    } catch {
        Write-Host ("  ERROR enabling Recycle Bin: {0}" -f $_) -ForegroundColor Red
        return $false
    }
}

function Get-DomainDevices {
    param([string]$Dom, [int]$DisThr, [int]$DelThr, [int]$DisDel)
    $Props = @(
        "Name","DNSHostName","OperatingSystem","OperatingSystemVersion",
        "whenCreated","whenChanged","LastLogonDate","PasswordLastSet",
        "Enabled","DistinguishedName","Description","IPv4Address","ManagedBy"
    )
    $Computers = Get-ADComputer -Filter * -Server $Dom -Properties $Props -ErrorAction Stop
    $Now  = Get-Date
    $List = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($C in $Computers) {
        $Signals      = @(@($C.LastLogonDate, $C.PasswordLastSet) | Where-Object { $_ })
        $LastActivity = if ($Signals.Count -gt 0) { @($Signals | Sort-Object -Descending)[0] } else { $null }
        $DaysInactive = if ($LastActivity) { [int]($Now - $LastActivity).TotalDays } else { [int]($Now - $C.whenCreated).TotalDays }
        $DaysCreated  = [int]($Now - $C.whenCreated).TotalDays
        $DaysChanged  = if ($C.whenChanged) { [int]($Now - $C.whenChanged).TotalDays } else { $DaysCreated }
        $CurStatus    = if ($C.Enabled -eq $true) { "Enabled" } elseif ($C.Enabled -eq $false) { "Disabled" } else { "Unknown" }

        $Rec = if ($C.Enabled -eq $false) {
                   if ($DaysChanged -gt $DisDel) { "DELETE (disabled >{0}d)" -f $DisDel }
                   else                          { "DISABLED - pending review" }
               } elseif (-not $LastActivity) {
                   if   ($DaysCreated -gt $DelThr)  { "DELETE (no logon, >{0}d)"  -f $DelThr  }
                   elseif ($DaysCreated -gt $DisThr) { "DISABLE (no logon, >{0}d)" -f $DisThr }
                   else                              { "MONITOR - new/not yet active" }
               } elseif ($DaysInactive -gt $DelThr) { "DELETE (inactive >{0}d)"  -f $DelThr  }
               elseif  ($DaysInactive -gt $DisThr)  { "DISABLE (inactive >{0}d)" -f $DisThr  }
               else                                  { "ACTIVE" }

        $OU = $C.DistinguishedName -replace "^CN=[^,]+,", ""
        $List.Add([PSCustomObject]@{
            "Domain"                = $Dom
            "Device Name"           = $C.Name
            "DNS Host Name"         = $C.DNSHostName
            "Current Status"        = $CurStatus
            "Operating System"      = $C.OperatingSystem
            "OS Version"            = $C.OperatingSystemVersion
            "Description"           = $C.Description
            "Created On"            = if ($C.whenCreated)     { $C.whenCreated.ToString("yyyy-MM-dd HH:mm")     } else { "" }
            "Last Modified (AD)"    = if ($C.whenChanged)     { $C.whenChanged.ToString("yyyy-MM-dd HH:mm")     } else { "" }
            "Last Logon"            = if ($C.LastLogonDate)   { $C.LastLogonDate.ToString("yyyy-MM-dd HH:mm")   } else { "Never / Unknown" }
            "Password Last Set"     = if ($C.PasswordLastSet) { $C.PasswordLastSet.ToString("yyyy-MM-dd HH:mm") } else { "Never / Unknown" }
            "Last Activity (best)"  = if ($LastActivity)      { $LastActivity.ToString("yyyy-MM-dd HH:mm")      } else { "Never / Unknown" }
            "Days Inactive"         = $DaysInactive
            "Days Since Created"    = $DaysCreated
            "Days Since AD Changed" = $DaysChanged
            "IPv4 Address"          = $C.IPv4Address
            "Managed By"            = $C.ManagedBy
            "Organisational Unit"   = $OU
            "Recommendation"        = $Rec
            "_DN"                   = $C.DistinguishedName
        })
    }

    return @($List | Sort-Object {
        switch -Wildcard ($_.Recommendation) {
            "DELETE*"    { 0 }
            "DISABLE (*" { 1 }
            "DISABLED*"  { 2 }
            "MONITOR*"   { 3 }
            default      { 4 }
        }
    }, "Days Inactive" -Descending)
}

# =============================================================================
#  STEP 1 - AD RECYCLE BIN
# =============================================================================
Write-Host "[1/4] Checking AD Recycle Bin status..." -ForegroundColor Yellow
Write-Host ""

$RBEnabled = Test-RecycleBin -Dom $Domain
if ($RBEnabled) {
    Write-Host ("  [OK]   {0,-30}  Recycle Bin: ENABLED" -f $Domain) -ForegroundColor Green
} else {
    Write-Host ("  [OFF]  {0,-30}  Recycle Bin: NOT ENABLED" -f $Domain) -ForegroundColor Yellow
    if ($WhatIfPreference) {
        Write-Host "         [WHATIF] Would attempt to enable - no change applied." -ForegroundColor Magenta
    } elseif (-not $IsAdmin) {
        Write-Host "         [SKIP]   Administrator rights required to enable Recycle Bin." -ForegroundColor DarkYellow
    } else {
        Write-Host "         Attempting to enable..." -ForegroundColor Yellow
        $OK = Enable-RecycleBin -Dom $Domain
        if ($OK) {
            Start-Sleep -Seconds 2
            $RBEnabled = Test-RecycleBin -Dom $Domain
            if ($RBEnabled) {
                Write-Host ("  [OK]   {0,-30}  Recycle Bin: ENABLED successfully" -f $Domain) -ForegroundColor Green
            } else {
                Write-Host ("  [WARN] {0,-30}  Enable ran but status unconfirmed - verify manually" -f $Domain) -ForegroundColor DarkYellow
            }
        } else {
            Write-Host ("  [FAIL] {0,-30}  Could not enable - Enterprise Admin rights required" -f $Domain) -ForegroundColor Red
            $Cont = Read-Host "  Continue anyway? Deletions will be IRREVERSIBLE. (Y/N)"
            if ($Cont -notmatch "^[Yy]") { Write-Host "Aborted." -ForegroundColor Yellow; exit 0 }
        }
    }
}
Write-Host ""

# =============================================================================
#  STEP 2 - QUERY
# =============================================================================
Write-Host "[2/4] Querying Active Directory: $Domain" -ForegroundColor Yellow
Write-Host ""

Write-Host ("  Querying: {0,-32}" -f $Domain) -ForegroundColor DarkCyan -NoNewline
try {
    $DeviceData = @(Get-DomainDevices -Dom $Domain -DisThr $FinalDisable -DelThr $FinalDelete -DisDel $FinalDisDel)
    $TotalCount = $DeviceData.Count
    Write-Host ("{0,5} objects found" -f $TotalCount) -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Error "AD query failed: $_"
    exit 1
}
Write-Host ""

# =============================================================================
#  STEP 3 - EXCEL
# =============================================================================
Write-Host "[3/4] Building Excel workbook..." -ForegroundColor Yellow

$DevSheetName = $Domain.Replace(".", "-")
if ($DevSheetName.Length -gt 31) { $DevSheetName = $DevSheetName.Substring(0, 31) }

$DisplayCols  = @($DeviceData[0].PSObject.Properties.Name | Where-Object { $_ -ne "_DN" })
$RecoColIdx   = [Array]::IndexOf($DisplayCols, "Recommendation") + 1
$StatusColIdx = [Array]::IndexOf($DisplayCols, "Current Status")  + 1

$nA  = @($DeviceData | Where-Object { $_.Recommendation -eq "ACTIVE"       }).Count
$nDi = @($DeviceData | Where-Object { $_.Recommendation -like "DISABLE (*" }).Count
$nDe = @($DeviceData | Where-Object { $_.Recommendation -like "DELETE*"    }).Count
$nDs = @($DeviceData | Where-Object { $_.Recommendation -like "DISABLED*"  }).Count
$nMo = @($DeviceData | Where-Object { $_.Recommendation -like "MONITOR*"   }).Count

$RBStatus  = if ($RBEnabled) { "ENABLED" } else { "NOT ENABLED - RISK!" }
$RunAsText = if ($IsAdmin) { "Administrator" } else { "Standard User (limited)" }

$SummaryRows = @(
    [PSCustomObject]@{ Item = "AD Device Cleanup Report";             Value = "" },
    [PSCustomObject]@{ Item = "Domain";                               Value = $Domain },
    [PSCustomObject]@{ Item = "Generated By";                         Value = $FinalAuthor },
    [PSCustomObject]@{ Item = "Report Date";                          Value = (Get-Date -Format "yyyy-MM-dd HH:mm") },
    [PSCustomObject]@{ Item = "Script Version";                       Value = $ScriptVersion },
    [PSCustomObject]@{ Item = "PowerShell Version";                   Value = "PowerShell $PSVer" },
    [PSCustomObject]@{ Item = "Execution Context";                    Value = $RunAsText },
    [PSCustomObject]@{ Item = "AI Assistant";                         Value = "Claude Sonnet (Anthropic) - claude.ai" },
    [PSCustomObject]@{ Item = "";                                     Value = "" },
    [PSCustomObject]@{ Item = "POLICY THRESHOLDS";                    Value = "" },
    [PSCustomObject]@{ Item = "DISABLE threshold (inactive days)";    Value = "$FinalDisable days" },
    [PSCustomObject]@{ Item = "DELETE  threshold (inactive days)";    Value = "$FinalDelete days" },
    [PSCustomObject]@{ Item = "DELETE already-disabled after";        Value = "$FinalDisDel days since last AD change" },
    [PSCustomObject]@{ Item = "";                                     Value = "" },
    [PSCustomObject]@{ Item = "AD RECYCLE BIN";                       Value = $RBStatus },
    [PSCustomObject]@{ Item = "";                                     Value = "" },
    [PSCustomObject]@{ Item = "TOTALS";                               Value = "" },
    [PSCustomObject]@{ Item = "Total objects";                        Value = "$TotalCount" },
    [PSCustomObject]@{ Item = "  Active";                             Value = "$nA" },
    [PSCustomObject]@{ Item = "  Recommend DISABLE";                  Value = "$nDi" },
    [PSCustomObject]@{ Item = "  Recommend DELETE";                   Value = "$nDe" },
    [PSCustomObject]@{ Item = "  Already Disabled (in AD)";           Value = "$nDs" },
    [PSCustomObject]@{ Item = "  Monitor";                            Value = "$nMo" },
    [PSCustomObject]@{ Item = "";                                     Value = "" },
    [PSCustomObject]@{ Item = "LEGEND";                               Value = "" },
    [PSCustomObject]@{ Item = "ACTIVE";                               Value = "Last activity < $FinalDisable days - no action" },
    [PSCustomObject]@{ Item = "DISABLE (inactive >Xd)";               Value = "Inactive $FinalDisable - $FinalDelete days - disable" },
    [PSCustomObject]@{ Item = "DELETE  (inactive >Xd)";               Value = "Inactive > $FinalDelete days - delete" },
    [PSCustomObject]@{ Item = "DISABLED - pending review";            Value = "Already disabled - review before deletion" },
    [PSCustomObject]@{ Item = "DELETE (disabled >Xd)";               Value = "Disabled > $FinalDisDel days since last change" },
    [PSCustomObject]@{ Item = "MONITOR";                              Value = "No logon, recently created - watch 30 days" }
)

$WB = New-Workbook
try {
    Save-Workbook -WB $WB -Path $FinalOutput `
        -DeviceData $DeviceData -DisplayCols $DisplayCols `
        -RecoCol $RecoColIdx -StatusCol $StatusColIdx `
        -DevSheetName $DevSheetName -SummaryRows $SummaryRows -ActionLog $null
    Write-Host ("  Saved: {0}" -f $FinalOutput) -ForegroundColor Green
} catch {
    Write-Error "Excel export failed: $_"
    exit 1
}

# =============================================================================
#  STEP 4 - ACTION MENU
# =============================================================================
Write-Host ""
Write-Host "[4/4] Action Phase" -ForegroundColor Yellow
Write-Host ""

$ActionLog = [System.Collections.Generic.List[PSCustomObject]]::new()

if (-not $Interactive) {
    Write-Host "  Export-only mode. Add -Interactive to access the action menu." -ForegroundColor DarkGray
} else {

    :ActionLoop while ($true) {

        $nDi = @($DeviceData | Where-Object { $_.Recommendation -like "DISABLE (*" }).Count
        $nDe = @($DeviceData | Where-Object { $_.Recommendation -like "DELETE*"    }).Count
        $nDs = @($DeviceData | Where-Object { $_.Recommendation -like "DISABLED*"  }).Count

        Write-Host ""
        Write-Host "  +-------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host ("  |  ACTION MENU  -  {0,-42}|" -f $Domain)                       -ForegroundColor Cyan
        Write-Host ("  |  Thresholds : Disable>{0}d  Delete>{1}d  Disabled>{2}d       |" -f $FinalDisable, $FinalDelete, $FinalDisDel) -ForegroundColor Cyan
        Write-Host ("  |  Targets    : DISABLE={0,-5} DELETE={1,-5} ALREADY-DISABLED={2,-3}|" -f $nDi, $nDe, $nDs) -ForegroundColor Cyan
        Write-Host "  +-------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host "  |  [0]  Exit - no AD changes                                  |" -ForegroundColor Gray
        Write-Host ("  |  [1]  DISABLE devices inactive > {0,-3} days                 |" -f $FinalDisable) -ForegroundColor Yellow
        Write-Host ("  |  [2]  DELETE  devices inactive > {0,-3} days                 |" -f $FinalDelete)  -ForegroundColor Red
        Write-Host ("  |  [3]  DISABLE>{0}d AND DELETE>{1}d (combined)              |" -f $FinalDisable, $FinalDelete) -ForegroundColor Magenta
        Write-Host ("  |  [4]  DELETE  already-disabled > {0,-3} days                 |" -f $FinalDisDel) -ForegroundColor DarkRed
        Write-Host "  |  [5]  Change thresholds                                     |" -ForegroundColor Cyan
        Write-Host "  +-------------------------------------------------------------+" -ForegroundColor Cyan
        Write-Host ""
        $Choice = Read-Host "  Enter choice (0-5)"

        if ($Choice -eq "0") { Write-Host "  No AD changes made." -ForegroundColor DarkGray; break ActionLoop }

        if ($Choice -eq "5") {
            $FinalDisable = Read-IntPrompt "  New DISABLE threshold (days)" "$FinalDisable"
            $FinalDelete  = Read-IntPrompt "  New DELETE  threshold (days)" "$FinalDelete"
            $FinalDisDel  = Read-IntPrompt "  New DISABLED-DELETE threshold (days)" "$FinalDisDel"
            Write-Host "  Refreshing data..." -ForegroundColor Yellow
            try {
                $DeviceData = @(Get-DomainDevices -Dom $Domain -DisThr $FinalDisable -DelThr $FinalDelete -DisDel $FinalDisDel)
                $TotalCount = $DeviceData.Count
            } catch {
                Write-Host "  Warning: re-query failed - using cached data." -ForegroundColor DarkYellow
            }
            continue ActionLoop
        }

        if ($Choice -notin "1","2","3","4") { Write-Host "  Invalid choice." -ForegroundColor DarkGray; continue ActionLoop }

        $ToDisable = [System.Collections.Generic.List[object]]::new()
        $ToDelete  = [System.Collections.Generic.List[object]]::new()

        foreach ($Dev in $DeviceData) {
            switch ($Choice) {
                "1" { if ($Dev.Recommendation -like "DISABLE (*")        { $ToDisable.Add($Dev) } }
                "2" { if ($Dev.Recommendation -like "DELETE*")            { $ToDelete.Add($Dev)  } }
                "3" { if ($Dev.Recommendation -like "DISABLE (*")        { $ToDisable.Add($Dev) }
                       if ($Dev.Recommendation -like "DELETE*")           { $ToDelete.Add($Dev)  } }
                "4" { if ($Dev.Recommendation -like "DELETE (disabled*") { $ToDelete.Add($Dev)  } }
            }
        }

        if ($ToDisable.Count + $ToDelete.Count -eq 0) {
            Write-Host "  No targets found for this selection." -ForegroundColor DarkGray
            continue ActionLoop
        }

        Write-Host ("  Devices to DISABLE : {0}" -f $ToDisable.Count) -ForegroundColor Yellow
        Write-Host ("  Devices to DELETE  : {0}" -f $ToDelete.Count)  -ForegroundColor Red

        if ($WhatIfPreference) {
            Write-Host "  [WHATIF] No changes applied." -ForegroundColor Magenta
            $ToDisable | ForEach-Object { Write-Host ("    DISABLE: {0}" -f $_."Device Name") -ForegroundColor DarkYellow }
            $ToDelete  | ForEach-Object { Write-Host ("    DELETE : {0}" -f $_."Device Name") -ForegroundColor DarkRed    }
        } else {
            $Conf = Read-Host ("  Apply changes to {0}? (Y/N)" -f $Domain)
            if ($Conf -notmatch "^[Yy]") { Write-Host "  Skipped." -ForegroundColor DarkGray; continue ActionLoop }

            foreach ($Dev in $ToDisable) {
                $Res = "OK"; $Err = ""
                try {
                    Disable-ADAccount -Identity $Dev._DN -Server $Domain -Confirm:$false
                    Write-Host ("  DISABLED: {0}" -f $Dev."Device Name") -ForegroundColor Yellow
                } catch {
                    $Res = "ERROR"; $Err = $_.Exception.Message
                    Write-Host ("  ERROR: {0}: {1}" -f $Dev."Device Name", $Err) -ForegroundColor Red
                }
                $ActionLog.Add([PSCustomObject]@{
                    Timestamp       = (Get-Date -f "yyyy-MM-dd HH:mm:ss")
                    Domain          = $Domain
                    Action          = "DISABLE"
                    "Device Name"   = $Dev."Device Name"
                    "Days Inactive" = $Dev."Days Inactive"
                    DN              = $Dev._DN
                    Result          = $Res
                    Error           = $Err
                })
            }
            foreach ($Dev in $ToDelete) {
                $Res = "OK"; $Err = ""
                try {
                    Remove-ADComputer -Identity $Dev._DN -Server $Domain -Confirm:$false
                    Write-Host ("  DELETED : {0}" -f $Dev."Device Name") -ForegroundColor Red
                } catch {
                    $Res = "ERROR"; $Err = $_.Exception.Message
                    Write-Host ("  ERROR: {0}: {1}" -f $Dev."Device Name", $Err) -ForegroundColor Red
                }
                $ActionLog.Add([PSCustomObject]@{
                    Timestamp       = (Get-Date -f "yyyy-MM-dd HH:mm:ss")
                    Domain          = $Domain
                    Action          = "DELETE"
                    "Device Name"   = $Dev."Device Name"
                    "Days Inactive" = $Dev."Days Inactive"
                    DN              = $Dev._DN
                    Result          = $Res
                    Error           = $Err
                })
            }
        }

        $Again = Read-Host "  Perform another action? (Y/N)"
        if ($Again -notmatch "^[Yy]") { break ActionLoop }

    } # end ActionLoop

    if ($ActionLog.Count -gt 0) {
        Write-Host "  Writing action log to workbook..." -ForegroundColor Yellow
        try {
            $WB2 = New-Workbook
            Save-Workbook -WB $WB2 -Path $FinalOutput `
                -DeviceData $DeviceData -DisplayCols $DisplayCols `
                -RecoCol $RecoColIdx -StatusCol $StatusColIdx `
                -DevSheetName $DevSheetName -SummaryRows $SummaryRows -ActionLog $ActionLog
            Write-Host "  Action log saved." -ForegroundColor Green
        } catch {
            Write-Host ("  Warning: action log write failed - {0}" -f $_) -ForegroundColor DarkYellow
        }
    }
}

# =============================================================================
#  FINAL SUMMARY
# =============================================================================
$nA  = @($DeviceData | Where-Object { $_.Recommendation -eq "ACTIVE"       }).Count
$nDi = @($DeviceData | Where-Object { $_.Recommendation -like "DISABLE (*" }).Count
$nDe = @($DeviceData | Where-Object { $_.Recommendation -like "DELETE*"    }).Count
$nDs = @($DeviceData | Where-Object { $_.Recommendation -like "DISABLED*"  }).Count
$nMo = @($DeviceData | Where-Object { $_.Recommendation -like "MONITOR*"   }).Count

Write-Host ""
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "|                      S U M M A R Y                         |" -ForegroundColor Cyan
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("|  Domain              : {0,-37}|" -f $Domain)                  -ForegroundColor White
Write-Host ("|  Total objects       : {0,-37}|" -f $TotalCount)              -ForegroundColor White
Write-Host ("|  Active              : {0,-37}|" -f $nA)                      -ForegroundColor Green
Write-Host ("|  Recommend DISABLE   : {0,-37}|" -f $nDi)                     -ForegroundColor Yellow
Write-Host ("|  Recommend DELETE    : {0,-37}|" -f $nDe)                     -ForegroundColor Red
Write-Host ("|  Already Disabled    : {0,-37}|" -f $nDs)                     -ForegroundColor DarkYellow
Write-Host ("|  Monitor             : {0,-37}|" -f $nMo)                     -ForegroundColor Cyan
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("|  Disable threshold   : {0,-37}|" -f "$FinalDisable days")     -ForegroundColor White
Write-Host ("|  Delete  threshold   : {0,-37}|" -f "$FinalDelete days")      -ForegroundColor White
Write-Host ("|  Disabled-delete     : {0,-37}|" -f "$FinalDisDel days")      -ForegroundColor White
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ("|  Recycle Bin         : {0,-37}|" -f $RBStatus)                -ForegroundColor White
Write-Host ("|  PowerShell          : {0,-37}|" -f "PowerShell $PSVer")      -ForegroundColor White
Write-Host ("|  Run As              : {0,-37}|" -f $RunAsText) -ForegroundColor $(if ($IsAdmin) { "White" } else { "Yellow" })
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host "|  Output:                                                    |" -ForegroundColor White
Write-Host ("|  {0,-59}|" -f $FinalOutput)                                   -ForegroundColor White
Write-Host "+-------------------------------------------------------------+" -ForegroundColor Cyan
Write-Host ""
if ($WhatIfPreference) {
    Write-Host "  [WHATIF mode - no AD changes were applied]" -ForegroundColor Magenta
    Write-Host ""
}
