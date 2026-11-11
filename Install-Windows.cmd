<# :
@echo off
setlocal
title Client Export Tool - Installer
set "CET_SELF=%~f0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression ([System.IO.File]::ReadAllText($env:CET_SELF))"
set "CET_RC=%ERRORLEVEL%"
echo.
pause
exit /b %CET_RC%
#>

# ---------------------------------------------------------------------------
# Client Export Tool - Windows installer
#
# Double-click Install-Windows.cmd. The few batch lines above only launch the
# PowerShell below (this file is valid as both). It:
#   1. copies ClientExportTool.xlam into %APPDATA%\Microsoft\AddIns
#   2. removes the "downloaded from the internet" block (Mark of the Web)
#   3. registers the add-in with Excel and removes stale entries left by
#      earlier installs (for example one pointing at the Downloads folder)
# Run it again at any time to update to a newer version.
#
# Keep this file ASCII-only with CRLF line endings (see .gitattributes).
# ---------------------------------------------------------------------------

$ErrorActionPreference = 'Stop'
$AddinFile = 'ClientExportTool.xlam'
$ExcelKey  = 'Software\Microsoft\Office\16.0\Excel'
$Pattern   = 'ClientExportTool[^\\"]*\.xlam'

function Ok($msg)   { Write-Host $msg -ForegroundColor Green }
function Warn($msg) { Write-Host $msg -ForegroundColor Yellow }
function Fail($msg) { Write-Host ''; Write-Host "ERROR: $msg" -ForegroundColor Red; exit 1 }

Write-Host 'Client Export Tool - Windows installer' -ForegroundColor Cyan
Write-Host ''

# 1. Find the add-in next to this installer --------------------------------
$here  = Split-Path -Parent $env:CET_SELF
$found = @(Get-ChildItem -LiteralPath $here -Filter 'ClientExportTool*.xlam' -File -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending)
if ($found.Count -eq 0) {
    if ($here -match '\.zip(\\|$)') {
        Fail "The installer was opened from inside the zip. Right-click the zip, choose 'Extract All...', then run Install-Windows.cmd from the extracted folder."
    }
    Fail "Couldn't find $AddinFile in $here. Keep Install-Windows.cmd in the same folder as the add-in. If you opened it from inside the zip, extract the zip first (right-click > Extract All...)."
}
$src = $found[0].FullName
if ($found.Count -gt 1) { Warn "More than one add-in file is in this folder; using the newest: $($found[0].Name)" }

# 2. Excel must be closed (it holds the add-in open and rewrites its add-in
#    list from memory when it exits) -----------------------------------------
while (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) {
    Warn 'Excel is running. Save your work, close ALL Excel windows, then press Enter here.'
    Warn '(If no Excel window is visible, end EXCEL.EXE in Task Manager. Press Ctrl+C to cancel.)'
    [void](Read-Host)
}

# 3. Copy and unblock -------------------------------------------------------
$destDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$dest    = Join-Path $destDir $AddinFile
try {
    New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    if ($src -ne $dest) { Copy-Item -LiteralPath $src -Destination $dest -Force }
    Unblock-File -LiteralPath $dest
} catch {
    Fail "Couldn't copy the add-in to $destDir. $($_.Exception.Message)"
}
if (Get-Item -LiteralPath $dest -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue) {
    Fail "Windows still marks $dest as downloaded. Right-click it > Properties > tick 'Unblock' > OK, then run this installer again."
}
Ok "Installed and unblocked: $dest"

# 4. Register with Excel ----------------------------------------------------
# Excel loads checked add-ins from OPEN, OPEN1, OPEN2... under Options (read in
# order, so the numbering must stay contiguous). Unchecked add-ins are listed
# under "Add-in Manager". Drop any ClientExportTool entry, then add ours.
$hkcu  = [Microsoft.Win32.Registry]::CurrentUser
$stale = 0
try {
    $opts = $hkcu.CreateSubKey("$ExcelKey\Options")
    $openNames = @($opts.GetValueNames() | Where-Object { $_ -match '^OPEN\d*$' } |
                   Sort-Object { if ($_.Length -eq 4) { 0 } else { [int]$_.Substring(4) } })
    $keep = New-Object 'System.Collections.Generic.List[string]'
    foreach ($n in $openNames) {
        $v = [string]$opts.GetValue($n)
        if ($v -match $Pattern) {
            if ($v.IndexOf($dest, [StringComparison]::OrdinalIgnoreCase) -lt 0) { $stale++ }
        } else {
            $keep.Add($v)
        }
    }
    $keep.Add('"' + $dest + '"')
    foreach ($n in $openNames) { $opts.DeleteValue($n) }
    for ($i = 0; $i -lt $keep.Count; $i++) {
        $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
        $opts.SetValue($name, $keep[$i], [Microsoft.Win32.RegistryValueKind]::String)
    }
    $opts.Close()

    $mgr = $hkcu.OpenSubKey("$ExcelKey\Add-in Manager", $true)
    if ($mgr) {
        foreach ($n in $mgr.GetValueNames()) {
            if ($n -match $Pattern) {
                if ($n.IndexOf($dest, [StringComparison]::OrdinalIgnoreCase) -lt 0) { $stale++ }
                $mgr.DeleteValue($n)
            }
        }
        $mgr.Close()
    }
} catch {
    Fail "The add-in was copied but couldn't be registered with Excel: $($_.Exception.Message). Finish in Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse > $dest"
}
Ok 'Registered with Excel.'
if ($stale -gt 0) { Write-Host "Removed $stale outdated add-in entry(s) from a previous install." }

Write-Host ''
Ok 'Done! Open Excel and go to Review > Client Export > Create Client Export.'
exit 0
