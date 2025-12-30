<#
.SYNOPSIS
    Advanced Batch File Renamer
    Renames files based on an Excel (.xlsx) or CSV mapping.

.DESCRIPTION
    Takes a CSV mapping file, finds the old filenames in the target directory,
    and renames them to the new filenames.
    Supports Undo, Dry-Run, Recursion, and Conflict Resolution strategies.

.PARAMETER inputfile
    Path to the .csv file containing the mapping.
    Expected columns: "Folder Path", "Old Name", "New Name"
    OR Legacy columns: "Old Name", "New Name" (uses -targetdir)

.PARAMETER targetdir
    The directory containing the files to be renamed. 
    Only used if CSV doesn't have a "Folder Path" column.

.PARAMETER map
    Column indices (0-based) for OldName,NewName. Default "0,1".
    Example: "0,1" means Column A is Old, Column B is New.

.PARAMETER conflict
    Strategy for dealing with existing files: "skip", "overwrite", "autonumber".
    Default: "skip"

.PARAMETER subfolders
    If set, searches for "Old Name" in subdirectories.

.PARAMETER dryrun
    If set, only previews the valid operations.

.PARAMETER undo
    If set, attempts to undo the last batch of operations from the log file.
#>

param(
    [string]$inputfile,
    [string]$targetdir = ".",
    [string]$map = "0,1",
    [ValidateSet("skip", "overwrite", "autonumber")]
    [string]$conflict = "skip",
    [switch]$subfolders,
    [switch]$dryrun,
    [switch]$undo
)

$ErrorActionPreference = "Stop"
$LogFile = Join-Path $targetdir "rename_log.json"

# --- Helper Functions ---

function Write-Success ($msg) { Write-Host "[SUCCESS] $msg" -ForegroundColor Green }
function Write-Warn ($msg)    { Write-Host "[WARNING] $msg" -ForegroundColor Yellow }
function Write-ErrorMsg ($msg){ Write-Host "[ERROR] $msg" -ForegroundColor Red }
function Write-Info ($msg)    { Write-Host "[INFO] $msg" -ForegroundColor Cyan }

function Get-UniqueFilename ($dir, $filename, $occupiedPaths) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($filename)
    $ext = [System.IO.Path]::GetExtension($filename)
    $count = 1
    $newVal = $filename
    
    # Check physical existence AND virtual existence (for DryRun/Batch consistency)
    while ((Test-Path (Join-Path $dir $newVal)) -or ($occupiedPaths -contains (Join-Path $dir $newVal))) {
        $newVal = "${name}_${count}${ext}"
        $count++
    }

function Process-FileRename {
    param($file, $newName, $conflict, [ref]$OccupiedPaths, $dryrun, [ref]$ops)
    
    $dir = $file.DirectoryName
    $destPath = Join-Path $dir $newName
    
    # Check Conflict
    if ((Test-Path $destPath) -or ($OccupiedPaths.Value -contains $destPath)) {
        if ($conflict -eq "skip") {
            Write-Warn "Target '$newName' exists. Skipping."
            return
        } elseif ($conflict -eq "autonumber") {
            $newName = Get-UniqueFilename -dir $dir -filename $newName -occupiedPaths $OccupiedPaths.Value
            $destPath = Join-Path $dir $newName
            Write-Info "Conflict resolved: Renaming to '$newName'"
        } elseif ($conflict -eq "overwrite") {
            Write-Warn "Overwriting '$newName'"
            if (-not $dryrun) { Remove-Item $destPath -Force }
        }
    }
    
    $OccupiedPaths.Value += $destPath
    
    if ($dryrun) {
        Write-Host "[PREVIEW] '$($file.Name)' -> '$newName'" -ForegroundColor Cyan
    } else {
        try {
            Rename-Item -LiteralPath $file.FullName -NewName $newName
            Write-Success "'$($file.Name)' -> '$newName'"
            $ops.Value += @{ source = $file.FullName; dest = $destPath }
        } catch {
            Write-ErrorMsg "Failed to rename '$($file.Name)': $($_.Exception.Message)"
        }
    }
}
    return $newVal
}

function Read-CsvData {
    param($Path)
    try {
        $lines = Get-Content $Path
        $data = @()
        
        # Detect mode based on column count of first data row
        if ($lines.Count -gt 1) {
            $cols = $lines[1] -split ","
            if ($cols.Count -ge 3) {
                # 3-Column Mode: Folder, Old, New
                Write-Info "Detected 3-column CSV (Folder, Old, New)"
                for ($i = 1; $i -lt $lines.Count; $i++) {
                    $cols = $lines[$i] -split ","
                    if ($cols.Count -ge 3) {
                        $p = $cols[0].Trim('"').Trim()
                        $o = $cols[1].Trim('"').Trim()
                        $n = $cols[2].Trim('"').Trim()
                        if ($o -and $n) {
                            $data += [PSCustomObject]@{ Path=$p; OldName=$o; NewName=$n; Mode="Direct" }
                        }
                    }
                }
            } else {
                # 2-Column Mode: Old, New
                Write-Info "Detected 2-column CSV (Old, New) - Using TargetDir"
                for ($i = 1; $i -lt $lines.Count; $i++) {
                    $cols = $lines[$i] -split ","
                    if ($cols.Count -ge 2) {
                        $o = $cols[0].Trim('"').Trim()
                        $n = $cols[1].Trim('"').Trim()
                        if ($o -and $n) {
                            $data += [PSCustomObject]@{ Path=$null; OldName=$o; NewName=$n; Mode="Scan" }
                        }
                    }
                }
            }
        }
        return $data
    }
    catch {
        Write-ErrorMsg "Failed to read CSV."
        return @()
    }
}

function Save-Log {
    param($ops)
    $batch = @{
        id = (Get-Date).Ticks
        timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        operations = $ops
    }
    
    $allLogs = @()
    if (Test-Path $LogFile) {
        try { $allLogs = Get-Content $LogFile -Raw | ConvertFrom-Json } catch {}
        if (-not $allLogs) { $allLogs = @() } # Handle null if empty
        if ($allLogs -isnot [array]) { $allLogs = @($allLogs) } # Ensure array
    }
    
    $allLogs += $batch
    $allLogs | ConvertTo-Json -Depth 4 | Set-Content $LogFile
}

function Undo-Last {
    if (-not (Test-Path $LogFile)) { Write-ErrorMsg "No undo log found."; return }
    
    $json = Get-Content $LogFile -Raw | ConvertFrom-Json
    if (-not $json) { Write-Info "Log is empty."; return }
    if ($json -isnot [array]) { $json = @($json) }
    
    $lastBatch = $json[-1] # Get last
    Write-Info "Undoing batch from $($lastBatch.timestamp)..."
    
    $ops = $lastBatch.operations
    # Reverse operations
    for ($i = $ops.Count - 1; $i -ge 0; $i--) {
        $op = $ops[$i]
        $curr = $op.dest
        $orig = $op.source
        
        if (Test-Path $curr) {
            if (Test-Path $orig) {
                Write-Warn "Cannot undo '$curr' -> '$orig': Original file already exists."
            } else {
                Rename-Item -Path $curr -NewName ([System.IO.Path]::GetFileName($orig))
                Write-Success "Restored: $orig"
            }
        } else {
            Write-Warn "File missing: $curr"
        }
    }
    
    # Remove last batch from log
    $newJson = $json[0..($json.Count - 2)]
    if ($null -eq $newJson) { $newJson = @() } # handle empty
    $newJson | ConvertTo-Json -Depth 4 | Set-Content $LogFile
}

# --- Main Execution ---

if ($undo) {
    Undo-Last
    exit
}

if (-not $inputfile) {
    Write-ErrorMsg "Please provide an -inputfile (CSV)."
    exit 1
}

$FullPathInput = Resolve-Path $inputfile
if (-not (Test-Path $FullPathInput)) { Write-ErrorMsg "Input file not found."; exit 1 }

# Parse Mapping Indices
$indices = $map -split ","
$oidx = [int]$indices[0]
$nidx = [int]$indices[1]

Write-Info "Reading mapping..."
$Mapping = @()
if ($inputfile -match "\.csv$") {
    $Mapping = Read-CsvData -Path $FullPathInput
} else {
    Write-ErrorMsg "Unsupported file type. Please use .csv"
    exit 1
}

if ($Mapping.Count -eq 0) { Write-Info "No renaming rules found."; exit }

# Get Files
$ResolvedTarget = Resolve-Path $targetdir
Write-Info "Target Directory: $ResolvedTarget"
# Only scan if we have rules that need scanning (Mode=Scan)
$hasScanRules = $Mapping | Where-Object { $_.Mode -eq "Scan" }
$files = @()
if ($hasScanRules) {
    Write-Info "Scanning for files (Legacy Mode)..."
    $files = Get-ChildItem -Path $ResolvedTarget -File -Recurse:$subfolders
}


Write-Info "Found $($files.Count) files. Processing $($Mapping.Count) rules..."

$ops = @()
$OccupiedPaths = @() # Track paths we reserve in this run

# --- Processing Logic ---

# 1. Handle Direct Mode (Grouped by Folder for Speed)
$DirectRules = $Mapping | Where-Object { $_.Mode -eq "Direct" }
if ($DirectRules) {
    Write-Info "Processing rules (Grouped by Directory)..."
    $Groups = $DirectRules | Group-Object Path
    
    foreach ($grp in $Groups) {
        $targetPath = $grp.Name
        # Validate Folder Once
        if (-not (Test-Path $targetPath)) {
            Write-Warn "Folder not found: '$targetPath'. Skipping $($grp.Count) rules."
            continue
        }
        
        foreach ($rule in $grp.Group) {
            $fullOldPath = Join-Path $targetPath $rule.OldName
            if (-not (Test-Path $fullOldPath)) {
                Write-Warn "File not found: '$($rule.OldName)' in '$targetPath'"
                continue
            }
            
            # Use 'Get-Item' to get a file object similar to the scan mode
            $file = Get-Item $fullOldPath
            
            Process-FileRename -file $file -newName $rule.NewName -conflict $conflict -OccupiedPaths ([ref]$OccupiedPaths) -dryrun $dryrun -ops ([ref]$ops)
        }
    }
}

# 2. Handle Scan Mode (Legacy 2-Column)
$ScanRules = $Mapping | Where-Object { $_.Mode -eq "Scan" }
if ($ScanRules) {
    Write-Info "Processing legacy Scan Rules..."
    foreach ($rule in $ScanRules) {
        $matches = $files | Where-Object { $_.Name -eq $rule.OldName }
        if (-not $matches) {
            Write-Warn "File not found: '$($rule.OldName)'"
            continue
        }
        
        foreach ($file in $matches) {
            Process-FileRename -file $file -newName $rule.NewName -conflict $conflict -OccupiedPaths ([ref]$OccupiedPaths) -dryrun $dryrun -ops ([ref]$ops)
        }
    }
}

if (-not $dryrun -and $ops.Count -gt 0) {
    Save-Log $ops
    Write-Host "Tip: You can undo this operation by running: rename_tool.bat -undo -targetdir `"$ResolvedTarget`"" -ForegroundColor Magenta
}

Write-Info "Done."
