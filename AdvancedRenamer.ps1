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
    Expected columns: "Old Name", "New Name" (or specified via -map).

.PARAMETER targetdir
    The directory containing the files to be renamed. Default is current directory.

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
    return $newVal
}

function Read-CsvData {
    param($Path, $OldIdx, $NewIdx)
    try {
        # Import CSV without headers to treat by index
        $raw = Import-Csv $Path -Header "C0","C1","C2","C3","C4","C5" # Dummy headers for up to 6 cols
        # Skip user header if needed? Usually Import-Csv handles headers nicely if we know names.
        # But we are using indices.
        
        # Better approach: Read raw content? No, Import-Csv is best.
        # Let's assume standard headers "Old Name" and "New Name" OR use indices logic manually.
        # For simplicity in this hybrid tool, let's treat it by index manually.
        
        $lines = Get-Content $Path
        $data = @()
        # Skip header row 1
        for ($i = 1; $i -lt $lines.Count; $i++) {
            $cols = $lines[$i] -split ","
            if ($cols.Count -gt $NewIdx) {
                $o = $cols[$OldIdx].Trim('"').Trim()
                $n = $cols[$NewIdx].Trim('"').Trim()
                if ($o -and $n) {
                    $data += [PSCustomObject]@{ OldName = $o; NewName = $n }
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
    $Mapping = Read-CsvData -Path $FullPathInput -OldIdx $oidx -NewIdx $nidx
} else {
    Write-ErrorMsg "Unsupported file type. Please use .csv"
    exit 1
}

if ($Mapping.Count -eq 0) { Write-Info "No renaming rules found."; exit }

# Get Files
$ResolvedTarget = Resolve-Path $targetdir
Write-Info "Target Directory: $ResolvedTarget"
$files = Get-ChildItem -Path $ResolvedTarget -File -Recurse:$subfolders

Write-Info "Found $($files.Count) files. Processing $($Mapping.Count) rules..."

$ops = @()
$OccupiedPaths = @() # Track paths we reserve in this run

foreach ($rule in $Mapping) {
    # Find the file in our list (Case insensitive matches roughly in Windows)
    # Simple search: Match by Name property
    $matches = $files | Where-Object { $_.Name -eq $rule.OldName }
    
    if (-not $matches) {
        Write-Warn "File not found: '$($rule.OldName)'"
        continue
    }
    
    foreach ($file in $matches) {
        $dir = $file.DirectoryName
        $newName = $rule.NewName
        $destPath = Join-Path $dir $newName
        
        # Check against Disk OR our reserved list
        if ((Test-Path $destPath) -or ($OccupiedPaths -contains $destPath)) {
            if ($conflict -eq "skip") {
                Write-Warn "Target '$newName' exists. Skipping."
                continue
            } elseif ($conflict -eq "autonumber") {
                $newName = Get-UniqueFilename -dir $dir -filename $newName -occupiedPaths $OccupiedPaths
                $destPath = Join-Path $dir $newName
                Write-Info "Conflict resolved: Renaming to '$newName'"
            } elseif ($conflict -eq "overwrite") {
                Write-Warn "Overwriting '$newName'"
                # Remove dest so rename works
                if (-not $dryrun) { Remove-Item $destPath -Force }
            }
        }
        
        # Mark this path as occupied for subsequent checks in this batch
        $OccupiedPaths += $destPath

        if ($dryrun) {
            Write-Host "[PREVIEW] '$($file.Name)' -> '$newName'" -ForegroundColor Cyan
        } else {
            try {
                Rename-Item -LiteralPath $file.FullName -NewName $newName
                Write-Success "'$($file.Name)' -> '$newName'"
                $ops += @{ source = $file.FullName; dest = $destPath }
            } catch {
                Write-ErrorMsg "Failed to rename '$($file.Name)': $($_.Exception.Message)"
            }
        }
    }
}

if (-not $dryrun -and $ops.Count -gt 0) {
    Save-Log $ops
    Write-Host "Tip: You can undo this operation by running: rename_tool.bat -undo -targetdir `"$ResolvedTarget`"" -ForegroundColor Magenta
}

Write-Info "Done."
