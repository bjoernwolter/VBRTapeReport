<#
.SYNOPSIS
    Displays the current tape media in Veeam Backup & Replication, with information on barcode, media set,
    expiration date, last write time, slot, and free space.
.DESCRIPTION
    The script filters out cleaning tapes, checks the expiration date, and highlights with colors:
        - Red: Expiration date is in the future
        - Green: Tape has expired
        - Yellow: No expiration date available
    Additionally, the last write date and slot are displayed.

.AUTHOR
    Björn Wolter

.VERSION
    1.1

.LASTEDIT
    2025-05-19

.REQUIREMENTS
    Veeam Backup & Replication PowerShell Module

#>

# Load Module
Import-Module Veeam.Backup.PowerShell -DisableNameChecking

# get tapes, without cleaning tapes
$tapes = Get-VBRTapeMedium | Where-Object {
    $_.Location -like "Slot" -and $_.Barcode -notlike "CLN*"
}

# sort ob mediaset name
$tapes = $tapes | Sort-Object { if ($_.MediaSet) { $_.MediaSet.Name } else { "" } }

# Header
Write-Host ("{0,-15} {1,-42} {2,-25} {3,-12} {4,-6} {5,8}" -f "Barcode", "Mediaset", "Expiration", "Last Write", "Slot", "Free (GB)")
Write-Host ("".PadRight(114, "-"))

# Printout lines with colors
foreach ($tape in $tapes) {
    $barcode = $tape.Barcode

    # MediaSet
    $mediaset = if ($tape.MediaSet -and $tape.MediaSet.Name) {
        $tape.MediaSet.Name.Substring(0, [math]::Min(40, $tape.MediaSet.Name.Length))
    } else {
        "n/a"
    }

    # Slot
    $slot = if ($tape.Location -and $tape.Location.SlotAddress -ne $null) {
        $tape.Location.SlotAddress + 1
    } else {
        "n/a"
    }

    # Free space
    $freeGB = if ($tape.Free -ne $null) {
        [math]::Round($tape.Free / 1TB, 2)
    } else {
        "n/a"
    }

    # Last write time
    $lastWrite = if ($tape.LastWriteTime -ne $null) {
        $tape.LastWriteTime.ToString("yyyy-MM-dd")
    } else {
        "n/a"
    }

    # Expiration handling
    $now = Get-Date
    $expDate = $tape.ExpirationDate

    if ($expDate -eq $null) {
        $expString = "n/a"
        $color = "Yellow"
    }
    elseif ($expDate -lt $now) {
        $expString = "Expired ($($expDate.ToString('yyyy-MM-dd')))"
        $color = "Green"
    } else {
        $expString = $expDate.ToString('yyyy-MM-dd')
        $color = "Red"
    }

    # write everything to console
    Write-Host ("{0,-15} {1,-42} {2,-25} {3,-12} {4,-6} {5,8}" -f $barcode, $mediaset, $expString, $lastWrite, $slot, $freeGB) -ForegroundColor $color
}

