[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [int]$Depth = -1,

    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "$PSScriptRoot\PermissionsReport.xlsx"
)

begin {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    Write-Host "`n>>> Starting..." -ForegroundColor Magenta
}

process {
    if (!(Get-Module -ListAvailable ImportExcel)) { Write-Error "ImportExcel missing."; return }
    Import-Module ImportExcel
    Add-Type -AssemblyName System.Drawing

    # --- 1. SCANNING ---
    function Get-SimplifiedACL($folderPath) {
        try {
            $acl = Get-Acl $folderPath -ErrorAction Stop
            return ($acl.Access | Where-Object { $_.IdentityReference -notmatch "SYSTEM|NT AUTHORITY" } | 
                    Select-Object IdentityReference, FileSystemRights, AccessControlType | 
                    Sort-Object IdentityReference | Out-String).Trim()
        } catch { return "ACCESS_DENIED" }
    }

    Write-Host ">>> Step 1: Scanning for permission changes..." -ForegroundColor Cyan
    $rootFolder = Get-Item $Path
    $uniqueFolders = New-Object System.Collections.Generic.List[PSObject]
    $uniqueFolders.Add($rootFolder)

    $rawFolders = Get-ChildItem -Path $Path -Directory -Recurse -ErrorAction SilentlyContinue
    foreach ($f in $rawFolders) {
        $parent = Split-Path $f.FullName -Parent
        if ((Get-SimplifiedACL $f.FullName) -ne (Get-SimplifiedACL $parent)) {
            $uniqueFolders.Add($f)
        }
    }
    Write-Host "    Found $($uniqueFolders.Count) unique columns." -ForegroundColor Gray

    # --- 2. MAPPING ---
    Write-Host ">>> Step 2: Mapping hierarchy..." -ForegroundColor Cyan
    $baseParent = Split-Path $rootFolder.FullName -Parent
    $columnsPathMap = New-Object System.Collections.Generic.List[PSObject]
    $maxDepth = 0
    foreach ($folder in $uniqueFolders) {
        $parts = ($folder.FullName.Substring($baseParent.Length).TrimStart('\')).Split('\')
        if ($parts.Count -gt $maxDepth) { $maxDepth = $parts.Count }
        $columnsPathMap.Add($parts)
    }
    [int]$headerRowsCount = $maxDepth

    # --- 3. USERS ---
    Write-Host ">>> Step 3: Resolving ACL Identities..." -ForegroundColor Cyan
    $userPermissionsMaster = @{} 
    foreach ($folder in $uniqueFolders) {
        try {
            $acl = Get-Acl $folder.FullName
            foreach ($access in $acl.Access) {
                if ($access.IdentityReference -match "NT AUTHORITY|BUILTIN|SYSTEM") { continue }
                $cleanId = $access.IdentityReference.Value.Split('\')[-1]
                $rights = if ($access.FileSystemRights.ToString() -match "FullControl|Modify|Write") { "W" } else { "R" }
                $uKey = $cleanId.ToLower().Trim()
                if (!$userPermissionsMaster.ContainsKey($uKey)) { $userPermissionsMaster[$uKey] = @{} }
                if ($userPermissionsMaster[$uKey][$folder.FullName] -ne "W") { $userPermissionsMaster[$uKey][$folder.FullName] = $rights }
            }
        } catch {}
    }

    # --- 4. DATA ---
    Write-Host ">>> Step 4: Filling Excel Data Rows..." -ForegroundColor Cyan
    $excel = New-Object OfficeOpenXml.ExcelPackage
    $sheet = $excel.Workbook.Worksheets.Add("Permissions")
    $sortedUsers = $userPermissionsMaster.Keys | Sort-Object
    [int]$colTracker = 3
    foreach ($folder in $uniqueFolders) {
        $currentRow = $headerRowsCount + 1
        foreach ($u in $sortedUsers) {
            if ($colTracker -eq 3) {
                $sheet.Cells[$currentRow, 1].Value = $u
                $sheet.Cells[$currentRow, 2].Value = $u
            }
            $val = if ($userPermissionsMaster[$u][$folder.FullName]) { $userPermissionsMaster[$u][$folder.FullName] } else { "-" }
            $cell = $sheet.Cells[$currentRow, $colTracker]
            $cell.Value = $val
            $cell.Style.Fill.PatternType = 'Solid'
            $color = if ($val -eq "W") { [System.Drawing.Color]::LightGreen } elseif ($val -eq "R") { [System.Drawing.Color]::LightSkyBlue } else { [System.Drawing.Color]::LightPink }
            $cell.Style.Fill.BackgroundColor.SetColor($color)
            $currentRow++
        }
        $colTracker++
    }

    # --- 5. HEADERS ---
    Write-Host ">>> Step 5: Formatting tree header..." -ForegroundColor Cyan
    for ($i = 0; $i -lt $columnsPathMap.Count; $i++) {
        $parts = $columnsPathMap[$i]
        [int]$targetCol = $i + 3
        for ($r = 1; $r -le $headerRowsCount; $r++) {
            if ($r -le $parts.Count) { $sheet.Cells[$r, $targetCol].Value = $parts[$r-1] }
        }
    }

    [int]$maxColIndex = $columnsPathMap.Count + 2

    # Horizontal Merge
    for ($r = 1; $r -le $headerRowsCount; $r++) {
        for ($c = 3; $c -le $maxColIndex; $c++) {
            $val = $sheet.Cells[$r, $c].Value
            if ([string]::IsNullOrWhiteSpace($val)) { continue }
            [int]$startCol = $c
            [int]$nextC = $c + 1
            while ($nextC -le $maxColIndex) {
                if ($sheet.Cells[$r, $nextC].Value -eq $val) {
                    $c = $nextC
                    $nextC++
                } else { break }
            }
            if ($c -gt $startCol) { $sheet.Cells[$r, $startCol, $r, $c].Merge = $true }
        }
    }

    # Vertical Merge
    for ($vc = 3; $vc -le $maxColIndex; $vc++) {
        for ($vr = 1; $vr -le $headerRowsCount; $vr++) {
            if ($sheet.Cells[$vr, $vc].Merge) { continue }
            $vVal = $sheet.Cells[$vr, $vc].Value
            if ([string]::IsNullOrWhiteSpace($vVal)) { continue }
            [int]$startR = $vr
            [int]$nextR = $vr + 1
            while ($nextR -le $headerRowsCount) {
                $belowVal = $sheet.Cells[$nextR, $vc].Value
                if ([string]::IsNullOrWhiteSpace($belowVal)) {
                    $vr = $nextR
                    $nextR++
                } else { break }
            }
            if ($vr -gt $startR) { $sheet.Cells[$startR, $vc, $vr, $vc].Merge = $true }
        }
    }

    # --- 6. STYLES & SAVE ---
    Write-Host ">>> Step 6: Saving file..." -ForegroundColor Cyan
    $sheet.Cells[1, 1].Value = "Login"; $sheet.Cells[1, 2].Value = "Full Name"
    if ($headerRowsCount -gt 1) {
        $sheet.Cells[1, 1, $headerRowsCount, 1].Merge = $true
        $sheet.Cells[1, 2, $headerRowsCount, 2].Merge = $true
    }
    $headerRange = $sheet.Cells[1, 1, $headerRowsCount, $maxColIndex]
    $headerRange.Style.HorizontalAlignment = 'Center'
    $headerRange.Style.VerticalAlignment = 'Center'
    $headerRange.Style.Font.Bold = $true
    $headerRange.Style.Border.BorderAround('Thin')

    $sheet.Cells.AutoFitColumns()
    $excel.SaveAs($OutputFile)
    $excel.Dispose()
    Write-Host "`n>>> Success! Final Report: $OutputFile" -ForegroundColor Green
}