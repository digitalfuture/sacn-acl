[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [int]$Depth = 100,

    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "$PSScriptRoot\PermissionsReport.xlsx"
)

process {
    if (!(Get-Module -ListAvailable ImportExcel)) { Write-Error "ImportExcel module missing."; return }
    Import-Module ImportExcel
    Add-Type -AssemblyName System.Drawing

    # AD Module Check
    $adAvailable = Get-Module -ListAvailable ActiveDirectory
    if ($adAvailable) { try { Import-Module ActiveDirectory -ErrorAction SilentlyContinue } catch { $adAvailable = $false } }

    $groupCache = @{}

    function Get-Members ($identityName) {
        if (!$adAvailable) { return @($identityName) }
        if ($groupCache.ContainsKey($identityName)) { return $groupCache[$identityName] }
        try {
            $adObj = Get-ADObject -Filter "SamAccountName -eq '$identityName'" -ErrorAction SilentlyContinue
            if ($null -ne $adObj -and $adObj.ObjectClass -eq "group") {
                $m = Get-ADGroupMember -Identity $identityName -Recursive | Select-Object -ExpandProperty SamAccountName
                $groupCache[$identityName] = $m
                return $m
            }
        } catch {}
        return @($identityName)
    }

    # 1. SCANNING (DACL-only filter)
    Write-Host ">>> Step 1: Scanning folders..." -ForegroundColor Cyan
    $root = Get-Item $Path
    $folderList = New-Object System.Collections.Generic.List[PSObject]
    $folderList.Add($root)

    $allDirs = Get-ChildItem -Path $Path -Directory -Recurse -Depth ($Depth - 1) -ErrorAction SilentlyContinue
    
    foreach ($dir in $allDirs) {
        $currentAcl = Get-Acl $dir.FullName
        $parentAcl = Get-Acl (Split-Path $dir.FullName -Parent)
        # Compare only DACL strings to avoid noise from Owner changes
        if ($currentAcl.AccessToString -ne $parentAcl.AccessToString) {
            $folderList.Add($dir)
        }
    }

    # 2. DATA COLLECTION
    $userMap = @{} 
    $rightsMap = @{} 

    Write-Host ">>> Step 2: Processing users and AD names..." -ForegroundColor Cyan
    foreach ($f in $folderList) {
        $acl = Get-Acl $f.FullName
        foreach ($acc in $acl.Access) {
            $identity = $acc.IdentityReference.Value
            if ($identity -match "SYSTEM|NT AUTHORITY|BUILTIN") { continue }
            $rawName = if ($identity -match "\\") { $identity.Split('\')[-1] } else { $identity }
            $rights = if ($acc.FileSystemRights.ToString() -match "FullControl|Modify|Write") { "W" } else { "R" }

            $members = Get-Members $rawName
            foreach ($m in $members) {
                if ($null -eq $m) { continue }
                if (!$userMap.ContainsKey($m)) {
                    $dName = $m
                    if ($adAvailable) {
                        try {
                            $user = Get-ADUser -Identity $m -Properties DisplayName -ErrorAction SilentlyContinue
                            if ($user.DisplayName) { $dName = $user.DisplayName }
                        } catch {}
                    }
                    $userMap[$m] = [PSCustomObject]@{ FullName = $dName; Login = $m }
                }
                $key = "$m|$($f.FullName)"
                if ($rightsMap[$key] -ne "W") { $rightsMap[$key] = $rights }
            }
        }
    }

    # 3. EXCEL CONSTRUCTION
    $excel = New-Object OfficeOpenXml.ExcelPackage
    $ws = $excel.Workbook.Worksheets.Add("Permissions")
    $sortedLogins = $userMap.Keys | Sort-Object
    $baseLen = $root.Parent.FullName.Length
    $maxH = ($folderList | ForEach-Object { ($_.FullName.Substring($baseLen).TrimStart('\').Split('\')).Count } | Measure-Object -Maximum).Maximum
    if ($maxH -lt 2) { $maxH = 2 }

    for ($c = 0; $c -lt $folderList.Count; $c++) {
        $col = $c + 3
        $fPath = $folderList[$c].FullName
        $parts = $fPath.Substring($baseLen).TrimStart('\').Split('\')
        for ($i = 0; $i -lt $parts.Count; $i++) { $ws.SetValue(($i + 1), $col, $parts[$i]) }

        for ($u = 0; $u -lt $sortedLogins.Count; $u++) {
            $row = $maxH + 1 + $u
            $login = $sortedLogins[$u]
            if ($c -eq 0) {
                $ws.SetValue($row, 1, $userMap[$login].FullName)
                $ws.SetValue($row, 2, $userMap[$login].Login)
            }
            $val = if ($rightsMap["$login|$fPath"]) { $rightsMap["$login|$fPath"] } else { "-" }
            $ws.SetValue($row, $col, $val)
            
            $cell = $ws.Cells[$row, $col]
            $cell.Style.Fill.PatternType = 'Solid'
            if ($val -eq "W") { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen) }
            elseif ($val -eq "R") { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSkyBlue) }
            else { 
                $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 199, 206))
                $cell.Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
            }
        }
    }

    # 4. FORMATTING & ALIGNMENT
    $totalR = $maxH + $sortedLogins.Count
    $totalC = $folderList.Count + 2
    $allRange = $ws.Cells[1, 1, $totalR, $totalC]

    # Global Borders and LEFT alignment
    $allRange.Style.Border.Top.Style = $allRange.Style.Border.Bottom.Style = 'Thin'
    $allRange.Style.Border.Left.Style = $allRange.Style.Border.Right.Style = 'Thin'
    $allRange.Style.HorizontalAlignment = 'Left'
    $allRange.Style.VerticalAlignment = 'Center'
    $allRange.Style.Indent = 1

    # User Headers Style
    $userHdr = $ws.Cells[1, 1, $maxH, 2]
    $userHdr.Style.Fill.PatternType = 'Solid'
    $userHdr.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::SlateGray)
    $userHdr.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    
    $ws.Cells[1, 1, $maxH, 1].Merge = $true; $ws.Cells[1, 1].Value = "Full Name"
    $ws.Cells[1, 2, $maxH, 2].Merge = $true; $ws.Cells[1, 2].Value = "Login"
    $ws.Cells[1, 3, $maxH, 3].Merge = $true

    # Folder Header Merging
    for ($r = 1; $r -le $maxH; $r++) {
        for ($c = 4; $c -le ($folderList.Count + 2); $c++) {
            $v = $ws.GetValue($r, $c); if (!$v) { continue }
            $s = $c
            while ($c + 1 -le ($folderList.Count + 2) -and $ws.GetValue($r, $c + 1) -eq $v) { $c++ }
            if ($c -gt $s) { $ws.Cells[$r, $s, $r, $c].Merge = $true }
        }
    }

    # Folder Headers Style
    $folderHdrArea = $ws.Cells[1, 3, $maxH, $totalC]
    $folderHdrArea.Style.Fill.PatternType = 'Solid'
    $folderHdrArea.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
    $ws.Cells[1, 1, $maxH, $totalC].Style.Font.Bold = $true

    $ws.Cells.AutoFitColumns()

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
    $excel.SaveAs($OutputFile)
    $excel.Dispose()
    Write-Host "`n>>> Success! Report generated in English. Path: $OutputFile" -ForegroundColor Green
}