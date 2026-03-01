[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [string[]]$TargetPaths,
    [Parameter(Mandatory=$false)]
    [int]$GlobalDepth = 100,
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "$PSScriptRoot\PermissionsReport.xlsx"
)

process {
    if (!(Get-Module -ListAvailable ImportExcel)) { Write-Error "ImportExcel module missing."; return }
    Import-Module ImportExcel
    Add-Type -AssemblyName System.Drawing

    $adAvailable = Get-Module -ListAvailable ActiveDirectory
    if ($adAvailable) { try { Import-Module ActiveDirectory -ErrorAction SilentlyContinue } catch { $adAvailable = $false } }
    $groupCache = @{}

    function Get-Members ($identityName) {
        if (!$adAvailable) { return @($identityName) }
        if ($groupCache.ContainsKey($identityName)) { return $groupCache[$identityName] }
        try {
            $obj = Get-ADObject -Filter "SamAccountName -eq '$identityName'" -ErrorAction SilentlyContinue
            if ($null -ne $obj -and $obj.ObjectClass -eq "group") {
                $m = Get-ADGroupMember -Identity $identityName -Recursive | Select-Object -ExpandProperty SamAccountName
                $groupCache[$identityName] = $m; return $m
            }
        } catch {}
        return @($identityName)
    }

    $rawFolderList = New-Object System.Collections.Generic.List[PSObject]
    $rootMap = @{}

    foreach ($entry in $TargetPaths) {
        $path = $entry; $depth = $GlobalDepth
        if ($entry -match "(.+):(\d+)$") { $path = $Matches[1]; $depth = [int]$Matches[2] }
        if (!(Test-Path $path)) { continue }
        
        $rootItem = Get-Item $path
        $rawFolderList.Add($rootItem)
        $rootMap[$rootItem.FullName] = $rootItem.FullName

        $subDirs = Get-ChildItem -Path $path -Directory -Recurse -Depth ($depth - 1) -ErrorAction SilentlyContinue
        foreach ($dir in $subDirs) {
            try {
                $cACL = Get-Acl $dir.FullName; $pACL = Get-Acl (Split-Path $dir.FullName -Parent)
                $f1 = ($cACL.Access | Where-Object {$_.IdentityReference -notmatch "SYSTEM|NT AUTHORITY|BUILTIN"} | ForEach-Object {"$($_.IdentityReference.Value):$($_.FileSystemRights)"} | Sort-Object) -join ";"
                $f2 = ($pACL.Access | Where-Object {$_.IdentityReference -notmatch "SYSTEM|NT AUTHORITY|BUILTIN"} | ForEach-Object {"$($_.IdentityReference.Value):$($_.FileSystemRights)"} | Sort-Object) -join ";"
                if ($f1 -ne $f2) { $rawFolderList.Add($dir); $rootMap[$dir.FullName] = $rootItem.FullName }
            } catch {}
        }
    }

    $allFolderList = $rawFolderList | Sort-Object FullName
    $userMap = @{} ; $rightsMap = @{} 
    foreach ($f in $allFolderList) {
        try {
            $acl = Get-Acl $f.FullName
            foreach ($acc in $acl.Access) {
                $identity = $acc.IdentityReference.Value
                if ($identity -match "SYSTEM|NT AUTHORITY|BUILTIN") { continue }
                $raw = $identity.Split('\')[-1]
                $rights = if ($acc.FileSystemRights.ToString() -match "FullControl|Modify|Write") { "W" } else { "R" }
                foreach ($m in (Get-Members $raw)) {
                    if (!$userMap.ContainsKey($m)) {
                        $dn = $m; if ($adAvailable) { try { $u = Get-ADUser $m; $dn = $u.Name } catch {} }
                        $userMap[$m] = [PSCustomObject]@{FullName=$dn; Login=$m}
                    }
                    $key = "$m|$($f.FullName)"
                    if ($rightsMap[$key] -ne "W") { $rightsMap[$key] = $rights }
                }
            }
        } catch {}
    }

    $excel = New-Object OfficeOpenXml.ExcelPackage
    $ws = $excel.Workbook.Worksheets.Add("Permissions")
    $logins = $userMap.Keys | Sort-Object
    
    $maxH = 1
    foreach ($f in $allFolderList) {
        $rel = $f.FullName.Replace($rootMap[$f.FullName], "").TrimStart('\')
        $d = if ($rel -eq "") { 1 } else { ($rel.Split('\')).Count + 1 }
        if ($d -gt $maxH) { $maxH = $d }
    }

    for ($c = 0; $c -lt $allFolderList.Count; $c++) {
        $col = $c + 3; $f = $allFolderList[$c]; $root = $rootMap[$f.FullName]
        $ws.SetValue(1, $col, $root)
        $rel = $f.FullName.Replace($root, "").TrimStart('\')
        if ($rel -ne "") {
            $parts = $rel.Split('\')
            for ($p = 0; $p -lt $parts.Count; $p++) { $ws.SetValue(($p + 2), $col, $parts[$p]) }
        }
        for ($u = 0; $u -lt $logins.Count; $u++) {
            $row = $maxH + 1 + $u; $login = $logins[$u]
            if ($c -eq 0) { 
                $ws.SetValue($row, 1, $userMap[$login].FullName); $ws.SetValue($row, 2, $userMap[$login].Login)
            }
            $v = if ($rightsMap["$login|$($f.FullName)"]) { $rightsMap["$login|$($f.FullName)"] } else { "-" }
            $ws.SetValue($row, $col, $v)
            $cell = $ws.Cells[$row, $col]; $cell.Style.Fill.PatternType = 'Solid'
            switch ($v) {
                "W" { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen) }
                "R" { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSkyBlue) }
                default { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 199, 206)) }
            }
            $cell.Style.HorizontalAlignment = 'Center'
        }
    }

    $totalC = $allFolderList.Count + 2

    # --- FINAL HYBRID MERGE LOGIC ---

    # 1. Горизонтальное объединение (создаем блоки)
    for ($r = 1; $r -le $maxH; $r++) {
        for ($c = 3; $c -lt $totalC; $c++) {
            $v1 = $ws.GetValue($r, $c)
            if ($null -eq $v1) { continue }
            $startCol = $c
            while ($c + 1 -le $totalC -and $ws.GetValue($r, $c + 1) -eq $v1) {
                if ($r -gt 1 -and $ws.GetValue($r - 1, $startCol) -ne $ws.GetValue($r - 1, $c + 1)) { break }
                $c++
            }
            if ($c -gt $startCol) { try { $ws.Cells[$r, $startCol, $r, $c].Merge = $true } catch {} }
        }
    }

    # 2. Вертикальное объединение (растягиваем блоки вниз)
    # Сначала Full Name / Login
    $ws.Cells[1,1,$maxH,1].Merge = $true; $ws.Cells[1,1].Value = "Full Name"
    $ws.Cells[1,2,$maxH,2].Merge = $true; $ws.Cells[1,2].Value = "Login"

    # Теперь папки
    for ($col = 3; $col -le $totalC; $col++) {
        for ($row = 1; $row -le $maxH; $row++) {
            $val = $ws.GetValue($row, $col)
            if ($null -ne $val) {
                $startR = $row
                $nextR = $row + 1
                # Ищем пустые ячейки строго под этой
                while ($nextR -le $maxH -and $null -eq $ws.GetValue($nextR, $col)) {
                    $nextR++
                }
                if ($nextR - 1 -gt $startR) {
                    # Пытаемся объединить. Если ячейка уже в горизонтальном мерже, 
                    # EPPlus корректно расширит область.
                    try { $ws.Cells[$startR, $col, $nextR - 1, $col].Merge = $true } catch {}
                }
                $row = $nextR - 1
            }
        }
    }

    # Styles
    $fullRange = $ws.Cells[1, 1, ($maxH + $logins.Count), $totalC]
    $fullRange.Style.VerticalAlignment = 'Center'
    $fullRange.Style.Border.Top.Style = $fullRange.Style.Border.Bottom.Style = $fullRange.Style.Border.Left.Style = $fullRange.Style.Border.Right.Style = 'Thin'
    $ws.Cells[1, 1, $maxH, $totalC].Style.Fill.PatternType = 'Solid'
    $ws.Cells[1, 1, $maxH, $totalC].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGray)
    $ws.Cells[1, 1, $maxH, $totalC].Style.Font.Bold = $true
    $ws.Cells[1, 1, $maxH, $totalC].Style.HorizontalAlignment = 'Left'
    $ws.Cells[($maxH + 1), 1, ($maxH + $logins.Count), 2].Style.HorizontalAlignment = 'Left'
    
    $ws.Cells.AutoFitColumns(12, 60)

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
    $excel.SaveAs($OutputFile); $excel.Dispose()
    Write-Host "`n>>> Done! Hybrid Merge (H then V) applied." -ForegroundColor Green
}