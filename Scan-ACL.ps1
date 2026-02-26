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

Process {
    # Проверка ImportExcel (обязателен для файла)
    if (!(Get-Module -ListAvailable ImportExcel)) { 
        Write-Error "Module 'ImportExcel' is required. Run: Install-Module ImportExcel -Scope CurrentUser"
        return 
    }
    Import-Module ImportExcel
    Add-Type -AssemblyName System.Drawing

    # Проверка AD (опционально)
    $ADAvailable = $false
    if (Get-Module -ListAvailable ActiveDirectory) {
        Import-Module ActiveDirectory
        $ADAvailable = $true
        Write-Host "Active Directory module found. Groups will be expanded." -ForegroundColor Green
    } else {
        Write-Host "Active Directory module NOT found. Showing raw identities (no group expansion)." -ForegroundColor Yellow
    }

    Write-Host "Step 1: Scanning folders..." -ForegroundColor Cyan
    $dirParams = @{ Path = $Path; Directory = $true; Recurse = $true; ErrorAction = 'SilentlyContinue' }
    if ($Depth -ge 0) { $dirParams["Depth"] = $Depth }
    $rootFolder = Get-Item $Path
    $allFolders = @($rootFolder) + (Get-ChildItem @dirParams | Sort-Object FullName)

    # Высота шапки
    $maxLevel = 0
    foreach ($f in $allFolders) {
        $level = ($f.FullName.Replace($Path, "").Split('\', [System.StringSplitOptions]::RemoveEmptyEntries)).Count
        if ($level -gt $maxLevel) { $maxLevel = $level }
    }
    $headerRowsCount = [int]$maxLevel + 1

    Write-Host "Step 2: Processing permissions..." -ForegroundColor Cyan
    $userPermissionsMaster = @{} 
    $userFullNameCache = @{}
    $groupCache = @{}

    foreach ($folder in $allFolders) {
        $acl = Get-Acl $folder.FullName
        $folderPathKey = $folder.FullName

        foreach ($access in $acl.Access) {
            $identity = $access.IdentityReference.Value
            if ($identity -match "NT AUTHORITY|BUILTIN|SYSTEM") { continue }
            
            # Чистим имя от домена
            $cleanIdentity = if ($identity -match '\\') { $identity.Split('\')[-1] } else { $identity }

            $rights = $access.FileSystemRights.ToString()
            $accessType = if ($rights -match "FullControl|Modify|Write") { "W" } else { "R" }

            $uLogins = @()
            if ($ADAvailable) {
                if ($groupCache.ContainsKey($cleanIdentity.ToLower())) {
                    $uLogins = $groupCache[$cleanIdentity.ToLower()]
                } else {
                    try {
                        $adObj = Get-ADObject -Filter "sAMAccountName -eq '$cleanIdentity'" -Properties DisplayName, SamAccountName, ObjectClass -ErrorAction SilentlyContinue
                        if ($adObj -and $adObj.ObjectClass -eq "group") {
                            $members = Get-ADGroupMember -Identity $adObj.distinguishedName -Recursive | Where-Object { $_.objectClass -eq "user" }
                            $uLogins = $members.SamAccountName
                            foreach ($m in $members) {
                                $mKey = $m.SamAccountName.ToLower()
                                if (!$userFullNameCache.ContainsKey($mKey)) { $userFullNameCache[$mKey] = $m.name }
                            }
                        } elseif ($adObj) {
                            $adSamName = $adObj.SamAccountName
                            $uLogins = @($adSamName)
                            if ($adObj.DisplayName) { $userFullNameCache[$adSamName.ToLower()] = $adObj.DisplayName }
                        } else {
                            $uLogins = @($cleanIdentity)
                        }
                    } catch { $uLogins = @($cleanIdentity) }
                    $groupCache[$cleanIdentity.ToLower()] = $uLogins
                }
            } else {
                # Если AD нет, просто используем имя из ACL
                $uLogins = @($cleanIdentity)
            }

            foreach ($u in $uLogins) {
                $uKey = $u.ToLower().Trim()
                if (!$userPermissionsMaster.ContainsKey($uKey)) { $userPermissionsMaster[$uKey] = @{} }
                if ($userPermissionsMaster[$uKey][$folderPathKey] -ne "W") { $userPermissionsMaster[$uKey][$folderPathKey] = $accessType }
            }
        }
    }

    Write-Host "Step 3: Building Excel..." -ForegroundColor Cyan
    $excel = New-Object OfficeOpenXml.ExcelPackage
    $sheet = $excel.Workbook.Worksheets.Add("Permissions")

    $currentCol = 3
    $sheet.Cells[[int]1, [int]1].Value = "Login"
    $sheet.Cells[[int]1, [int]2].Value = "Full Name"
    $sheet.Cells[[int]1, [int]1, [int]$headerRowsCount, [int]1].Merge = $true
    $sheet.Cells[[int]1, [int]2, [int]$headerRowsCount, [int]2].Merge = $true

    $sortedUserKeys = $userPermissionsMaster.Keys | Sort-Object

    foreach ($folder in $allFolders) {
        $relativeDir = $folder.FullName.Replace($Path, $rootFolder.Name)
        $parts = $relativeDir.Split('\', [System.StringSplitOptions]::RemoveEmptyEntries)
        
        for ($i = 0; $i -lt $parts.Count; $i++) {
            $sheet.Cells[[int]($i + 1), [int]$currentCol].Value = $parts[$i]
        }
        
        $currentRow = [int]$headerRowsCount + 1
        foreach ($uKey in $sortedUserKeys) {
            if ($currentCol -eq 3) { 
                $sheet.Cells[$currentRow, [int]1].Value = $uKey
                $dispName = if ($userFullNameCache.ContainsKey($uKey)) { $userFullNameCache[$uKey] } else { $uKey }
                $sheet.Cells[$currentRow, [int]2].Value = $dispName
            }
            $val = if ($userPermissionsMaster[$uKey][$folder.FullName]) { $userPermissionsMaster[$uKey][$folder.FullName] } else { "-" }
            $cell = $sheet.Cells[$currentRow, [int]$currentCol]
            $cell.Value = $val
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            
            if ($val -eq "W") { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen) }
            elseif ($val -eq "R") { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSkyBlue) }
            elseif ($val -eq "-") { $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightPink) }
            $currentRow++
        }
        $currentCol++
    }

    # Merging
    for ($r = 1; $r -le $headerRowsCount; $r++) {
        $startCol = 3
        for ($c = 3; $c -le ($currentCol - 1); $c++) {
            $cur = $sheet.Cells[$r, $c].Value
            $nxt = $sheet.Cells[$r, $c + 1].Value
            if ($cur -ne $nxt -or $c -eq ($currentCol - 1)) {
                if ($c -gt $startCol) { $sheet.Cells[$r, $startCol, $r, $c].Merge = $true }
                $startCol = $c + 1
            }
        }
    }

    $allRange = $sheet.Cells[[int]1, [int]1, [int]($currentRow - 1), [int]($currentCol - 1)]
    $allRange.Style.Border.Top.Style = $allRange.Style.Border.Left.Style = $allRange.Style.Border.Right.Style = $allRange.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    
    $headerRange = $sheet.Cells[[int]1, [int]1, [int]$headerRowsCount, [int]($currentCol - 1)]
    $headerRange.Style.Font.Bold = $true
    $headerRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $headerRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

    $sheet.Cells.AutoFitColumns()

    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }
    $excel.SaveAs($OutputFile)
    $excel.Dispose()

    Write-Host "Success! Report created: $OutputFile" -ForegroundColor Green
}