[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$Path,

    [Parameter(Mandatory=$false)]
    [int]$Depth = -1,

    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "$PSScriptRoot\PermissionsReport.csv"
)

Process {
    $dirParams = @{
        Path      = $Path
        Directory = $true
        Recurse   = $true
        ErrorAction = 'SilentlyContinue'
    }
    
    if ($Depth -ge 0) { $dirParams["Depth"] = $Depth }

    Write-Host "Scanning folders... please wait." -ForegroundColor Cyan
    $rootFolder = Get-Item $Path
    $subFolders = Get-ChildItem @dirParams
    $allFolders = @($rootFolder) + $subFolders

    $allUsers = $allFolders | ForEach-Object { 
        (Get-Acl $_.FullName).Access.IdentityReference 
    } | Select-Object -Unique | Where-Object { 
        $_ -notmatch "NT AUTHORITY|BUILTIN|SYSTEM" 
    }

    $results = New-Object System.Collections.Generic.List[PSObject]

    foreach ($user in $allUsers) {
        $userRow = [ordered]@{ "User" = $user }
        
        foreach ($folder in $allFolders) {
            $relativeName = $folder.FullName.Replace($Path, $rootFolder.Name).Replace("\", " - ")
            
            $acl = Get-Acl $folder.FullName
            $userPerms = $acl.Access | Where-Object { $_.IdentityReference -eq $user }
            
            if ($userPerms) {
                $rights = $userPerms.FileSystemRights -join ","
                if ($rights -match "FullControl|Modify|Write") {
                    $accessType = "Write"
                } elseif ($rights -match "Read|ListDirectory") {
                    $accessType = "Read"
                } else {
                    $accessType = "Special"
                }
            } else {
                $accessType = "-"
            }
            $userRow[$relativeName] = $accessType
        }
        $results.Add([PSCustomObject]$userRow)
    }

    # Ensure output directory exists (especially if user provided a custom path)
    $outputDir = Split-Path $OutputFile
    if ($outputDir -and !(Test-Path $outputDir)) { 
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null 
    }

    $results | Export-Csv -Path $OutputFile -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host "Process Finished!" -ForegroundColor Green
    Write-Host "Folders found: $($allFolders.Count)"
    Write-Host "Report saved: $OutputFile" -ForegroundColor Yellow
}