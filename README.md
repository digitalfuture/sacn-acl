# NTFSPermissionsReport

A PowerShell tool for generating a consolidated Excel matrix of NTFS permissions across multiple root directories with Active Directory group expansion.

## Key Features

- **Multi-Root Support**: Scan multiple paths in one run.
- **Smart Depth Control**: Set a global depth or override it for specific paths using `Path:Depth` syntax.
- **DACL Fingerprinting**: Automatically hides folders that inherit permissions without changes (ignores Owner/Technical noise).
- **AD Group Expansion**: Recursively expands AD groups to show individual user access.
- **Excel Matrix**: Color-coded output (Green=Write, Blue=Read, Red=No Access) with left-aligned formatting.

## Prerequisites

1. **PowerShell 5.1+**
2. **ImportExcel Module**: `Install-Module ImportExcel`
3. **RSAT (Active Directory)**: Required for group expansion. If missing, the script will fallback to raw identity names.

## Usage

### Basic Command

Run the script and provide target paths. You can specify depth for each path individually.

```powershell
.\Scan-ACL.ps1 -TargetPaths "C:\Projects", "D:\Archive:2", "\\Server\Share:1"
```

## Parameters

- **TargetPaths**: (Required) Array of strings. Syntax: Path or Path:Depth.

- **GlobalDepth**: (Optional) Default depth for paths without a specific limit. Default is 100.

- **OutputFile**: (Optional) Custom path for the resulting .xlsx file.

## Examples

Global depth with overrides:

```powershell
.\Scan-ACL.ps1 -TargetPaths "C:\Data:1", "D:\Backups" -GlobalDepth 5
```

### Output Legend

- **W (Light Green)**: Full Control, Modify, or Write permissions.

- **R (Light Blue)**: Read-only permissions.

- **(Light Red)**: No permissions detected for the user on this specific folder.
