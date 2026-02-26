# Folder Permissions Analyzer

This PowerShell script scans a directory and its subfolders to create a permissions matrix. It exports the result to a CSV file.

## Features

- **Mandatory Path**: Requires a target directory for analysis.
- **Auto-Output**: By default, saves the report in the same folder as the script.
- **Custom Depth**: Option to limit how deep the script should scan.
- **Flattened Headers**: Folder hierarchy is displayed as `Root - Folder - SubFolder`.

## Parameters

- `-Path`: (Mandatory) The full path to the folder to analyze.
- `-Depth`: (Optional) How many levels of subfolders to include. Default is `-1` (full recursion).
- `-OutputFile`: (Optional) Custom path for the CSV. Default is `PermissionsReport.csv` in the script's directory.

---

## Examples

### 1. Default Scan

Saves `PermissionsReport.csv` in the current script folder:

```powershell
.\Scan-ACL.ps1 -Path "C:\Data\Shared"
```
"# sacn-acl" 
