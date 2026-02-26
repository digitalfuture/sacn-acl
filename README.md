# Folder Permissions Matrix (AD Expanded)

This script generates a professional Excel report of folder permissions. It expands Active Directory groups to show individual users and their access levels.

## Features

- **AD Group Expansion**: Shows real people, not just group names.
- **Excel Output**: Creates `.xlsx` files with filters and frozen headers.
- **Path Hierarchy**: Uses `\` in headers for a natural folder structure view.
- **Write/Read Logic**: Simplifies complex NTFS rights.

## Prerequisites

Before running the script, you must install the following dependencies:

### 1. Active Directory Module (RSAT)

Required to expand groups.

- **Windows 10/11**: Settings > Apps > Optional features > Add a feature > "RSAT: Active Directory Domain Services".
- **Windows Server**: `Install-WindowsFeature RSAT-AD-PowerShell`.

### 2. ImportExcel Module

Required to generate Excel files without Microsoft Office.
Run this in PowerShell as Administrator:

```powershell
Install-Module ImportExcel -Scope CurrentUser
```
