# -arieldavenport-m365-assessment-script

Exports users with assigned licenses and tenant product/license data to CSV files in one run.

## Requirements

- PowerShell 7+
- Microsoft Graph PowerShell SDK (`Install-Module Microsoft.Graph -Scope CurrentUser`)

## Usage

```powershell
pwsh ./export-users-and-licenses.ps1 -OutputDirectory ./exports
```

This creates:

- `users-with-assigned-licenses.csv`
- `tenant-products-licenses.csv`
