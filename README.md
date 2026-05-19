# m365-assessment-script

One-shot Microsoft 365 tenant assessment. Paste it into the M365 admin
center Cloud Shell (or any PowerShell 7+ host with `Microsoft.Graph`
installed) and walk away with three CSVs.

## Quick start (one line)

Open the **Cloud Shell** in the Microsoft 365 admin center (PowerShell mode),
paste this, hit enter:

```powershell
irm https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/Invoke-M365Assessment.ps1 | iex
```

Consent to the Graph scopes when prompted. Three CSVs land in the current
directory; download them from the Cloud Shell file browser.

To pass parameters (e.g. a different stale-user cutoff), wrap it:

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/Invoke-M365Assessment.ps1))) -StaleDays 60
```

## What you get

| File | What's in it |
| --- | --- |
| `M365_Users_<tenant>_<ts>.csv`        | Mirrors the admin center "Active users" export plus sign-in activity (last interactive / non-interactive / successful sign-in, days since last activity). |
| `M365_Products_<tenant>_<ts>.csv`     | Subscribed SKUs across all billing accounts the tenant can see (total / consumed / available licenses, service plans, friendly product names). |
| `M365_StaleUsers_<tenant>_<ts>.csv`   | Enabled, non-guest accounts whose most recent sign-in is older than `-StaleDays` (default 90), or that have never signed in and were created longer ago than the threshold. |

## Alternate usage

If you'd rather upload the file into Cloud Shell and run it locally:

```powershell
./Invoke-M365Assessment.ps1
```

A browser / device-code prompt asks you to consent to these Graph scopes
the first time:

- `User.Read.All`
- `Organization.Read.All`
- `Directory.Read.All`
- `AuditLog.Read.All`

CSVs land in the current directory; download them from the Cloud Shell
file browser.

### Parameters

| Parameter | Default | Notes |
| --- | --- | --- |
| `-OutputPath`         | current dir | Where the CSVs are written. |
| `-StaleDays`          | `90`        | Inactivity threshold for the stale-users CSV. |
| `-SkipFriendlyNames`  | off         | Skip the one-time download of Microsoft's SKU friendly-name map; CSV keeps raw `SkuPartNumber` values. |

## Requirements

- PowerShell 7+ (Cloud Shell already has it).
- `Microsoft.Graph` modules (auto-installed to the current user if missing).
- An account with at least **Global Reader** on the tenant.
- **Microsoft Entra ID P1 or P2** for the `signInActivity` property. Without
  it the script still produces the user and product CSVs, but the stale-user
  list falls back to a created-date heuristic and prints a warning.
