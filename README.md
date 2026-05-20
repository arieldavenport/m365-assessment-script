# m365-assessment-script

One-shot Microsoft 365 tenant assessment. Paste a single line into the
Microsoft 365 admin center **Cloud Shell (Bash)** and walk away with a
zip containing three CSVs.

## Quick start (one line)

Open Cloud Shell in **Bash** mode (top-left "Switch to Bash" button if it's in
PowerShell), then paste:

```bash
curl -s https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/m365-assessment.sh | bash
```

Cloud Shell pops the file-download dialog at the end with the zip
pre-selected. No second sign-in (reuses your existing Azure CLI session),
no module installs, no Conditional Access friction.

Need different options? Pass arguments through:

```bash
curl -s https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/m365-assessment.sh | bash -s -- --stale-days 60
```

## What you get

| File in zip | What's in it |
| --- | --- |
| `M365_Users_<tenant>_<ts>.csv`        | Mirrors the admin center "Active users" export plus sign-in activity (last interactive / non-interactive / successful sign-in, days since last activity). |
| `M365_Products_<tenant>_<ts>.csv`     | Subscribed SKUs across all billing accounts the tenant can see (total / consumed / available licenses, service plans, friendly product names). |
| `M365_StaleUsers_<tenant>_<ts>.csv`   | Enabled, non-guest accounts whose most recent sign-in is older than `--stale-days` (default 90), or that have never signed in and were created longer ago than the threshold. |

## Options

| Flag | Default | Notes |
| --- | --- | --- |
| `--stale-days N`         | `90`        | Inactivity threshold for the stale-users CSV. |
| `--output-dir PATH`      | `$HOME`     | Where the CSVs / zip are written. |
| `--skip-friendly-names`  | off         | Skip downloading Microsoft's SKU friendly-name map; CSV keeps raw `SkuPartNumber` values. |
| `--no-download`          | off         | Don't auto-trigger the Cloud Shell download dialog; just print the zip path. |

## Requirements

- Microsoft 365 admin center / Azure Cloud Shell **Bash** mode (already has `az`, `curl`, `jq`, `zip`, `python3`).
- An account with at least **Global Reader** on the tenant.
- The Azure CLI session in Cloud Shell already covers Graph delegated permissions via `Directory.AccessAsUser.All` — enough for users + products.
- `signInActivity` on user rows requires `AuditLog.Read.All` consent on the Azure CLI app. If it's not granted, the user export still produces successfully but the sign-in columns are blank and the stale list falls back to a created-date heuristic.

## Running locally

If you'd rather not pipe from `curl`:

```bash
wget https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/m365-assessment.sh
chmod +x m365-assessment.sh
./m365-assessment.sh --stale-days 90
```
