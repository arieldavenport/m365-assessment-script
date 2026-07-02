# m365-assessment-script

One-shot Microsoft 365 tenant assessment. Paste a single line into the
Microsoft 365 admin center **Cloud Shell (Bash)** and walk away with a
zip containing a full tenant inventory *and* a security review: Conditional
Access, admin roles, MFA registration, guest access, legacy auth, OAuth app
consent risk, Identity Protection risk signals, and Secure Score.

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
| `M365_ConditionalAccessPolicies_<tenant>_<ts>.csv` | Every Conditional Access policy: state, scope (users/apps/client app types), grant/session controls. |
| `M365_AdminRoleAssignments_<tenant>_<ts>.csv` | Active Microsoft Entra directory role assignments, flagged by whether the role is high-privilege. |
| `M365_MFARegistration_<tenant>_<ts>.csv` | Per-user authentication method registration: MFA/SSPR/passwordless status, `IsAdmin`, methods registered. |
| `M365_GuestUsers_<tenant>_<ts>.csv`   | Guest (B2B) accounts, same columns as the Users CSV. |
| `M365_LegacyAuthSignIns_<tenant>_<ts>.csv` | Sign-ins using legacy/basic auth protocols (IMAP, POP, ActiveSync, etc.) in the lookback window, aggregated by user/app/client. |
| `M365_OAuthConsent_<tenant>_<ts>.csv` | Every delegated and application permission grant to an enterprise app, with high-risk permissions flagged. |
| `M365_RiskyUsers_<tenant>_<ts>.csv`   | Microsoft Entra ID Protection risky users (requires Entra ID P2). |
| `M365_SecureScore_<tenant>_<ts>.csv`  | Microsoft Secure Score, per control, worst-implemented first. |
| `M365_SecurityFindings_<tenant>_<ts>.csv` | Consolidated High/Medium/Low findings synthesized across all of the above, sorted by severity. |

## Options

| Flag | Default | Notes |
| --- | --- | --- |
| `--stale-days N`         | `90`        | Inactivity threshold for the stale-users CSV. |
| `--output-dir PATH`      | `$HOME`     | Where the CSVs / zip are written. |
| `--skip-friendly-names`  | off         | Skip downloading Microsoft's SKU friendly-name map; CSV keeps raw `SkuPartNumber` values. |
| `--skip-security-review` | off         | Skip the security review entirely and only produce the original Users/Products/StaleUsers CSVs. |
| `--legacy-auth-days N`   | `7`         | Sign-in log lookback window for the legacy-auth scan. |
| `--no-download`          | off         | Don't auto-trigger the Cloud Shell download dialog; just print the zip path. |

## Requirements

- Microsoft 365 admin center / Azure Cloud Shell **Bash** mode (already has `az`, `curl`, `jq`, `zip`, `python3`).
- An account with at least **Global Reader** on the tenant.
- The Azure CLI session in Cloud Shell already covers Graph delegated permissions via `Directory.AccessAsUser.All` — enough for users, products, and most of the security review (Conditional Access, admin roles, MFA registration, guests, legacy auth, OAuth consent, Secure Score).
- `signInActivity` on user rows, MFA registration details, and the legacy-auth scan require `AuditLog.Read.All` consent on the Azure CLI app. If it's not granted, those exports still produce successfully but the relevant columns/CSVs are blank, and the stale list falls back to a created-date heuristic.
- Risky users (`M365_RiskyUsers_*.csv`) require an **Entra ID P2** license — on P1/free tenants the CSV is written with headers only.
- Every security-review sub-check degrades gracefully: if the signed-in account lacks the Graph permission for it, the script prints a warning, writes a header-only CSV, and continues rather than aborting.

## Running locally

If you'd rather not pipe from `curl`:

```bash
wget https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/m365-assessment.sh
chmod +x m365-assessment.sh
./m365-assessment.sh --stale-days 90
```
