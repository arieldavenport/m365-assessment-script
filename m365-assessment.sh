#!/usr/bin/env bash
# m365-assessment.sh
# One-shot Microsoft 365 tenant assessment: users, products, and stale accounts.
#
# Designed to run in the Microsoft 365 admin center / Azure Cloud Shell (Bash).
# Reuses the existing Azure CLI session for Graph auth (no device code, no
# Conditional Access friction), writes three CSVs, zips them, and triggers
# the Cloud Shell download dialog.
#
# Usage:
#   ./m365-assessment.sh [--stale-days N] [--output-dir PATH] [--skip-friendly-names] [--no-download]
#
# One-liner (paste into Cloud Shell, Bash mode):
#   curl -s https://raw.githubusercontent.com/arieldavenport/m365-assessment-script/main/m365-assessment.sh | bash

set -euo pipefail

STALE_DAYS=90
OUTPUT_DIR="$HOME"
SKIP_FRIENDLY=0
NO_DOWNLOAD=0
SKIP_SECURITY=0
LEGACY_AUTH_DAYS=7

usage() {
    cat <<EOF
Usage: $0 [options]
  --stale-days N           Stale account threshold in days (default 90)
  --output-dir PATH        Output directory (default: \$HOME)
  --skip-friendly-names    Skip downloading Microsoft's SKU friendly-name map
  --skip-security-review   Skip the security review checks (CA, MFA, roles, OAuth apps, etc.)
  --legacy-auth-days N     Sign-in log lookback window for legacy-auth detection (default 7)
  --no-download            Skip the Cloud Shell download dialog at the end
  -h, --help               This help
EOF
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --stale-days)          STALE_DAYS="$2"; shift 2 ;;
        --output-dir)          OUTPUT_DIR="$2"; shift 2 ;;
        --skip-friendly-names) SKIP_FRIENDLY=1; shift ;;
        --skip-security-review) SKIP_SECURITY=1; shift ;;
        --legacy-auth-days)    LEGACY_AUTH_DAYS="$2"; shift 2 ;;
        --no-download)         NO_DOWNLOAD=1; shift ;;
        -h|--help)             usage; exit 0 ;;
        *)                     echo "Unknown argument: $1" >&2; usage; exit 1 ;;
    esac
done

# ---------------------------------------------------------------------------
# Dependencies
# ---------------------------------------------------------------------------
for cmd in az curl jq zip python3; do
    command -v "$cmd" >/dev/null 2>&1 || {
        echo "ERROR: required command not found: $cmd" >&2
        exit 1
    }
done

mkdir -p "$OUTPUT_DIR"

# ---------------------------------------------------------------------------
# Auth: reuse the existing Cloud Shell Azure CLI session
# ---------------------------------------------------------------------------
echo "Getting Microsoft Graph access token from existing Azure CLI session..."
TOKEN=$(az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv)
AUTH="Authorization: Bearer $TOKEN"

# ---------------------------------------------------------------------------
# Org / tenant info
# ---------------------------------------------------------------------------
ORG=$(curl -sf -H "$AUTH" 'https://graph.microsoft.com/v1.0/organization')
TENANT_NAME=$(jq -r '.value[0].displayName' <<<"$ORG")
TENANT_ID=$(jq -r '.value[0].id' <<<"$ORG")
TENANT_TAG=$(echo "$TENANT_NAME" | tr -c 'A-Za-z0-9' '_' | sed 's/__*/_/g; s/^_//; s/_$//')
TIMESTAMP=$(date +%Y%m%d-%H%M%S)

USERS_CSV="$OUTPUT_DIR/M365_Users_${TENANT_TAG}_${TIMESTAMP}.csv"
PRODUCTS_CSV="$OUTPUT_DIR/M365_Products_${TENANT_TAG}_${TIMESTAMP}.csv"
STALE_CSV="$OUTPUT_DIR/M365_StaleUsers_${TENANT_TAG}_${TIMESTAMP}.csv"
CA_CSV="$OUTPUT_DIR/M365_ConditionalAccessPolicies_${TENANT_TAG}_${TIMESTAMP}.csv"
ROLES_CSV="$OUTPUT_DIR/M365_AdminRoleAssignments_${TENANT_TAG}_${TIMESTAMP}.csv"
MFA_CSV="$OUTPUT_DIR/M365_MFARegistration_${TENANT_TAG}_${TIMESTAMP}.csv"
GUESTS_CSV="$OUTPUT_DIR/M365_GuestUsers_${TENANT_TAG}_${TIMESTAMP}.csv"
LEGACYAUTH_CSV="$OUTPUT_DIR/M365_LegacyAuthSignIns_${TENANT_TAG}_${TIMESTAMP}.csv"
OAUTH_CSV="$OUTPUT_DIR/M365_OAuthConsent_${TENANT_TAG}_${TIMESTAMP}.csv"
RISKYUSERS_CSV="$OUTPUT_DIR/M365_RiskyUsers_${TENANT_TAG}_${TIMESTAMP}.csv"
SECURESCORE_CSV="$OUTPUT_DIR/M365_SecureScore_${TENANT_TAG}_${TIMESTAMP}.csv"
FINDINGS_CSV="$OUTPUT_DIR/M365_SecurityFindings_${TENANT_TAG}_${TIMESTAMP}.csv"
ZIP_PATH="$OUTPUT_DIR/M365_Assessment_${TENANT_TAG}_${TIMESTAMP}.zip"

echo "Connected to: $TENANT_NAME  ($TENANT_ID)"

# ---------------------------------------------------------------------------
# SKU friendly-name map (GUID -> Product display name)
# ---------------------------------------------------------------------------
SKU_MAP_JSON='{}'
if [[ $SKIP_FRIENDLY -ne 1 ]]; then
    echo "Downloading Microsoft SKU friendly-name map..."
    SKU_CSV_TMP=$(mktemp)
    SKU_URL='https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    if curl -sf -o "$SKU_CSV_TMP" "$SKU_URL"; then
        SKU_MAP_JSON=$(python3 - "$SKU_CSV_TMP" <<'PY'
import csv, json, sys
m = {}
with open(sys.argv[1], newline='', encoding='utf-8-sig') as f:
    for row in csv.DictReader(f):
        g = (row.get('GUID') or '').strip()
        if g and g not in m:
            m[g] = (row.get('Product_Display_Name') or '').strip()
print(json.dumps(m))
PY
)
        echo "Loaded $(jq 'length' <<<"$SKU_MAP_JSON") SKU GUID -> name mappings."
    else
        echo "WARNING: SKU map download failed; using raw SkuPartNumber." >&2
    fi
    rm -f "$SKU_CSV_TMP"
fi

# ---------------------------------------------------------------------------
# Paginated GET — accumulates .value across @odata.nextLink pages
# ---------------------------------------------------------------------------
get_all_pages() {
    local url="$1"
    local merged='[]'
    while [[ -n "$url" && "$url" != "null" ]]; do
        local page
        page=$(curl -sf -H "$AUTH" "$url") || return 1
        merged=$(jq -c --argjson m "$merged" '$m + .value' <<<"$page") || return 1
        url=$(jq -r '."@odata.nextLink" // empty' <<<"$page")
    done
    echo "$merged"
}

# ---------------------------------------------------------------------------
# Products / subscribed SKUs + commerce subscriptions (for renewal dates)
# ---------------------------------------------------------------------------
NOW_EPOCH=$(date -u +%s)

echo "Collecting subscribed SKUs..."
SKUS=$(get_all_pages 'https://graph.microsoft.com/v1.0/subscribedSkus')

echo "Collecting subscription renewal info (companySubscription)..."
SUBS_LIST='[]'
SUBS_AVAILABLE=0
# Try v1.0 first, fall back to beta. Either may return 403 on tenants where the
# Azure CLI app doesn't have Directory.Read.All consent for subscriptions; in
# that case renewal columns will simply be blank.
for SUBS_URL in 'https://graph.microsoft.com/v1.0/directory/subscriptions' \
                'https://graph.microsoft.com/beta/directory/subscriptions'; do
    if SUBS_TRY=$(get_all_pages "$SUBS_URL" 2>/dev/null); then
        SUBS_LIST="$SUBS_TRY"
        SUBS_AVAILABLE=1
        echo "Loaded $(jq 'length' <<<"$SUBS_LIST") subscriptions from ${SUBS_URL##*/com}."
        break
    fi
done
if [[ $SUBS_AVAILABLE -ne 1 ]]; then
    echo "WARNING: could not read directory/subscriptions; renewal date column will be blank." >&2
fi

# Build combined input and emit one row per subscription, with SKU-level
# consumption joined in. SKUs that have no matching subscription (free /
# derived products) get an "orphan" row with blank renewal fields, so the
# CSV remains a complete inventory.
jq -n -r \
    --argjson skus "$SKUS" \
    --argjson subs "$SUBS_LIST" \
    --argjson map  "$SKU_MAP_JSON" \
    --argjson now  "$NOW_EPOCH" '
    ($skus | map({(.skuId): .}) | add // {}) as $sku_by_id |
    ($subs | map(.skuId)) as $sub_skus |

    # Approximate whole-and-fractional months until the given renewal date.
    # Average month = 365.25/12 days = 2629800 seconds. Negative => overdue.
    def months_until($iso):
        if ($iso == null or $iso == "") then null
        else ((($iso | fromdateiso8601) - $now) / 2629800 * 10 | round / 10)
        end;

    def header: [
        "ProductName","SkuPartNumber",
        "NextRenewalDate","MonthsUntilNextRenewal","SubscriptionStatus","IsTrial","SubscriptionCreatedDate",
        "SubscriptionLicenses",
        "SkuTotalLicenses","SkuConsumedLicenses","SkuAvailableLicenses",
        "ServicePlans","CommerceSubscriptionId","SkuId","AppliesTo","CapabilityStatus"
    ];

    def sub_row:
        . as $sub | ($sku_by_id[.skuId] // {}) as $sku | [
            ($map[.skuId] // .skuPartNumber // $sku.skuPartNumber),
            (.skuPartNumber // $sku.skuPartNumber),
            .nextLifecycleDateTime,
            months_until(.nextLifecycleDateTime),
            .status,
            .isTrial,
            .createdDateTime,
            (.totalLicenses // 0),
            ($sku.prepaidUnits.enabled // 0),
            ($sku.consumedUnits // 0),
            (($sku.prepaidUnits.enabled // 0) - ($sku.consumedUnits // 0)),
            ([$sku.servicePlans[]?.servicePlanName] | join(";")),
            (.commerceSubscriptionId // .id // ""),
            .skuId,
            ($sku.appliesTo // ""),
            ($sku.capabilityStatus // "")
        ];

    def sku_only_row: [
        ($map[.skuId] // .skuPartNumber),
        .skuPartNumber,
        "", "", "", "", "",
        "",
        (.prepaidUnits.enabled // 0),
        (.consumedUnits // 0),
        ((.prepaidUnits.enabled // 0) - (.consumedUnits // 0)),
        ([.servicePlans[]?.servicePlanName] | join(";")),
        "", .skuId, (.appliesTo // ""), (.capabilityStatus // "")
    ];

    header,
    ($subs | sort_by(.nextLifecycleDateTime // "9999")[] | sub_row),
    ($skus | map(select(.skuId as $s | $sub_skus | index($s) | not))
           | sort_by($map[.skuId] // .skuPartNumber)[]
           | sku_only_row)
    | @csv
' > "$PRODUCTS_CSV"

PRODUCT_ROWS=$(($(wc -l < "$PRODUCTS_CSV") - 1))
echo "Wrote $PRODUCT_ROWS product rows -> $PRODUCTS_CSV"

# ---------------------------------------------------------------------------
# Users (+ signInActivity if accessible)
# ---------------------------------------------------------------------------
echo "Collecting users..."
PROPS_BASE='id,userPrincipalName,displayName,givenName,surname,mail,userType,accountEnabled,createdDateTime,department,jobTitle,officeLocation,mobilePhone,businessPhones,city,state,country,usageLocation,proxyAddresses,assignedLicenses'
PROPS_FULL="${PROPS_BASE},signInActivity"

SIGNIN_AVAILABLE=1
if ! USERS=$(get_all_pages "https://graph.microsoft.com/v1.0/users?\$select=${PROPS_FULL}&\$top=999" 2>/dev/null); then
    echo "WARNING: signInActivity unavailable (no AuditLog.Read.All consent); retrying without it." >&2
    SIGNIN_AVAILABLE=0
    USERS=$(get_all_pages "https://graph.microsoft.com/v1.0/users?\$select=${PROPS_BASE}&\$top=999")
fi

CUTOFF_EPOCH=$((NOW_EPOCH - STALE_DAYS * 86400))

USERS_FILE=$(mktemp)
echo "$USERS" > "$USERS_FILE"

# Shared jq prelude: functions used by both the full user CSV and the stale subset.
read -r -d '' JQ_PRELUDE <<'JQ' || true
def days_since($iso):
    if ($iso == null or $iso == "") then null
    else (($now - ($iso | fromdateiso8601)) / 86400 | floor)
    end;
def most_recent:
    [ .signInActivity.lastSignInDateTime,
      .signInActivity.lastNonInteractiveSignInDateTime,
      .signInActivity.lastSuccessfulSignInDateTime ]
    | map(select(. != null and . != ""))
    | sort | reverse | .[0] // null;
def license_names:
    [ (.assignedLicenses // [])[] | ($map[.skuId] // .skuId) ] | join(";");
def project: [
    .userPrincipalName, .displayName, .givenName, .surname, .mail,
    .userType, .accountEnabled, .createdDateTime,
    .department, .jobTitle, .officeLocation, .mobilePhone,
    ((.businessPhones // []) | join(";")),
    .city, .state, .country, .usageLocation,
    ((.proxyAddresses // []) | join(";")),
    license_names,
    ((.assignedLicenses // []) | length),
    .signInActivity.lastSignInDateTime,
    .signInActivity.lastNonInteractiveSignInDateTime,
    .signInActivity.lastSuccessfulSignInDateTime,
    days_since(most_recent),
    .id
];
def header: [
    "UserPrincipalName","DisplayName","FirstName","LastName","Mail",
    "UserType","AccountEnabled","CreatedDateTime",
    "Department","JobTitle","OfficeLocation","MobilePhone","BusinessPhones",
    "City","State","Country","UsageLocation","ProxyAddresses",
    "AssignedLicenses","LicenseCount",
    "LastSignInDateTime","LastNonInteractiveSignInDateTime","LastSuccessfulSignInDateTime",
    "DaysSinceLastActivity","UserId"
];
JQ

jq -r --argjson map "$SKU_MAP_JSON" --argjson now "$NOW_EPOCH" "
$JQ_PRELUDE
header, (.[] | project) | @csv
" < "$USERS_FILE" > "$USERS_CSV"

TOTAL_USERS=$(($(wc -l < "$USERS_CSV") - 1))
echo "Wrote $TOTAL_USERS user rows -> $USERS_CSV"

# ---------------------------------------------------------------------------
# Stale users (enabled members past the threshold, or never signed in & old)
# ---------------------------------------------------------------------------
jq -r --argjson map "$SKU_MAP_JSON" --argjson now "$NOW_EPOCH" --argjson stale "$STALE_DAYS" --argjson cutoff "$CUTOFF_EPOCH" "
$JQ_PRELUDE
def is_stale:
    (.accountEnabled == true)
    and (.userType != \"Guest\")
    and (
        (days_since(most_recent)) as \$d |
        (\$d != null and \$d >= \$stale)
        or
        (\$d == null and (.createdDateTime != null) and (.createdDateTime != \"\")
            and (.createdDateTime | fromdateiso8601) < \$cutoff)
    );
header, (.[] | select(is_stale) | project) | @csv
" < "$USERS_FILE" > "$STALE_CSV"

TOTAL_STALE=$(($(wc -l < "$STALE_CSV") - 1))
echo "Wrote $TOTAL_STALE stale-user rows -> $STALE_CSV"

# ---------------------------------------------------------------------------
# Security review: Conditional Access, security defaults, privileged roles,
# MFA registration, guests, legacy auth, OAuth app consent risk, Identity
# Protection risk signals, Secure Score. Every sub-check degrades gracefully
# (header-only CSV + warning) if the signed-in account lacks the Graph
# permission for it, the same way signInActivity/subscriptions already do.
# ---------------------------------------------------------------------------
CA_AVAILABLE=0
ROLES_AVAILABLE=0
MFA_AVAILABLE=0
LEGACYAUTH_AVAILABLE=0
OAUTH_AVAILABLE=0
RISKYUSERS_AVAILABLE=0
SECURESCORE_AVAILABLE=0
SECDEFAULTS_ENABLED=""
GA_COUNT=0
SECURESCORE_CURRENT=""
SECURESCORE_MAX=""
TOTAL_FINDINGS=0
HIGH_RISK_PERMS_JSON='["Mail.ReadWrite","Mail.Send","Mail.ReadWrite.All","Mail.Send.All","MailboxSettings.ReadWrite","Files.ReadWrite.All","Sites.ReadWrite.All","Sites.FullControl.All","Sites.Manage.All","Directory.ReadWrite.All","Directory.AccessAsUser.All","RoleManagement.ReadWrite.Directory","User.ReadWrite.All","User.Export.All","Group.ReadWrite.All","GroupMember.ReadWrite.All","Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","Policy.ReadWrite.ConditionalAccess","full_access_as_app","Contacts.ReadWrite","EWS.AccessAsUser.All","Exchange.ManageAsApp"]'

if [[ $SKIP_SECURITY -ne 1 ]]; then

echo ""
echo "Running security review..."

# --- Conditional Access policies + security defaults ---
echo "Collecting Conditional Access policies..."
CA_POLICIES='[]'
if CA_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999' 2>/dev/null); then
    CA_POLICIES="$CA_TRY"
    CA_AVAILABLE=1
else
    echo "WARNING: could not read Conditional Access policies (requires Policy.Read.All); skipping." >&2
fi

if SECDEFAULTS_JSON=$(curl -sf -H "$AUTH" 'https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy' 2>/dev/null); then
    SECDEFAULTS_ENABLED=$(jq -r '.isEnabled' <<<"$SECDEFAULTS_JSON")
else
    echo "WARNING: could not read the security defaults policy (requires Policy.Read.All)." >&2
fi

jq -r '
def header: ["DisplayName","State","CreatedDateTime","ModifiedDateTime","IncludeUsers","ExcludeUsers","IncludeApplications","ExcludeApplications","ClientAppTypes","GrantControls","GrantOperator","SessionControls","PolicyId"];
def row: [
    .displayName, .state, .createdDateTime, .modifiedDateTime,
    ((.conditions.users.includeUsers // []) | join(";")),
    ((.conditions.users.excludeUsers // []) | join(";")),
    ((.conditions.applications.includeApplications // []) | join(";")),
    ((.conditions.applications.excludeApplications // []) | join(";")),
    ((.conditions.clientAppTypes // []) | join(";")),
    ((.grantControls.builtInControls // []) | join(";")),
    (.grantControls.operator // ""),
    (if .sessionControls == null then "" else ([(.sessionControls | keys[])] | join(";")) end),
    .id
];
header, (.[] | row) | @csv
' <<<"$CA_POLICIES" > "$CA_CSV"
echo "Wrote $(($(wc -l < "$CA_CSV") - 1)) Conditional Access policy rows -> $CA_CSV"

# --- Privileged directory role assignments ---
echo "Collecting privileged role assignments..."
ROLE_ASSIGNMENTS='[]'
if ROLE_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=principal($select=id,displayName,userPrincipalName,userType,accountEnabled),roleDefinition&$top=999' 2>/dev/null); then
    ROLE_ASSIGNMENTS="$ROLE_TRY"
    ROLES_AVAILABLE=1
else
    echo "WARNING: could not read directory role assignments (requires RoleManagement.Read.Directory); skipping." >&2
fi

jq -r '
def privileged_roles: ["Global Administrator","Privileged Role Administrator","Privileged Authentication Administrator",
  "Security Administrator","Application Administrator","Cloud Application Administrator","Exchange Administrator",
  "SharePoint Administrator","User Administrator","Helpdesk Administrator","Conditional Access Administrator",
  "Authentication Administrator","Partner Tier2 Support","Hybrid Identity Administrator","Domain Name Administrator"];
def header: ["RoleDisplayName","IsHighPrivilegeRole","PrincipalType","PrincipalDisplayName","PrincipalUserPrincipalName","PrincipalId","DirectoryScopeId","RoleTemplateId"];
def ptype: (.principal."@odata.type" // "" | sub("#microsoft.graph."; ""));
def row: [
    (.roleDefinition.displayName // .roleDefinitionId),
    ((.roleDefinition.displayName // "") as $n | privileged_roles | index($n) != null),
    ptype,
    (.principal.displayName // ""),
    (.principal.userPrincipalName // ""),
    .principalId,
    .directoryScopeId,
    (.roleDefinition.templateId // .roleDefinitionId)
];
header, (.[] | row) | @csv
' <<<"$ROLE_ASSIGNMENTS" > "$ROLES_CSV"
echo "Wrote $(($(wc -l < "$ROLES_CSV") - 1)) role assignment rows -> $ROLES_CSV"

GA_COUNT=$(jq '[.[] | select(.roleDefinition.displayName=="Global Administrator") | .principalId] | unique | length' <<<"$ROLE_ASSIGNMENTS")

# --- MFA / authentication method registration ---
echo "Collecting authentication method registration details..."
MFA_FILE=$(mktemp)
echo '[]' > "$MFA_FILE"
if MFA_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$top=999' 2>/dev/null); then
    echo "$MFA_TRY" > "$MFA_FILE"
    MFA_AVAILABLE=1
else
    echo "WARNING: could not read authentication method registration details (requires AuditLog.Read.All); skipping." >&2
fi

jq -r '
def header: ["UserPrincipalName","DisplayName","UserType","IsAdmin","IsMfaRegistered","IsMfaCapable","IsSsprRegistered","IsSsprCapable","IsPasswordlessCapable","DefaultMfaMethod","MethodsRegistered","LastUpdatedDateTime"];
def row: [
    .userPrincipalName, .userDisplayName, .userType,
    .isAdmin, .isMfaRegistered, .isMfaCapable, .isSsprRegistered, .isSsprCapable, .isPasswordlessCapable,
    .defaultMfaMethod, ((.methodsRegistered // []) | join(";")), .lastUpdatedDateTime
];
header, (.[] | row) | @csv
' < "$MFA_FILE" > "$MFA_CSV"
echo "Wrote $(($(wc -l < "$MFA_CSV") - 1)) MFA registration rows -> $MFA_CSV"

NO_MFA_ADMINS=$(jq -c '[.[] | select(.isAdmin == true and .isMfaRegistered == false) | {userPrincipalName, userDisplayName}]' < "$MFA_FILE")
rm -f "$MFA_FILE"

# --- Guest users (derived from the already-fetched user list, no extra call) ---
echo "Collecting guest users..."
jq -r --argjson map "$SKU_MAP_JSON" --argjson now "$NOW_EPOCH" "
$JQ_PRELUDE
header, (.[] | select(.userType == \"Guest\") | project) | @csv
" < "$USERS_FILE" > "$GUESTS_CSV"
echo "Wrote $(($(wc -l < "$GUESTS_CSV") - 1)) guest user rows -> $GUESTS_CSV"

PRIV_UPNS_JSON=$(jq -c '[.[] | .principal.userPrincipalName // empty] | unique' <<<"$ROLE_ASSIGNMENTS")
STALE_ADMINS=$(jq -c --argjson upns "$PRIV_UPNS_JSON" --argjson now "$NOW_EPOCH" --argjson stale "$STALE_DAYS" "
$JQ_PRELUDE
[ .[]
  | select(.userPrincipalName as \$u | \$upns | index(\$u) != null)
  | . as \$row
  | (days_since(most_recent)) as \$d
  | select(\$d != null and \$d >= \$stale)
  | {userPrincipalName: \$row.userPrincipalName, days: \$d} ]
" < "$USERS_FILE")

# --- Legacy authentication sign-ins ---
if [[ $SIGNIN_AVAILABLE -eq 1 ]]; then
    echo "Scanning sign-in logs for legacy authentication (last ${LEGACY_AUTH_DAYS}d)..."
    LEGACY_CUTOFF_ISO=$(date -u -d "@$((NOW_EPOCH - LEGACY_AUTH_DAYS * 86400))" +%Y-%m-%dT%H:%M:%SZ)
    LEGACY_APPS=("Authenticated SMTP" "Autodiscover" "Exchange ActiveSync" "Exchange Online PowerShell" "Exchange Web Services" "IMAP4" "MAPI over HTTP" "Offline Address Book" "Outlook Anywhere (RPC over HTTP)" "Outlook Service" "POP3" "Reporting Web Services" "Other clients")
    OR_CLAUSE=""
    for app in "${LEGACY_APPS[@]}"; do
        [[ -n "$OR_CLAUSE" ]] && OR_CLAUSE="${OR_CLAUSE} or "
        OR_CLAUSE="${OR_CLAUSE}clientAppUsed eq '${app}'"
    done
    LEGACY_FILTER="createdDateTime ge ${LEGACY_CUTOFF_ISO} and (${OR_CLAUSE})"
    LEGACY_FILTER_ENC=$(python3 -c "import urllib.parse,sys; print(urllib.parse.quote(sys.argv[1]))" "$LEGACY_FILTER")
    LEGACY_URL="https://graph.microsoft.com/v1.0/auditLogs/signIns?\$filter=${LEGACY_FILTER_ENC}&\$select=userPrincipalName,userDisplayName,appDisplayName,clientAppUsed,ipAddress,createdDateTime,status&\$top=999"

    LEGACY_MERGED='[]'
    LEGACY_PAGE_COUNT=0
    LEGACY_MAX_PAGES=25
    url="$LEGACY_URL"
    LEGACY_FETCH_OK=1
    while [[ -n "$url" && "$url" != "null" && $LEGACY_PAGE_COUNT -lt $LEGACY_MAX_PAGES ]]; do
        page=$(curl -sf -H "$AUTH" "$url") || { LEGACY_FETCH_OK=0; break; }
        LEGACY_MERGED=$(jq -c --argjson m "$LEGACY_MERGED" '$m + .value' <<<"$page")
        url=$(jq -r '."@odata.nextLink" // empty' <<<"$page")
        LEGACY_PAGE_COUNT=$((LEGACY_PAGE_COUNT + 1))
    done

    if [[ $LEGACY_FETCH_OK -eq 1 ]]; then
        LEGACYAUTH_AVAILABLE=1
        if [[ -n "$url" && "$url" != "null" ]]; then
            echo "NOTE: legacy-auth sign-in scan hit the ${LEGACY_MAX_PAGES}-page cap; results are truncated. Narrow with --legacy-auth-days." >&2
        fi
    else
        echo "WARNING: could not read sign-in logs for legacy-auth detection; skipping." >&2
    fi

    LEGACY_AGG=$(jq -c '
        group_by([.userPrincipalName, .clientAppUsed, .appDisplayName]) |
        map(
            (.[0].userPrincipalName) as $u | (.[0].clientAppUsed) as $c | (.[0].appDisplayName) as $a |
            (map(select((.status.errorCode // 0) == 0)) | length) as $succ |
            (map(.createdDateTime) | sort) as $times |
            (map(.ipAddress) | unique | .[0:3]) as $ips |
            {userPrincipalName: $u, clientAppUsed: $c, appDisplayName: $a,
             attemptCount: length, successCount: $succ, failureCount: (length - $succ),
             firstSeen: ($times[0] // ""), lastSeen: ($times[-1] // ""), sampleIps: $ips}
        )
    ' <<<"$LEGACY_MERGED")

    jq -r '
    def header: ["UserPrincipalName","ClientAppUsed","ApplicationDisplayName","AttemptCount","SuccessCount","FailureCount","FirstSeen","LastSeen","SampleIpAddresses"];
    def row: [.userPrincipalName, .clientAppUsed, .appDisplayName, .attemptCount, .successCount, .failureCount, .firstSeen, .lastSeen, (.sampleIps | join(";"))];
    header, (.[] | row) | @csv
    ' <<<"$LEGACY_AGG" > "$LEGACYAUTH_CSV"
    echo "Wrote $(($(wc -l < "$LEGACYAUTH_CSV") - 1)) legacy-auth aggregate rows -> $LEGACYAUTH_CSV"
else
    echo "Skipping legacy-auth sign-in scan (signInActivity unavailable)." >&2
    LEGACY_AGG='[]'
    printf 'UserPrincipalName,ClientAppUsed,ApplicationDisplayName,AttemptCount,SuccessCount,FailureCount,FirstSeen,LastSeen,SampleIpAddresses\n' > "$LEGACYAUTH_CSV"
fi

# --- OAuth app consent / enterprise app risk ---
echo "Collecting OAuth app consent grants..."
GRAPH_SP_APPID='00000003-0000-0000-c000-000000000000'
GRAPH_SP_JSON='{}'
if GRAPH_SP_TRY=$(curl -sf -H "$AUTH" "https://graph.microsoft.com/v1.0/servicePrincipals(appId='${GRAPH_SP_APPID}')?\$select=id,appRoles" 2>/dev/null); then
    GRAPH_SP_JSON="$GRAPH_SP_TRY"
fi
GRAPH_SP_ID=$(jq -r '.id // empty' <<<"$GRAPH_SP_JSON")

SP_FILE=$(mktemp); echo '[]' > "$SP_FILE"
DEL_FILE=$(mktemp); echo '[]' > "$DEL_FILE"
APPASSIGN_FILE=$(mktemp); echo '[]' > "$APPASSIGN_FILE"

if [[ -n "$GRAPH_SP_ID" ]]; then
    if SP_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,publisherName,verifiedPublisher,appOwnerOrganizationId,signInAudience,passwordCredentials,keyCredentials&$top=999' 2>/dev/null); then
        echo "$SP_TRY" > "$SP_FILE"
        if DEL_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants?$top=999' 2>/dev/null); then
            echo "$DEL_TRY" > "$DEL_FILE"
        fi
        if APP_TRY=$(get_all_pages "https://graph.microsoft.com/v1.0/servicePrincipals/${GRAPH_SP_ID}/appRoleAssignedTo?\$top=999" 2>/dev/null); then
            echo "$APP_TRY" > "$APPASSIGN_FILE"
        fi
        OAUTH_AVAILABLE=1
    fi
fi

if [[ $OAUTH_AVAILABLE -ne 1 ]]; then
    echo "WARNING: could not read service principals / OAuth grants (requires Directory.Read.All); skipping app consent review." >&2
fi

GRAPH_APP_ROLES_JSON=$(jq -c '.appRoles // []' <<<"$GRAPH_SP_JSON")

jq -r --slurpfile sp "$SP_FILE" --slurpfile delegated "$DEL_FILE" --slurpfile appassignments "$APPASSIGN_FILE" \
      --argjson graphroles "$GRAPH_APP_ROLES_JSON" --argjson now "$NOW_EPOCH" --argjson hr "$HIGH_RISK_PERMS_JSON" '
def is_ms_owned: (.appOwnerOrganizationId // "") == "f8cdef31-a31e-4b4a-93e4-5f571e91255a";
def cred_count: ((.passwordCredentials // []) | length) + ((.keyCredentials // []) | length);
def has_expired_cred($now): ([(.passwordCredentials // [])[], (.keyCredentials // [])[]] | map(select(.endDateTime != null and (.endDateTime | fromdateiso8601) < $now)) | length) > 0;

($sp[0] // [] | map({(.id): .}) | add // {}) as $sp_by_id |
($graphroles // [] | map({(.id): .value}) | add // {}) as $role_by_id |

def client_fields($spId):
    ($sp_by_id[$spId] // {}) as $c | [
        ($c.displayName // $spId), ($c.appId // ""), ($c | is_ms_owned),
        ($c.verifiedPublisher.displayName // ""), ($c | cred_count), ($c | has_expired_cred($now))
    ];

def app_row:
    . as $a | (client_fields($a.principalId)) as $cf | ($role_by_id[$a.appRoleId] // $a.appRoleId) as $perm | [
        $cf[0], $cf[1], $cf[2], $cf[3], $cf[4], $cf[5],
        "Application", ($a.resourceDisplayName // "Microsoft Graph"), $perm, "N/A", "",
        ($hr | index($perm) != null)
    ];

def delegated_rows:
    . as $g | (client_fields($g.clientId)) as $cf | (($sp_by_id[$g.resourceId].displayName) // $g.resourceId) as $resourceName |
    (($g.scope // "") | split(" ") | map(select(length > 0))) as $scopes |
    $scopes[] as $perm | [
        $cf[0], $cf[1], $cf[2], $cf[3], $cf[4], $cf[5],
        "Delegated", $resourceName, $perm, $g.consentType, (if $g.consentType == "Principal" then ($g.principalId // "") else "" end),
        ($hr | index($perm) != null)
    ];

def header: ["ClientAppDisplayName","ClientAppId","ClientAppOwnerIsMicrosoft","ClientAppVerifiedPublisher","ClientAppCredentialCount","ClientAppHasExpiredCredential","PermissionType","ResourceDisplayName","Permission","ConsentType","ConsentGrantedToPrincipalId","HighRisk"];

header,
(($appassignments[0] // [])[] | app_row),
(($delegated[0] // [])[] | delegated_rows)
| @csv
' <<< 'null' > "$OAUTH_CSV"
echo "Wrote $(($(wc -l < "$OAUTH_CSV") - 1)) OAuth consent rows -> $OAUTH_CSV"

OAUTH_RISKY=$(jq -c --slurpfile sp "$SP_FILE" --slurpfile delegated "$DEL_FILE" --slurpfile appassignments "$APPASSIGN_FILE" \
      --argjson graphroles "$GRAPH_APP_ROLES_JSON" --argjson hr "$HIGH_RISK_PERMS_JSON" '
def is_ms_owned: (.appOwnerOrganizationId // "") == "f8cdef31-a31e-4b4a-93e4-5f571e91255a";
($sp[0] // [] | map({(.id): .}) | add // {}) as $sp_by_id |
($graphroles // [] | map({(.id): .value}) | add // {}) as $role_by_id |
[
  (($appassignments[0] // [])[] | . as $a | ($sp_by_id[$a.principalId] // {}) as $c |
    select(($c | is_ms_owned) | not) |
    ($role_by_id[$a.appRoleId] // $a.appRoleId) as $perm |
    select($hr | index($perm) != null) |
    {clientAppDisplayName: ($c.displayName // $a.principalId), permission: $perm, permissionType: "Application", consentType: "N/A"}
  ),
  (($delegated[0] // [])[] | . as $g | ($sp_by_id[$g.clientId] // {}) as $c |
    select(($c | is_ms_owned) | not) |
    (($g.scope // "") | split(" ") | map(select(length > 0)))[] as $perm |
    select($hr | index($perm) != null) |
    {clientAppDisplayName: ($c.displayName // $g.clientId), permission: $perm, permissionType: "Delegated", consentType: $g.consentType}
  )
]
' <<< 'null')

EXPIRED_CRED_APPS=$(jq -c --slurpfile sp "$SP_FILE" --argjson now "$NOW_EPOCH" '
[ $sp[0][]? | select( ([(.passwordCredentials // [])[], (.keyCredentials // [])[]] | map(select(.endDateTime != null and (.endDateTime | fromdateiso8601) < $now)) | length) > 0 ) | .displayName ] | unique
' <<< 'null')

rm -f "$SP_FILE" "$DEL_FILE" "$APPASSIGN_FILE"

# --- Identity Protection risky users ---
echo "Collecting Identity Protection risky users..."
RISKY_USERS='[]'
if RISKY_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/identityProtection/riskyUsers?$top=500' 2>/dev/null); then
    RISKY_USERS="$RISKY_TRY"
    RISKYUSERS_AVAILABLE=1
else
    echo "WARNING: could not read Identity Protection risky users (requires Entra ID P2 + IdentityRiskyUser.Read.All); skipping." >&2
fi

jq -r '
def header: ["UserPrincipalName","DisplayName","RiskLevel","RiskState","RiskDetail","RiskLastUpdatedDateTime","IsProcessing"];
def row: [.userPrincipalName, .userDisplayName, .riskLevel, .riskState, .riskDetail, .riskLastUpdatedDateTime, .isProcessing];
header, (.[] | row) | @csv
' <<<"$RISKY_USERS" > "$RISKYUSERS_CSV"
echo "Wrote $(($(wc -l < "$RISKYUSERS_CSV") - 1)) risky-user rows -> $RISKYUSERS_CSV"

# --- Secure Score ---
echo "Collecting Secure Score..."
SECURE_SCORE_ALL='[]'
if SS_TRY=$(get_all_pages 'https://graph.microsoft.com/v1.0/security/secureScores?$top=999' 2>/dev/null); then
    SECURE_SCORE_ALL="$SS_TRY"
    SECURESCORE_AVAILABLE=1
else
    echo "WARNING: could not read Secure Score (requires SecurityEvents.Read.All); skipping." >&2
fi
SECURE_SCORE_LATEST=$(jq -c '(sort_by(.createdDateTime) | last) // {}' <<<"$SECURE_SCORE_ALL")
SCORE_PROFILES='[]'
if [[ $SECURESCORE_AVAILABLE -eq 1 ]]; then
    SP_TRY2=$(get_all_pages 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles?$top=999' 2>/dev/null) && SCORE_PROFILES="$SP_TRY2"
fi

jq -r --argjson profiles "$SCORE_PROFILES" '
($profiles | map({(.id): .}) | add // {}) as $p_by_id |
(.controlScores // []) as $scores |
def header: ["ControlName","Title","Category","CurrentScore","MaxScore","PercentImplemented","Rank","Tier","UserImpact","ImplementationCost","ActionType","Remediation","RemediationImpact"];
def row:
    .controlName as $cn | ($p_by_id[$cn] // {}) as $p |
    (.score // 0) as $cur | ($p.maxScore // 0) as $max | [
        $cn, ($p.title // .description // $cn), ($p.controlCategory // .controlCategory // ""),
        $cur, $max, (if $max > 0 then ((100 * $cur / $max) * 10 | round / 10) else "" end),
        $p.rank, $p.tier, $p.userImpact, $p.implementationCost, $p.actionType, $p.remediation, $p.remediationImpact
    ];
header, ($scores | sort_by((.score // 0)) | .[] | row) | @csv
' <<<"$SECURE_SCORE_LATEST" > "$SECURESCORE_CSV"
echo "Wrote $(($(wc -l < "$SECURESCORE_CSV") - 1)) Secure Score control rows -> $SECURESCORE_CSV"

SECURESCORE_CURRENT=$(jq -r '.currentScore // ""' <<<"$SECURE_SCORE_LATEST")
SECURESCORE_MAX=$(jq -r '.maxScore // ""' <<<"$SECURE_SCORE_LATEST")

# --- Consolidated findings ---
echo "Compiling security findings..."
SECDEFAULTS_BOOL="null"
[[ "$SECDEFAULTS_ENABLED" == "true" ]] && SECDEFAULTS_BOOL="true"
[[ "$SECDEFAULTS_ENABLED" == "false" ]] && SECDEFAULTS_BOOL="false"

jq -r --argjson secdefaults "$SECDEFAULTS_BOOL" --argjson ca "$CA_POLICIES" --argjson roles "$ROLE_ASSIGNMENTS" \
      --argjson noMfaAdmins "$NO_MFA_ADMINS" --argjson staleAdmins "$STALE_ADMINS" --argjson legacyAgg "$LEGACY_AGG" \
      --argjson oauthRisky "$OAUTH_RISKY" --argjson expiredCredApps "$EXPIRED_CRED_APPS" --argjson riskyUsers "$RISKY_USERS" '
def sev($s;$cat;$find;$obj;$rec): {severity:$s, category:$cat, finding:$find, affected:$obj, recommendation:$rec};

( if ($secdefaults == false) and (($ca // []) | map(select(.state=="enabled")) | length) == 0 then
    [sev("High";"Identity Baseline";"No baseline identity protection: security defaults are disabled and no Conditional Access policy is enabled";"Tenant";"Enable security defaults or create enforced Conditional Access policies requiring MFA")]
  else [] end ) as $baseline_findings |

[ ($ca // [])[] | select(.state=="disabled")
  | sev("Medium";"Conditional Access";"Conditional Access policy is disabled"; .displayName; "Review and enable or remove the disabled policy") ] as $disabled_ca_findings |

(($roles // []) | map(select(.roleDefinition.displayName=="Global Administrator")) | map(.principalId) | unique | length) as $ga_count |
( if $ga_count == 1 then
    [sev("High";"Privileged Roles";"Only one Global Administrator account exists (no break-glass redundancy)"; "Tenant"; "Provision at least one additional break-glass Global Administrator account")]
  elif $ga_count > 5 then
    [sev("Medium";"Privileged Roles";("Large number of Global Administrators (" + ($ga_count|tostring) + ")"); "Tenant"; "Review Global Administrator membership and remove unnecessary assignments")]
  else [] end ) as $ga_findings |

[ ($roles // [])[] | select(.principal.userType=="Guest")
  | sev("High";"Privileged Roles";"Guest account holds a privileged directory role"; (.principal.userPrincipalName // .principalId); ("Remove the " + (.roleDefinition.displayName // .roleDefinitionId) + " assignment or convert to a managed member account")) ] as $guest_admin_findings |

[ ($noMfaAdmins // [])[]
  | sev("High";"MFA";"Privileged account has no MFA method registered"; .userPrincipalName; "Register a strong MFA method for this admin account immediately") ] as $no_mfa_findings |

[ ($staleAdmins // [])[]
  | sev("Medium";"Privileged Roles";("Privileged account inactive for " + (.days|tostring) + " days"); .userPrincipalName; "Disable or remove the privileged role assignment if no longer needed") ] as $stale_admin_findings |

( if (($legacyAgg // []) | length) > 0 then
    [sev("Medium";"Legacy Authentication";(($legacyAgg | length | tostring) + " distinct user/app combinations used legacy (basic) authentication in the lookback window"); "Tenant"; "Block legacy authentication with a Conditional Access policy")]
  else [] end ) as $legacy_findings |

[ ($oauthRisky // [])[]
  | sev( (if .permissionType=="Delegated" and .consentType=="Principal" then "High" else "Medium" end);
         "OAuth Consent";
         ("High-risk " + .permissionType + " permission (" + .permission + ") granted to third-party app");
         .clientAppDisplayName;
         "Review app consent in Enterprise Applications and revoke if unnecessary" ) ] as $oauth_findings |

[ ($expiredCredApps // [])[]
  | sev("Low";"App Credentials";"App registration has an expired credential still present"; .; "Remove the expired client secret/certificate") ] as $expired_cred_findings |

[ ($riskyUsers // [])[] | select(.riskState=="atRisk" or .riskState=="confirmedCompromised")
  | sev( (if .riskLevel=="high" then "High" else "Medium" end);
         "Identity Protection"; ("User flagged as risky: " + (.riskDetail // "unspecified")); (.userPrincipalName // .userDisplayName); "Investigate in Entra ID Protection and remediate or confirm compromise" ) ] as $risky_user_findings |

($baseline_findings + $disabled_ca_findings + $ga_findings + $guest_admin_findings + $no_mfa_findings +
 $stale_admin_findings + $legacy_findings + $oauth_findings + $expired_cred_findings + $risky_user_findings) as $all |

def sevrank: if .severity=="High" then 0 elif .severity=="Medium" then 1 else 2 end;

(["Severity","Category","Finding","AffectedObject","Recommendation"]),
($all | sort_by(sevrank) | .[] | [.severity, .category, .finding, .affected, .recommendation])
| @csv
' <<< 'null' > "$FINDINGS_CSV"

TOTAL_FINDINGS=$(($(wc -l < "$FINDINGS_CSV") - 1))
echo "Wrote $TOTAL_FINDINGS security finding rows -> $FINDINGS_CSV"

fi

rm -f "$USERS_FILE"

# ---------------------------------------------------------------------------
# Zip everything
# ---------------------------------------------------------------------------
if [[ $SKIP_SECURITY -ne 1 ]]; then
    zip -j -q "$ZIP_PATH" "$USERS_CSV" "$PRODUCTS_CSV" "$STALE_CSV" \
        "$CA_CSV" "$ROLES_CSV" "$MFA_CSV" "$GUESTS_CSV" "$LEGACYAUTH_CSV" \
        "$OAUTH_CSV" "$RISKYUSERS_CSV" "$SECURESCORE_CSV" "$FINDINGS_CSV"
else
    zip -j -q "$ZIP_PATH" "$USERS_CSV" "$PRODUCTS_CSV" "$STALE_CSV"
fi

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
echo ""
echo "========================================="
echo "  M365 Tenant Assessment Complete"
echo "========================================="
echo "Tenant:                    $TENANT_NAME"
echo "Tenant ID:                 $TENANT_ID"
echo "Total users:               $TOTAL_USERS"
echo "Product rows:              $PRODUCT_ROWS"
echo "Stale users (>${STALE_DAYS}d):         $TOTAL_STALE"
if [[ $SUBS_AVAILABLE -ne 1 ]]; then
    echo "NOTE: subscription data was unavailable; NextRenewalDate column is blank."
fi
if [[ $SIGNIN_AVAILABLE -ne 1 ]]; then
    echo "NOTE: signInActivity was unavailable; stale list uses created-date heuristic only."
fi
if [[ $SKIP_SECURITY -ne 1 ]]; then
    echo ""
    echo "-- Security review --"
    echo "Security defaults enabled: ${SECDEFAULTS_ENABLED:-unknown}"
    echo "Global Administrators:     $GA_COUNT"
    if [[ -n "$SECURESCORE_CURRENT" ]]; then
        echo "Secure Score:              ${SECURESCORE_CURRENT} / ${SECURESCORE_MAX}"
    fi
    echo "Findings (high/med/low):   $TOTAL_FINDINGS total -> $FINDINGS_CSV"
    for flag_name in "CA_AVAILABLE:Conditional Access policies" \
                      "ROLES_AVAILABLE:privileged role assignments" \
                      "MFA_AVAILABLE:MFA registration details" \
                      "LEGACYAUTH_AVAILABLE:legacy-auth sign-in scan" \
                      "OAUTH_AVAILABLE:OAuth app consent review" \
                      "RISKYUSERS_AVAILABLE:Identity Protection risky users" \
                      "SECURESCORE_AVAILABLE:Secure Score"; do
        flag="${flag_name%%:*}"; label="${flag_name#*:}"
        if [[ "${!flag}" -ne 1 ]]; then
            echo "NOTE: $label unavailable (insufficient Graph permission or feature not licensed); CSV has headers only."
        fi
    done
else
    echo ""
    echo "Security review skipped (--skip-security-review)."
fi
echo ""
echo "Files:"
echo "  $USERS_CSV"
echo "  $PRODUCTS_CSV"
echo "  $STALE_CSV"
if [[ $SKIP_SECURITY -ne 1 ]]; then
    echo "  $CA_CSV"
    echo "  $ROLES_CSV"
    echo "  $MFA_CSV"
    echo "  $GUESTS_CSV"
    echo "  $LEGACYAUTH_CSV"
    echo "  $OAUTH_CSV"
    echo "  $RISKYUSERS_CSV"
    echo "  $SECURESCORE_CSV"
    echo "  $FINDINGS_CSV"
fi
echo "  $ZIP_PATH"
echo ""

# ---------------------------------------------------------------------------
# Trigger the Cloud Shell download dialog
# ---------------------------------------------------------------------------
if [[ $NO_DOWNLOAD -ne 1 ]] && command -v download >/dev/null 2>&1; then
    echo "Opening Cloud Shell download dialog for the zip bundle..."
    download "$ZIP_PATH"
else
    echo "Download path (paste into Manage files -> Download):"
    echo ""
    echo "$ZIP_PATH"
fi
