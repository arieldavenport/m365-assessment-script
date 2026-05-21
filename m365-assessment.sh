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

usage() {
    cat <<EOF
Usage: $0 [options]
  --stale-days N           Stale account threshold in days (default 90)
  --output-dir PATH        Output directory (default: \$HOME)
  --skip-friendly-names    Skip downloading Microsoft's SKU friendly-name map
  --no-download            Skip the Cloud Shell download dialog at the end
  -h, --help               This help
EOF
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --stale-days)          STALE_DAYS="$2"; shift 2 ;;
        --output-dir)          OUTPUT_DIR="$2"; shift 2 ;;
        --skip-friendly-names) SKIP_FRIENDLY=1; shift ;;
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
        page=$(curl -sf -H "$AUTH" "$url")
        merged=$(jq -c --argjson m "$merged" '$m + .value' <<<"$page")
        url=$(jq -r '."@odata.nextLink" // empty' <<<"$page")
    done
    echo "$merged"
}

# ---------------------------------------------------------------------------
# Products / subscribed SKUs + commerce subscriptions (for renewal dates)
# ---------------------------------------------------------------------------
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
    --argjson map  "$SKU_MAP_JSON" '
    ($skus | map({(.skuId): .}) | add // {}) as $sku_by_id |
    ($subs | map(.skuId)) as $sub_skus |

    def header: [
        "ProductName","SkuPartNumber",
        "NextRenewalDate","SubscriptionStatus","IsTrial","SubscriptionCreatedDate",
        "SubscriptionLicenses",
        "SkuTotalLicenses","SkuConsumedLicenses","SkuAvailableLicenses",
        "ServicePlans","CommerceSubscriptionId","SkuId","AppliesTo","CapabilityStatus"
    ];

    def sub_row:
        . as $sub | ($sku_by_id[.skuId] // {}) as $sku | [
            ($map[.skuId] // .skuPartNumber // $sku.skuPartNumber),
            (.skuPartNumber // $sku.skuPartNumber),
            .nextLifecycleDateTime,
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
        "", "", "", "",
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

NOW_EPOCH=$(date -u +%s)
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

rm -f "$USERS_FILE"

# ---------------------------------------------------------------------------
# Zip everything
# ---------------------------------------------------------------------------
zip -j -q "$ZIP_PATH" "$USERS_CSV" "$PRODUCTS_CSV" "$STALE_CSV"

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
echo ""
echo "Files:"
echo "  $USERS_CSV"
echo "  $PRODUCTS_CSV"
echo "  $STALE_CSV"
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
