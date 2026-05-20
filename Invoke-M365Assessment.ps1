<#
.SYNOPSIS
    One-shot Microsoft 365 tenant assessment: users, products, and stale accounts.

.DESCRIPTION
    Designed to run in the Microsoft 365 admin center / Azure Cloud Shell
    (PowerShell). Uses only Microsoft.Graph.Authentication (pre-loaded in
    Cloud Shell) and direct Graph REST calls via Invoke-MgGraphRequest, so
    no module installs and no assembly version conflicts.

    Produces three CSVs in the working directory:

        M365_Users_<tenant>_<timestamp>.csv
            Active users export plus signInActivity (LastSignInDateTime /
            LastNonInteractiveSignInDateTime / LastSuccessfulSignInDateTime)
            and DaysSinceLastActivity.

        M365_Products_<tenant>_<timestamp>.csv
            Subscribed SKUs (the "Products" / "Licenses" view) with friendly
            product names resolved from Microsoft's published mapping CSV.

        M365_StaleUsers_<tenant>_<timestamp>.csv
            Enabled, non-guest accounts inactive >= -StaleDays (default 90),
            or that have never signed in and were created longer ago than the
            threshold.

.PARAMETER OutputPath
    Folder where CSVs are written. Defaults to current directory.

.PARAMETER StaleDays
    Inactivity threshold for the stale-users CSV. Default 90.

.PARAMETER SkipFriendlyNames
    Skip the one-time download of Microsoft's SKU friendly-name map.

.NOTES
    Required Graph scopes (consented on first connect):
        User.Read.All, Organization.Read.All, Directory.Read.All, AuditLog.Read.All

    signInActivity requires AuditLog.Read.All AND Entra ID P1/P2. Without
    P1/P2 the script still produces the user and product CSVs; sign-in
    columns are blank and the stale list falls back to a created-date
    heuristic.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Get-Location).Path,
    [int]$StaleDays = 90,
    [switch]$SkipFriendlyNames
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 1. Minimal module bootstrap — only Microsoft.Graph.Authentication is needed
# ---------------------------------------------------------------------------
if (-not (Get-Module -Name Microsoft.Graph.Authentication)) {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host 'Installing Microsoft.Graph.Authentication ...' -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# 2. Connect
# ---------------------------------------------------------------------------
$scopes = @(
    'User.Read.All',
    'Organization.Read.All',
    'Directory.Read.All',
    'AuditLog.Read.All'
)

function Connect-ViaAzToken {
    # Reuse the existing Az session (Cloud Shell signs you in silently). This
    # avoids the device-code prompt entirely, which Conditional Access policies
    # frequently block. The Az PowerShell app holds Directory.AccessAsUser.All,
    # which is enough for users + SKUs. signInActivity may still 403 if the
    # tenant lacks AuditLog.Read.All consent on that app -- the script handles
    # that case by falling back.
    if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) { return $false }
    try {
        $tok = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com' -ErrorAction Stop
        $secure = if ($tok.Token -is [System.Security.SecureString]) {
            $tok.Token
        } else {
            ConvertTo-SecureString -String $tok.Token -AsPlainText -Force
        }
        Connect-MgGraph -AccessToken $secure -NoWelcome
        return $true
    } catch {
        Write-Warning "Could not reuse Az token ($($_.Exception.Message)). Falling back to interactive sign-in."
        return $false
    }
}

Write-Host 'Connecting to Microsoft Graph (reusing Cloud Shell Az session) ...' -ForegroundColor Cyan
if (-not (Connect-ViaAzToken)) {
    Write-Host "Interactive sign-in: $($scopes -join ', ')" -ForegroundColor Yellow
    Connect-MgGraph -Scopes $scopes -NoWelcome
}

$context  = Get-MgContext
$orgResp  = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -OutputType PSObject
$org      = $orgResp.value | Select-Object -First 1
$tenantTag = ($org.displayName -replace '[^a-zA-Z0-9]', '_')
$timestamp = (Get-Date).ToString('yyyyMMdd-HHmmss')

Write-Host "Connected to: $($org.displayName)  ($($context.TenantId))" -ForegroundColor Green

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$usersCsv    = Join-Path $OutputPath "M365_Users_${tenantTag}_${timestamp}.csv"
$productsCsv = Join-Path $OutputPath "M365_Products_${tenantTag}_${timestamp}.csv"
$staleCsv    = Join-Path $OutputPath "M365_StaleUsers_${tenantTag}_${timestamp}.csv"

# ---------------------------------------------------------------------------
# 3. Helpers
# ---------------------------------------------------------------------------
function Get-GraphAllPages {
    param([string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($page.value) { foreach ($v in $page.value) { [void]$all.Add($v) } }
        $next = $page.'@odata.nextLink'
    }
    return $all
}

$skuMap = @{}
if (-not $SkipFriendlyNames) {
    $skuCsvUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    try {
        Write-Host 'Downloading SKU friendly-name map ...' -ForegroundColor Cyan
        $tmp = [System.IO.Path]::GetTempFileName()
        Invoke-WebRequest -Uri $skuCsvUrl -OutFile $tmp -UseBasicParsing
        foreach ($row in (Import-Csv -Path $tmp)) {
            if ($row.GUID -and -not $skuMap.ContainsKey($row.GUID)) {
                $skuMap[$row.GUID] = $row.Product_Display_Name
            }
        }
        Remove-Item $tmp -ErrorAction SilentlyContinue
        Write-Host "Loaded $($skuMap.Count) SKU GUID -> name mappings." -ForegroundColor Green
    } catch {
        Write-Warning "Could not download SKU map: $($_.Exception.Message). Falling back to raw SkuPartNumber."
    }
}

function Resolve-SkuName {
    param($SkuId, $SkuPartNumber)
    if ($SkuId -and $skuMap.ContainsKey($SkuId)) { return $skuMap[$SkuId] }
    if ($SkuPartNumber) { return $SkuPartNumber }
    return $SkuId
}

# ---------------------------------------------------------------------------
# 4. Products / subscribed SKUs
# ---------------------------------------------------------------------------
Write-Host 'Collecting subscribed SKUs (products) ...' -ForegroundColor Cyan
$skus = Get-GraphAllPages 'https://graph.microsoft.com/v1.0/subscribedSkus'

$productRows = foreach ($s in $skus) {
    $enabled  = [int]$s.prepaidUnits.enabled
    $consumed = [int]$s.consumedUnits
    [pscustomobject]@{
        ProductName       = Resolve-SkuName -SkuId $s.skuId -SkuPartNumber $s.skuPartNumber
        SkuPartNumber     = $s.skuPartNumber
        SkuId             = $s.skuId
        AppliesTo         = $s.appliesTo
        CapabilityStatus  = $s.capabilityStatus
        TotalLicenses     = $enabled
        ConsumedLicenses  = $consumed
        AvailableLicenses = $enabled - $consumed
        SuspendedLicenses = [int]$s.prepaidUnits.suspended
        WarningLicenses   = [int]$s.prepaidUnits.warning
        AccountId         = $s.accountId
        AccountName       = $s.accountName
        ServicePlans      = (($s.servicePlans | ForEach-Object { $_.servicePlanName }) -join ';')
    }
}
$productRows | Sort-Object ProductName | Export-Csv -Path $productsCsv -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($productRows.Count) product rows -> $productsCsv" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 5. Users + sign-in activity
# ---------------------------------------------------------------------------
Write-Host 'Collecting users (may take a minute on large tenants) ...' -ForegroundColor Cyan

$baseProps  = 'id,userPrincipalName,displayName,givenName,surname,mail,userType,accountEnabled,createdDateTime,department,jobTitle,officeLocation,mobilePhone,businessPhones,city,state,country,usageLocation,proxyAddresses,assignedLicenses'
$withSignIn = "$baseProps,signInActivity"

$signInAvailable = $true
try {
    $users = Get-GraphAllPages "https://graph.microsoft.com/v1.0/users?`$select=$withSignIn&`$top=999"
} catch {
    Write-Warning "Could not retrieve signInActivity ($($_.Exception.Message)). Retrying without it; stale detection will fall back to created-date."
    $signInAvailable = $false
    $users = Get-GraphAllPages "https://graph.microsoft.com/v1.0/users?`$select=$baseProps&`$top=999"
}

$now = Get-Date
$userRows = foreach ($u in $users) {
    $licenseNames = foreach ($lic in $u.assignedLicenses) {
        Resolve-SkuName -SkuId $lic.skuId -SkuPartNumber $null
    }

    $lastSignIn   = $null
    $lastNonInter = $null
    $lastSuccess  = $null
    if ($u.signInActivity) {
        $lastSignIn   = $u.signInActivity.lastSignInDateTime
        $lastNonInter = $u.signInActivity.lastNonInteractiveSignInDateTime
        $lastSuccess  = $u.signInActivity.lastSuccessfulSignInDateTime
    }

    $mostRecent = @($lastSignIn, $lastNonInter, $lastSuccess) |
        Where-Object { $_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    $daysSince = $null
    if ($mostRecent) {
        $dt = if ($mostRecent -is [datetime]) { $mostRecent } else { [datetime]$mostRecent }
        $daysSince = [int]($now - $dt).TotalDays
    }

    [pscustomobject]@{
        UserPrincipalName                = $u.userPrincipalName
        DisplayName                      = $u.displayName
        FirstName                        = $u.givenName
        LastName                         = $u.surname
        Mail                             = $u.mail
        UserType                         = $u.userType
        AccountEnabled                   = $u.accountEnabled
        CreatedDateTime                  = $u.createdDateTime
        Department                       = $u.department
        JobTitle                         = $u.jobTitle
        OfficeLocation                   = $u.officeLocation
        MobilePhone                      = $u.mobilePhone
        BusinessPhones                   = ($u.businessPhones -join ';')
        City                             = $u.city
        State                            = $u.state
        Country                          = $u.country
        UsageLocation                    = $u.usageLocation
        ProxyAddresses                   = ($u.proxyAddresses -join ';')
        AssignedLicenses                 = ($licenseNames -join ';')
        LicenseCount                     = @($u.assignedLicenses).Count
        LastSignInDateTime               = $lastSignIn
        LastNonInteractiveSignInDateTime = $lastNonInter
        LastSuccessfulSignInDateTime     = $lastSuccess
        DaysSinceLastActivity            = $daysSince
        UserId                           = $u.id
    }
}
$userRows | Export-Csv -Path $usersCsv -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($userRows.Count) user rows -> $usersCsv" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 6. Stale users
# ---------------------------------------------------------------------------
$cutoff = $now.AddDays(-$StaleDays)
$staleRows = $userRows | Where-Object {
    $_.AccountEnabled -eq $true -and
    $_.UserType -ne 'Guest' -and
    (
        ($null -ne $_.DaysSinceLastActivity -and $_.DaysSinceLastActivity -ge $StaleDays) -or
        ($null -eq $_.DaysSinceLastActivity -and $_.CreatedDateTime -and ([datetime]$_.CreatedDateTime) -lt $cutoff)
    )
}
$staleRows | Export-Csv -Path $staleCsv -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($staleRows.Count) stale-user rows -> $staleCsv" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# 7. Summary
# ---------------------------------------------------------------------------
$enabledCount = ($userRows | Where-Object AccountEnabled).Count
$guestCount   = ($userRows | Where-Object { $_.UserType -eq 'Guest' }).Count
$totalLic     = ($productRows | Measure-Object TotalLicenses    -Sum).Sum
$consumedLic  = ($productRows | Measure-Object ConsumedLicenses -Sum).Sum

Write-Host ''
Write-Host '=========================================' -ForegroundColor Cyan
Write-Host '  M365 Tenant Assessment Complete'        -ForegroundColor Cyan
Write-Host '=========================================' -ForegroundColor Cyan
Write-Host ("Tenant:                    {0}" -f $org.displayName)
Write-Host ("Tenant ID:                 {0}" -f $context.TenantId)
Write-Host ("Total users:               {0}" -f $userRows.Count)
Write-Host ("  Enabled:                 {0}" -f $enabledCount)
Write-Host ("  Guests:                  {0}" -f $guestCount)
Write-Host ("Products / SKUs:           {0}" -f $productRows.Count)
Write-Host ("Licenses (consumed/total): {0} / {1}" -f $consumedLic, $totalLic)
Write-Host ("Stale users (>${StaleDays}d):         {0}" -f $staleRows.Count)
if (-not $signInAvailable) {
    Write-Host 'NOTE: signInActivity was unavailable. Stale users use created-date heuristic only.' -ForegroundColor Yellow
}
Write-Host ''
Write-Host 'Files:' -ForegroundColor Green
Write-Host "  $usersCsv"
Write-Host "  $productsCsv"
Write-Host "  $staleCsv"
