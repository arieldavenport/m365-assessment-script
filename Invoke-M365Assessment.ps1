<#
.SYNOPSIS
    One-shot Microsoft 365 tenant assessment: users, products, and stale accounts.

.DESCRIPTION
    Designed to be pasted into the Microsoft 365 admin center Cloud Shell (or any
    PowerShell 7+ host with the Microsoft.Graph module installed). Produces three
    CSVs in the working directory:

        M365_Users_<tenant>_<timestamp>.csv
            Mirrors the admin center Active users export, plus sign-in activity
            (LastSignInDateTime / LastNonInteractiveSignInDateTime /
            LastSuccessfulSignInDateTime) and DaysSinceLastActivity.

        M365_Products_<tenant>_<timestamp>.csv
            Subscribed SKUs (the "Products" / "Licenses" view), with friendly
            product names resolved via Microsoft's published mapping CSV.

        M365_StaleUsers_<tenant>_<timestamp>.csv
            Enabled member users whose most recent sign-in activity is older
            than -StaleDays (default 90), or who have never signed in and were
            created more than -StaleDays ago.

.PARAMETER OutputPath
    Folder where CSVs are written. Defaults to current directory.

.PARAMETER StaleDays
    Threshold in days for flagging stale accounts. Default 90.

.PARAMETER SkipFriendlyNames
    Skip the one-time download of Microsoft's SKU friendly-name map. The CSV
    will fall back to raw SkuPartNumber values.

.NOTES
    Required Graph scopes (consented on first connect):
        User.Read.All, Organization.Read.All, Directory.Read.All, AuditLog.Read.All

    signInActivity requires AuditLog.Read.All AND a Microsoft Entra ID P1 or P2
    license on the tenant. If the tenant lacks P1/P2 the script gracefully
    falls back to a user export without sign-in fields.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Get-Location).Path,
    [int]$StaleDays = 90,
    [switch]$SkipFriendlyNames
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# 1. Module bootstrap
# ---------------------------------------------------------------------------
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing module $module ..." -ForegroundColor Yellow
        Install-Module $module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $module -ErrorAction Stop
}

# ---------------------------------------------------------------------------
# 2. Connect to Microsoft Graph
# ---------------------------------------------------------------------------
$scopes = @(
    'User.Read.All',
    'Organization.Read.All',
    'Directory.Read.All',
    'AuditLog.Read.All'
)
Write-Host "Connecting to Microsoft Graph: $($scopes -join ', ')" -ForegroundColor Cyan
Connect-MgGraph -Scopes $scopes -NoWelcome

$context = Get-MgContext
$org     = Get-MgOrganization | Select-Object -First 1
$tenantTag = ($org.DisplayName -replace '[^a-zA-Z0-9]', '_')
$timestamp = (Get-Date).ToString('yyyyMMdd-HHmmss')

Write-Host "Connected to: $($org.DisplayName)  ($($context.TenantId))" -ForegroundColor Green

if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$usersCsv    = Join-Path $OutputPath "M365_Users_${tenantTag}_${timestamp}.csv"
$productsCsv = Join-Path $OutputPath "M365_Products_${tenantTag}_${timestamp}.csv"
$staleCsv    = Join-Path $OutputPath "M365_StaleUsers_${tenantTag}_${timestamp}.csv"

# ---------------------------------------------------------------------------
# 3. SKU friendly-name map (GUID -> Product display name)
# ---------------------------------------------------------------------------
$skuMap = @{}
if (-not $SkipFriendlyNames) {
    $skuCsvUrl = 'https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv'
    try {
        Write-Host "Downloading SKU friendly-name map ..." -ForegroundColor Cyan
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
Write-Host "Collecting subscribed SKUs (products) ..." -ForegroundColor Cyan
$skus = Get-MgSubscribedSku -All

$productRows = foreach ($s in $skus) {
    $enabled   = [int]$s.PrepaidUnits.Enabled
    $consumed  = [int]$s.ConsumedUnits
    [pscustomobject]@{
        ProductName       = Resolve-SkuName -SkuId $s.SkuId -SkuPartNumber $s.SkuPartNumber
        SkuPartNumber     = $s.SkuPartNumber
        SkuId             = $s.SkuId
        AppliesTo         = $s.AppliesTo
        CapabilityStatus  = $s.CapabilityStatus
        TotalLicenses     = $enabled
        ConsumedLicenses  = $consumed
        AvailableLicenses = $enabled - $consumed
        SuspendedLicenses = [int]$s.PrepaidUnits.Suspended
        WarningLicenses   = [int]$s.PrepaidUnits.Warning
        AccountId         = $s.AccountId
        AccountName       = $s.AccountName
        ServicePlans      = (($s.ServicePlans | ForEach-Object { $_.ServicePlanName }) -join ';')
    }
}
$productRows | Sort-Object ProductName | Export-Csv -Path $productsCsv -NoTypeInformation -Encoding UTF8
Write-Host "Wrote $($productRows.Count) product rows -> $productsCsv" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 5. Users + sign-in activity
# ---------------------------------------------------------------------------
Write-Host "Collecting users (may take a minute on large tenants) ..." -ForegroundColor Cyan

$userProperties = @(
    'id','userPrincipalName','displayName','givenName','surname','mail',
    'userType','accountEnabled','createdDateTime','department','jobTitle',
    'officeLocation','mobilePhone','businessPhones','city','state','country',
    'usageLocation','proxyAddresses','assignedLicenses','signInActivity'
)

$signInAvailable = $true
try {
    $users = Get-MgUser -All -Property $userProperties -ErrorAction Stop
} catch {
    Write-Warning "Could not retrieve signInActivity ($($_.Exception.Message)). Retrying without it; stale detection will be limited."
    $signInAvailable = $false
    $userProperties  = $userProperties | Where-Object { $_ -ne 'signInActivity' }
    $users = Get-MgUser -All -Property $userProperties
}

$now = Get-Date
$userRows = foreach ($u in $users) {
    $licenseNames = foreach ($lic in $u.AssignedLicenses) {
        Resolve-SkuName -SkuId $lic.SkuId -SkuPartNumber $null
    }

    $lastSignIn   = $u.SignInActivity.LastSignInDateTime
    $lastNonInter = $u.SignInActivity.LastNonInteractiveSignInDateTime
    $lastSuccess  = $u.SignInActivity.LastSuccessfulSignInDateTime

    $mostRecent = @($lastSignIn, $lastNonInter, $lastSuccess) |
        Where-Object { $_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    $daysSince = $null
    if ($mostRecent) { $daysSince = [int]($now - $mostRecent).TotalDays }

    [pscustomobject]@{
        UserPrincipalName                = $u.UserPrincipalName
        DisplayName                      = $u.DisplayName
        FirstName                        = $u.GivenName
        LastName                         = $u.Surname
        Mail                             = $u.Mail
        UserType                         = $u.UserType
        AccountEnabled                   = $u.AccountEnabled
        CreatedDateTime                  = $u.CreatedDateTime
        Department                       = $u.Department
        JobTitle                         = $u.JobTitle
        OfficeLocation                   = $u.OfficeLocation
        MobilePhone                      = $u.MobilePhone
        BusinessPhones                   = ($u.BusinessPhones -join ';')
        City                             = $u.City
        State                            = $u.State
        Country                          = $u.Country
        UsageLocation                    = $u.UsageLocation
        ProxyAddresses                   = ($u.ProxyAddresses -join ';')
        AssignedLicenses                 = ($licenseNames -join ';')
        LicenseCount                     = @($u.AssignedLicenses).Count
        LastSignInDateTime               = $lastSignIn
        LastNonInteractiveSignInDateTime = $lastNonInter
        LastSuccessfulSignInDateTime     = $lastSuccess
        DaysSinceLastActivity            = $daysSince
        UserId                           = $u.Id
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
Write-Host ("Tenant:               {0}"     -f $org.DisplayName)
Write-Host ("Tenant ID:            {0}"     -f $context.TenantId)
Write-Host ("Total users:          {0}"     -f $userRows.Count)
Write-Host ("  Enabled:            {0}"     -f $enabledCount)
Write-Host ("  Guests:             {0}"     -f $guestCount)
Write-Host ("Products / SKUs:      {0}"     -f $productRows.Count)
Write-Host ("Licenses (consumed/total): {0} / {1}" -f $consumedLic, $totalLic)
Write-Host ("Stale users (>${StaleDays}d):    {0}" -f $staleRows.Count)
if (-not $signInAvailable) {
    Write-Host 'NOTE: signInActivity was unavailable. Stale users fall back to created-date heuristic only.' -ForegroundColor Yellow
}
Write-Host ''
Write-Host 'Files:' -ForegroundColor Green
Write-Host "  $usersCsv"
Write-Host "  $productsCsv"
Write-Host "  $staleCsv"
