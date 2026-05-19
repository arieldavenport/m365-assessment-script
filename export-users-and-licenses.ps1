[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputDirectory = ".",

    [Parameter()]
    [switch]$SkipConnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    throw "Microsoft Graph PowerShell modules are required. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}

if (-not $SkipConnect) {
    Connect-MgGraph -Scopes "User.Read.All", "Organization.Read.All" | Out-Null
}

$subscribedSkus = Get-MgSubscribedSku -All
$skuMap = @{}
foreach ($sku in $subscribedSkus) {
    $skuMap[$sku.SkuId.Guid] = $sku.SkuPartNumber
}

$users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,AssignedLicenses"

$assignedLicenseRows = foreach ($user in $users) {
    foreach ($license in $user.AssignedLicenses) {
        $skuId = $license.SkuId.Guid
        [PSCustomObject]@{
            UserId                = $user.Id
            UserPrincipalName     = $user.UserPrincipalName
            DisplayName           = $user.DisplayName
            AssignedSkuId         = $skuId
            AssignedSkuPartNumber = if ($skuMap.ContainsKey($skuId)) { $skuMap[$skuId] } else { "" }
        }
    }
}

$tenantProductRows = foreach ($sku in $subscribedSkus) {
    [PSCustomObject]@{
        SkuId            = $sku.SkuId.Guid
        SkuPartNumber    = $sku.SkuPartNumber
        ConsumedUnits    = $sku.ConsumedUnits
        PrepaidEnabled   = $sku.PrepaidUnits.Enabled
        PrepaidSuspended = $sku.PrepaidUnits.Suspended
        PrepaidWarning   = $sku.PrepaidUnits.Warning
    }
}

$usersCsvPath = Join-Path -Path $OutputDirectory -ChildPath "users-with-assigned-licenses.csv"
$productsCsvPath = Join-Path -Path $OutputDirectory -ChildPath "tenant-products-licenses.csv"

$assignedLicenseRows | Export-Csv -Path $usersCsvPath -NoTypeInformation -Encoding utf8
$tenantProductRows | Export-Csv -Path $productsCsvPath -NoTypeInformation -Encoding utf8

Write-Host "Export complete:"
Write-Host " - $usersCsvPath"
Write-Host " - $productsCsvPath"
