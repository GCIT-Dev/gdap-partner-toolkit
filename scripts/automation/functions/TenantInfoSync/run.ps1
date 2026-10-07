# Azure Functions v4 timer function, PowerShell 7.4 or later (7.6 recommended).
# Collects tenant domains, licences and licensed users and, with MSPGDAP_APPLY_CHANGES=true and the IT Glue settings, updates IT Glue.
# Runs scripts/sync-tenant-info-itglue.ps1 from the function app root. See profile.ps1 for the publish layout and app settings.
# Replaces the Azure Functions v1 function in the original article: no stored password, AES key file,
# storage account key or MSOnline module.
param($Timer)

$ErrorActionPreference = 'Stop'
if ($Timer.IsPastDue) {
    Write-Warning 'The timer trigger is running later than scheduled.'
}

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath 'scripts', 'sync-tenant-info-itglue.ps1'
$scriptParams = @{}
if ($env:MSPGDAP_TENANT_IDS) {
    $scriptParams.TenantId = @($env:MSPGDAP_TENANT_IDS -split '[,;\s]+' | Where-Object { $_ })
}
else {
    $scriptParams.AllCustomers = $true
}
if ($env:MSPGDAP_ITGLUE_ASSET_TYPE_ID) {
    # The IT Glue API key is a secret in the token vault registered in profile.ps1.
    $scriptParams.ITGlueFlexibleAssetTypeId = [long]$env:MSPGDAP_ITGLUE_ASSET_TYPE_ID
    $scriptParams.ITGlueVaultName = 'MspGdapAutomation'
    if ($env:MSPGDAP_ITGLUE_KEY_SECRET_NAME) { $scriptParams.ITGlueApiKeySecretName = $env:MSPGDAP_ITGLUE_KEY_SECRET_NAME }
    if ($env:MSPGDAP_ITGLUE_BASE_URI) { $scriptParams.ITGlueBaseUri = $env:MSPGDAP_ITGLUE_BASE_URI }
    if ($env:MSPGDAP_ITGLUE_MAP_PATH) { $scriptParams.ITGlueOrganizationMapPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath $env:MSPGDAP_ITGLUE_MAP_PATH }
}
if ($env:MSPGDAP_APPLY_CHANGES -eq 'true') {
    $scriptParams.Apply = $true
}

$rows = @(& $scriptPath @scriptParams)

$failed = @($rows | Where-Object { $_.Error })
foreach ($failure in $failed) {
    Write-Warning "Customer $($failure.CustomerTenantId) ($($failure.CustomerName)): $($failure.Error)"
}

Push-OutputBinding -Name 'report' -Value (ConvertTo-Json -InputObject $rows -Depth 5)

Write-Information -MessageData ('{0} rows, {1} failed customers.' -f $rows.Count, $failed.Count) -InformationAction Continue
