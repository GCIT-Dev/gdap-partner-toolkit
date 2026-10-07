# Azure Functions v4 timer function, PowerShell 7.4 or later (7.6 recommended).
# Queues one message per SKU with unused licences, for Power Automate to email.
# Runs scripts/get-unused-licences.ps1 from the function app root. See profile.ps1 for the publish layout and app settings.
# Replaces the Azure Functions v1 function in the original article: no stored password, AES key file,
# storage account key or MSOnline module.
param($Timer)

$ErrorActionPreference = 'Stop'
if ($Timer.IsPastDue) {
    Write-Warning 'The timer trigger is running later than scheduled.'
}

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath 'scripts', 'get-unused-licences.ps1'
$scriptParams = @{}
if ($env:MSPGDAP_TENANT_IDS) {
    $scriptParams.TenantId = @($env:MSPGDAP_TENANT_IDS -split '[,;\s]+' | Where-Object { $_ })
}
else {
    $scriptParams.AllCustomers = $true
}

$rows = @(& $scriptPath @scriptParams)

$failed = @($rows | Where-Object { $_.Error })
foreach ($failure in $failed) {
    Write-Warning "Customer $($failure.CustomerTenantId) ($($failure.CustomerName)): $($failure.Error)"
}

$messages = @($rows | Where-Object { -not $_.Error -and $_.Action -ne 'Failed' } | ForEach-Object { ConvertTo-Json -InputObject $_ -Compress -Depth 5 })
if ($messages.Count -gt 0) {
    Push-OutputBinding -Name 'alerts' -Value $messages
}

Write-Information -MessageData ('{0} rows, {1} failed customers.' -f $rows.Count, $failed.Count) -InformationAction Continue
