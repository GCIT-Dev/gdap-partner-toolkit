# Azure Functions v4 timer function, PowerShell 7.4 or later (7.6 recommended).
# Report only: queues one message per NEW external forward. Removal stays an interactive task.
# As in the original (which kept forwards in Azure Table storage and only alerted on new ones), the
# forwards seen on the last run are kept in a state blob, so each forward is queued once. Customers that
# fail in a run keep their previous state, so they don't re-alert on the next successful run.
# Runs scripts/get-external-forwarding.ps1 from the function app root. See profile.ps1 for the publish layout and app settings.
# Replaces the Azure Functions v1 function in the original article: no stored password, AES key file,
# storage account key or MSOnline module.
param($Timer, $previousState)

$ErrorActionPreference = 'Stop'
if ($Timer.IsPastDue) {
    Write-Warning 'The timer trigger is running later than scheduled.'
}

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath 'scripts', 'get-external-forwarding.ps1'
$scriptParams = @{}
if ($env:MSPGDAP_TENANT_IDS) {
    $scriptParams.TenantId = @($env:MSPGDAP_TENANT_IDS -split '[,;\s]+' | Where-Object { $_ })
}
else {
    $scriptParams.AllCustomers = $true
}

# Optional app-only Exchange (Pattern A in docs/08). Set MSPGDAP_EXO_APP_ID, MSPGDAP_AUTOMATION_KEYVAULT_NAME and
# MSPGDAP_AUTOMATION_CERT_NAME to connect as the separate automation app prepared with
# enable-exchange-app-access.ps1. Without them Exchange runs as the delegated automation account.
$exchangeCertificate = $null
if ($env:MSPGDAP_EXO_APP_ID) {
    $pfxBase64 = Get-AzKeyVaultSecret -VaultName $env:MSPGDAP_AUTOMATION_KEYVAULT_NAME -Name $env:MSPGDAP_AUTOMATION_CERT_NAME -AsPlainText
    # Windows: keep the PFX's CSP provider (Exchange app-only does not support CNG). Linux: in memory only.
    $storageFlags = if ($IsWindows) {
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet
    }
    else {
        [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
    }
    $exchangeCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String($pfxBase64), [string]::Empty, $storageFlags)
    Remove-Variable -Name pfxBase64
    $scriptParams.ExchangeAppId = $env:MSPGDAP_EXO_APP_ID
    $scriptParams.ExchangeCertificate = $exchangeCertificate
}

try {
    $rows = @(& $scriptPath @scriptParams)
}
finally {
    if ($exchangeCertificate) {
        $exchangeCertificate.Dispose()
    }
}

$failed = @($rows | Where-Object { $_.Error })
foreach ($failure in $failed) {
    Write-Warning "Customer $($failure.CustomerTenantId) ($($failure.CustomerName)): $($failure.Error)"
}

function Get-ForwardKey {
    param($Row)
    '{0}|{1}|{2}|{3}|{4}' -f $Row.CustomerTenantId, $Row.Mailbox, $Row.Source, $Row.InboxRuleName, $Row.ExternalRecipient
}

$previousKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
if ($previousState) {
    $stateText = if ($previousState -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($previousState) } else { [string]$previousState }
    try {
        foreach ($key in @(ConvertFrom-Json -InputObject $stateText)) { $null = $previousKeys.Add([string]$key) }
    }
    catch {
        Write-Warning "The previous state blob could not be read, so every forward is treated as new: $($_.Exception.Message)"
    }
}

$current = @($rows | Where-Object { -not $_.Error -and $_.Action -ne 'Failed' })
$newRows = @($current | Where-Object { -not $previousKeys.Contains((Get-ForwardKey -Row $_)) })

$failedTenants = @($failed | ForEach-Object { '{0}|' -f $_.CustomerTenantId })
$carried = @($previousKeys | Where-Object { $key = $_; @($failedTenants | Where-Object { $key.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0 })
$state = @(@($current | ForEach-Object { Get-ForwardKey -Row $_ }) + $carried | Sort-Object -Unique)
Push-OutputBinding -Name 'state' -Value (ConvertTo-Json -InputObject $state -Compress)

$messages = @($newRows | ForEach-Object { ConvertTo-Json -InputObject $_ -Compress -Depth 5 })
if ($messages.Count -gt 0) {
    Push-OutputBinding -Name 'alerts' -Value $messages
}

Write-Information -MessageData ('{0} external forwards, {1} new, {2} failed customers.' -f $current.Count, $newRows.Count, $failed.Count) -InformationAction Continue
