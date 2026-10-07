# Azure Functions v4 timer function, PowerShell 7.4 or later (7.6 recommended).
# Reports, and with MSPGDAP_APPLY_CHANGES=true keeps current, the display name warning mail flow rule.
# Runs scripts/set-display-name-spoof-rule.ps1 from the function app root. See profile.ps1 for the publish layout and app settings.
# Replaces the Azure Functions v1 function in the original article: no stored password, AES key file,
# storage account key or MSOnline module.
param($Timer)

$ErrorActionPreference = 'Stop'
if ($Timer.IsPastDue) {
    Write-Warning 'The timer trigger is running later than scheduled.'
}

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath 'scripts', 'set-display-name-spoof-rule.ps1'
$scriptParams = @{}
if ($env:MSPGDAP_TENANT_IDS) {
    $scriptParams.TenantId = @($env:MSPGDAP_TENANT_IDS -split '[,;\s]+' | Where-Object { $_ })
}
else {
    $scriptParams.AllCustomers = $true
}
if ($env:MSPGDAP_APPLY_CHANGES -eq 'true') {
    $scriptParams.Apply = $true
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

Push-OutputBinding -Name 'report' -Value (ConvertTo-Json -InputObject $rows -Depth 5)

Write-Information -MessageData ('{0} rows, {1} failed customers.' -f $rows.Count, $failed.Count) -InformationAction Continue
