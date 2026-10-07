# Azure Functions v4 timer function, PowerShell 7.4 or later (7.6 recommended).
# Queues one message per NEW open Microsoft Defender XDR alert. Resolving alerts stays an interactive task.
# The timer runs every 30 minutes over a one-day window, so the alert IDs seen on the last run are kept in
# a state blob and each alert is queued once (the original synced alerts to a SharePoint list by ID).
# Customers that fail in a run keep their previous state, so they don't re-alert on the next run.
# Runs scripts/get-security-alerts.ps1 from the function app root. See profile.ps1 for the publish layout and app settings.
# Replaces the Azure Functions v1 function in the original article: no stored password, AES key file,
# storage account key or MSOnline module.
param($Timer, $previousState)

$ErrorActionPreference = 'Stop'
if ($Timer.IsPastDue) {
    Write-Warning 'The timer trigger is running later than scheduled.'
}

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath 'scripts', 'get-security-alerts.ps1'
$scriptParams = @{}
if ($env:MSPGDAP_TENANT_IDS) {
    $scriptParams.TenantId = @($env:MSPGDAP_TENANT_IDS -split '[,;\s]+' | Where-Object { $_ })
}
else {
    $scriptParams.AllCustomers = $true
}
$scriptParams.Days = 1

$rows = @(& $scriptPath @scriptParams)

$failed = @($rows | Where-Object { $_.Error })
foreach ($failure in $failed) {
    Write-Warning "Customer $($failure.CustomerTenantId) ($($failure.CustomerName)): $($failure.Error)"
}

$previousKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
if ($previousState) {
    $stateText = if ($previousState -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($previousState) } else { [string]$previousState }
    try {
        foreach ($key in @(ConvertFrom-Json -InputObject $stateText)) { $null = $previousKeys.Add([string]$key) }
    }
    catch {
        Write-Warning "The previous state blob could not be read, so every alert is treated as new: $($_.Exception.Message)"
    }
}

$current = @($rows | Where-Object { -not $_.Error -and $_.Action -ne 'Failed' })
$newRows = @($current | Where-Object { -not $previousKeys.Contains(('{0}|{1}' -f $_.CustomerTenantId, $_.AlertId)) })

$failedTenants = @($failed | ForEach-Object { '{0}|' -f $_.CustomerTenantId })
$carried = @($previousKeys | Where-Object { $key = $_; @($failedTenants | Where-Object { $key.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0 })
$state = @(@($current | ForEach-Object { '{0}|{1}' -f $_.CustomerTenantId, $_.AlertId }) + $carried | Sort-Object -Unique)
Push-OutputBinding -Name 'state' -Value (ConvertTo-Json -InputObject $state -Compress)

$messages = @($newRows | ForEach-Object { ConvertTo-Json -InputObject $_ -Compress -Depth 5 })
if ($messages.Count -gt 0) {
    Push-OutputBinding -Name 'alerts' -Value $messages
}

Write-Information -MessageData ('{0} open alerts, {1} new, {2} failed customers.' -f $current.Count, $newRows.Count, $failed.Count) -InformationAction Continue
