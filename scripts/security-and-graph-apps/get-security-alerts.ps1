#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft Defender XDR security alerts for customer tenants, and optionally resolves alerts
    with a comment.

.DESCRIPTION
    Reads alerts created in the last -Days days in each customer through the Microsoft Graph security
    API (GET /security/alerts_v2) and returns one row per alert with its title, severity, status, service
    source, category and a link to the alert in the Microsoft Defender portal. Resolved alerts are left out
    unless -IncludeResolved is used.

    With -ResolveAlertId and -Apply the script sets the named alerts to resolved
    (PATCH /security/alerts_v2/{id}), records -Classification and adds -Comment
    (POST /security/alerts_v2/{id}/comments). -WhatIf shows what would change.

    The legacy security alerts API (/security/alerts) that the original used retires on 15 October 2026.
    The SecurityAlertSync Azure Function in the functions folder runs the report every 30 minutes and queues
    each new alert once for Power Automate (it keeps the alert IDs it has seen in a state blob).

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER Days
    How many days of alerts to return. Default 7.

.PARAMETER MinimumSeverity
    Lowest severity to return: informational, low (default), medium or high.

.PARAMETER IncludeResolved
    Also return resolved alerts.

.PARAMETER ResolveAlertId
    IDs of alerts to resolve. Use with a single -TenantId and -Apply.

.PARAMETER Classification
    Classification recorded when resolving: unknown, falsePositive, truePositive or
    informationalExpectedActivity.

.PARAMETER Comment
    Comment added to each alert that is resolved.

.PARAMETER Apply
    Resolve the alerts named in -ResolveAlertId. Without it the script only reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-security-alerts.ps1 -AllCustomers -Days 1 -MinimumSeverity medium -OutputPath ./alerts.csv

    Lists open medium and high severity alerts from the last day across all customers.

.EXAMPLE
    ./get-security-alerts.ps1 -TenantId 'contoso.onmicrosoft.com' -ResolveAlertId 'da637551227677560813_-961444813' -Classification informationalExpectedActivity -Comment 'Approved admin activity.' -Apply -WhatIf

    Shows the alert that would be resolved in one customer, without resolving it.

.NOTES
    Replaces the original 2019 method: the AzureAD and AzureRM modules to create a multi-tenant app with a
    99-year client secret and add its service principal to the AdminAgents group (DAP), the legacy Graph
    security alerts API, and Azure Functions v1 timer and HTTP functions with the secret in their code.
    Required GDAP roles: Security Reader (report), Security Operator or Security Administrator (resolve).
    Required partner app permissions: Microsoft Graph delegated SecurityAlert.Read.All (report) and
    SecurityAlert.ReadWrite.All (resolve and comment). SecurityAlert.ReadWrite.All is in the full manifest.

.LINK
    https://gcit.com.au/knowledge-base/manage-office-365-customers-security-alerts-with-sharepoint-microsoft-flow-and-azure-functions/

.LINK
    docs/07-migrating-from-dap-msonline.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateRange(1, 180)]
    [int]$Days = 7,

    [ValidateSet('informational', 'low', 'medium', 'high')]
    [string]$MinimumSeverity = 'low',

    [switch]$IncludeResolved,

    [Parameter(ParameterSetName = 'Tenant')]
    [ValidateNotNullOrEmpty()]
    [string[]]$ResolveAlertId,

    [ValidateSet('unknown', 'falsePositive', 'truePositive', 'informationalExpectedActivity')]
    [string]$Classification,

    [string]$Comment,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $severities = @{ informational = 0; low = 1; medium = 2; high = 3 }

    function ConvertTo-ResultRow {
        param($Customer, $Alert, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId = $Customer.TenantId
            CustomerName     = $Customer.Name
            AlertId          = $Alert.id
            Title            = $Alert.title
            Severity         = $Alert.severity
            Status           = $Alert.status
            Classification   = $Alert.classification
            ServiceSource    = $Alert.serviceSource
            Category         = $Alert.category
            CreatedDateTime  = $Alert.createdDateTime
            IncidentId       = $Alert.incidentId
            AlertWebUrl      = $Alert.alertWebUrl
            Action           = $Action
            Error            = $ErrorMessage
        }
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }
    if ($ResolveAlertId -and $customers.Count -ne 1) {
        throw '-ResolveAlertId needs exactly one -TenantId, because alert IDs belong to one customer.'
    }
    $since = [datetime]::UtcNow.AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $minimum = $severities[$MinimumSeverity]

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            if ($ResolveAlertId) {
                foreach ($alertId in $ResolveAlertId) {
                    $alert = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "security/alerts_v2/$alertId"
                    if ([string]$alert.status -eq 'resolved') {
                        $action = 'AlreadyResolved'
                    }
                    elseif (-not $Apply) {
                        $action = 'WouldResolve'
                    }
                    elseif ($PSCmdlet.ShouldProcess("alert $alertId in $($customer.Name)", 'Resolve alert')) {
                        $body = @{ status = 'resolved' }
                        if ($Classification) { $body.classification = $Classification }
                        $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method PATCH -Uri "security/alerts_v2/$alertId" -Body $body -Confirm:$false
                        if ($Comment) {
                            $commentBody = @{ '@odata.type' = 'microsoft.graph.security.alertComment'; comment = $Comment }
                            $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri "security/alerts_v2/$alertId/comments" -Body $commentBody -Confirm:$false
                        }
                        $alert = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "security/alerts_v2/$alertId"
                        $action = if ([string]$alert.status -eq 'resolved') { 'Resolved' } else { 'ResolveNotConfirmed' }
                    }
                    else {
                        $action = 'WhatIf'
                    }
                    $row = ConvertTo-ResultRow -Customer $customer -Alert $alert -Action $action
                    $results.Add($row)
                    $row
                }
                continue
            }

            $alerts = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "security/alerts_v2?`$filter=createdDateTime ge $since")
            foreach ($alert in $alerts) {
                if (-not $IncludeResolved -and [string]$alert.status -eq 'resolved') { continue }
                $severity = $severities[[string]$alert.severity]
                if ($null -ne $severity -and $severity -lt $minimum) { continue }
                $row = ConvertTo-ResultRow -Customer $customer -Alert $alert -Action 'Reported'
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
