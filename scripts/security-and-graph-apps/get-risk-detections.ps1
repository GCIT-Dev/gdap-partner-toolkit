#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft Entra ID Protection risk detections, and optionally risky users, for customer tenants.

.DESCRIPTION
    Reads risk detections from the last -Days days in each customer through Microsoft Graph
    (GET /identityProtection/riskDetections) and returns one row per detection with the user, risk type,
    level, state, IP address and location.

    -IncludeRiskyUsers also returns users currently at risk (GET /identityProtection/riskyUsers) as rows
    with RecordType "RiskyUser".

    The risk detection API needs Microsoft Entra ID P1 or P2 in the customer tenant, and full detail needs
    P2. Customers without it return an error row and the script carries on.

    The output is flat, so it can be sent to a ticketing system, a SIEM or a SharePoint list through Power
    Automate. SharePoint "Alert me" notifications, which the original relied on, are being retired by
    Microsoft.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER Days
    How many days of risk detections to return. Default 7.

.PARAMETER MinimumRiskLevel
    Lowest risk level to return: low (default), medium or high.

.PARAMETER IncludeRiskyUsers
    Also return users whose risk state is atRisk.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-risk-detections.ps1 -AllCustomers -Days 1 -MinimumRiskLevel medium -OutputPath ./risk.csv

    Lists medium and high risk detections from the last day across all customers.

.EXAMPLE
    ./get-risk-detections.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeRiskyUsers

    Lists recent risk detections and users at risk in one customer.

.NOTES
    Replaces the original 2019 method: the AzureAD and AzureRM modules to create an app and add its service
    principal to the AdminAgents group (DAP), a client secret in the script, the v1 token endpoint with
    resource=, and the removed beta identityRiskEvents API.
    Required GDAP roles: Security Reader (or Global Reader, Security Operator).
    Required partner app permissions: Microsoft Graph delegated IdentityRiskEvent.Read.All and, for
    -IncludeRiskyUsers, IdentityRiskyUser.Read.All. Both are in the full manifest.
    Microsoft Learn: https://learn.microsoft.com/en-us/graph/api/riskdetection-list

.LINK
    https://gcit.com.au/knowledge-base/sync-azure-active-directory-risk-events-with-a-sharepoint-list-for-all-customer-tenants/

.LINK
    docs/04-preconsent-customers.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateRange(1, 90)]
    [int]$Days = 7,

    [ValidateSet('low', 'medium', 'high')]
    [string]$MinimumRiskLevel = 'low',

    [switch]$IncludeRiskyUsers,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $levels = @{ low = 1; medium = 2; high = 3 }

    function Test-RiskLevel {
        # 'hidden' and 'none' have no rank. They are kept only when every level is wanted.
        param([string]$Level)
        $value = $levels[$Level]
        if ($null -eq $value) { return $minimum -eq 1 }
        $value -ge $minimum
    }

    function ConvertTo-ResultRow {
        param($Customer, $RecordType, $Item, $ErrorMessage)
        $location = if ($Item.location) {
            (@($Item.location.city, $Item.location.state, $Item.location.countryOrRegion) | Where-Object { $_ }) -join ', '
        }
        else {
            $null
        }
        [pscustomobject][ordered]@{
            CustomerTenantId  = $Customer.TenantId
            CustomerName      = $Customer.Name
            RecordType        = $RecordType
            Id                = $Item.id
            UserPrincipalName = $Item.userPrincipalName
            UserDisplayName   = $Item.userDisplayName
            RiskEventType     = $Item.riskEventType
            RiskLevel         = $Item.riskLevel
            RiskState         = $Item.riskState
            RiskDetail        = $Item.riskDetail
            DateTime          = if ($Item.detectedDateTime) { $Item.detectedDateTime } else { $Item.riskLastUpdatedDateTime }
            IpAddress         = $Item.ipAddress
            Location          = $location
            Activity          = $Item.activity
            Source            = $Item.source
            Error             = $ErrorMessage
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
    $since = [datetime]::UtcNow.AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $minimum = $levels[$MinimumRiskLevel]

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $detections = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "identityProtection/riskDetections?`$filter=detectedDateTime ge $since")
            foreach ($detection in $detections) {
                if (-not (Test-RiskLevel -Level $detection.riskLevel)) { continue }
                $row = ConvertTo-ResultRow -Customer $customer -RecordType 'RiskDetection' -Item $detection
                $results.Add($row)
                $row
            }

            if ($IncludeRiskyUsers) {
                $riskyUsers = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "identityProtection/riskyUsers?`$filter=riskState eq 'atRisk'")
                foreach ($riskyUser in $riskyUsers) {
                    if (-not (Test-RiskLevel -Level $riskyUser.riskLevel)) { continue }
                    $row = ConvertTo-ResultRow -Customer $customer -RecordType 'RiskyUser' -Item $riskyUser
                    $results.Add($row)
                    $row
                }
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -RecordType 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
