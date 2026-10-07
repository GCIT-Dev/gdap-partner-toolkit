#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Pulls a user's unified audit log activity, with sign-in locations, to help decide whether a Microsoft 365
    account has been breached.

.DESCRIPTION
    For one user in each customer named in -TenantId, the script connects to Exchange Online through MspGdap
    and runs Search-UnifiedAuditLog for the last -Days days (paging with ReturnLargeSet, up to -ResultLimit
    records). It returns one row per audit record with the time, workload, operation, result, client IP
    and the full AuditData JSON (the original saved the raw records for later analysis).

    -IncludeSignInLocations adds the city and country for each client IP from the user's Microsoft Entra
    sign-in logs (GET /auditLogs/signIns), so no investigation data is sent to a third-party geolocation
    service. Sign-in logs need Microsoft Entra ID P1 or P2 and are kept for 30 days.

    Group the output to spot unusual countries and operations, for example:
    ... | Group-Object -Property Location | Sort-Object -Property Count -Descending
    ... | Where-Object Operation -match 'InboxRule|Set-Mailbox|Add-MailboxPermission'

    Auditing is on by default for most tenants, but check it. The script reports UnifiedAuditLogIngestionEnabled
    and warns if it is off. Audit (Standard) keeps records for 180 days. Audit (Premium) keeps them for a year.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.

.PARAMETER UserPrincipalName
    The user to investigate.

.PARAMETER Days
    How many days of audit records to return. Default 10.

.PARAMETER ResultLimit
    Most audit records to return per customer. Default 50000, the ReturnLargeSet maximum.

.PARAMETER IncludeSignInLocations
    Add city and country for each client IP from the Microsoft Entra sign-in logs.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-user-breach-indicators.ps1 -TenantId 'contoso.onmicrosoft.com' -UserPrincipalName 'jane@contoso.com' -IncludeSignInLocations -OutputPath ./jane-audit.csv

    Exports ten days of audit activity for one user with sign-in locations.

.EXAMPLE
    ./get-user-breach-indicators.ps1 -TenantId 'contoso.onmicrosoft.com' -UserPrincipalName 'jane@contoso.com' -Days 30 | Group-Object -Property Operation | Sort-Object -Property Count -Descending

    Summarises what the user did in the last 30 days, by operation.

.NOTES
    Replaces the original 2020 method: Connect-ExchangeOnline with the investigator's own account and
    Search-UnifiedAuditLog limited to 5000 records, with client IPs sent to the free ip-api.com service over
    plain HTTP, and references to the retired Security and Compliance Center (protection.office.com). The
    audit search itself is unchanged.
    Required GDAP roles: Compliance Administrator or Exchange Administrator (Search-UnifiedAuditLog needs the
    View-Only Audit Logs or Audit Logs role in Exchange Online), plus Reports Reader or Security Reader for
    -IncludeSignInLocations.
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage, Microsoft Graph
    delegated AuditLog.Read.All for sign-in logs (in the full manifest), and User.Read (organisation name
    lookup).

.LINK
    https://gcit.com.au/knowledge-base/how-to-detect-a-breach-in-microsoft-365/

.LINK
    docs/05-exchange-access.md
#>
[CmdletBinding()]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory)]
    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string]$UserPrincipalName,

    [ValidateRange(1, 180)]
    [int]$Days = 10,

    [ValidateRange(1, 50000)]
    [int]$ResultLimit = 50000,

    [switch]$IncludeSignInLocations,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function Get-NormalisedIp {
        param([string]$Address)
        if (-not $Address) { return $null }
        if ($Address -match '^\[(.+)\](:\d+)?$') { return $Matches[1] }
        if ($Address -match '^(\d{1,3}(\.\d{1,3}){3}):\d+$') { return $Matches[1] }
        $Address
    }

    function ConvertTo-ResultRow {
        param($Customer, $Record, $Data, $Location, $AuditEnabled, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId    = $Customer.TenantId
            CustomerName        = $Customer.Name
            UserPrincipalName   = $UserPrincipalName
            CreationTime        = if ($Data.CreationTime) { $Data.CreationTime } else { $Record.CreationDate }
            Workload            = $Data.Workload
            Operation           = if ($Data.Operation) { $Data.Operation } else { $Record.Operations }
            ResultStatus        = $Data.ResultStatus
            ClientIP            = Get-NormalisedIp -Address ([string]$Data.ClientIP)
            Location            = $Location
            UserAgent           = $Data.UserAgent
            ObjectId            = $Data.ObjectId
            AuditLogEnabled     = $AuditEnabled
            AuditData           = $Record.AuditData
            Error               = $ErrorMessage
        }
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    $end = [datetime]::UtcNow
    $start = $end.AddDays(-$Days)

    foreach ($customer in $customers) {
        try {
            $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
            $customer.Name = $organisation.displayName
            if ($organisation.id) { $customer.TenantId = $organisation.id }

            $locations = @{}
            if ($IncludeSignInLocations) {
                $since = $start.ToString('yyyy-MM-ddTHH:mm:ssZ')
                $escapedUpn = $UserPrincipalName.Replace("'", "''")
                Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "auditLogs/signIns?`$filter=userPrincipalName eq '$escapedUpn' and createdDateTime ge $since" |
                    ForEach-Object {
                        $ip = [string]$_.ipAddress
                        if ($ip -and -not $locations.ContainsKey($ip)) {
                            $locations[$ip] = (@($_.location.city, $_.location.state, $_.location.countryOrRegion) | Where-Object { $_ }) -join ', '
                        }
                    }
            }

            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId
            $customerRows = [System.Collections.Generic.List[object]]::new()
            try {
                $auditEnabled = [bool](Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled
                if (-not $auditEnabled) {
                    Write-Warning "Unified audit log ingestion is off in $($customer.Name). Turn it on with Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled `$true."
                }

                $sessionId = [guid]::NewGuid().ToString()
                $seen = [System.Collections.Generic.HashSet[string]]::new()
                do {
                    $batch = @(Search-UnifiedAuditLog -StartDate $start -EndDate $end -UserIds $UserPrincipalName -SessionId $sessionId -SessionCommand ReturnLargeSet -ResultSize 5000)
                    foreach ($record in $batch) {
                        if ($customerRows.Count -ge $ResultLimit) { break }
                        if (-not $seen.Add([string]$record.Identity)) { continue }
                        $data = $null
                        try { $data = $record.AuditData | ConvertFrom-Json -ErrorAction Stop } catch { Write-Verbose "Could not parse AuditData for $($record.Identity)." }
                        $ip = Get-NormalisedIp -Address ([string]$data.ClientIP)
                        $location = if ($ip) { $locations[$ip] } else { $null }
                        $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Record $record -Data $data -Location $location -AuditEnabled $auditEnabled))
                    }
                    $total = if ($batch.Count -gt 0) { [int]$batch[0].ResultCount } else { 0 }
                } while ($batch.Count -gt 0 -and $customerRows.Count -lt $ResultLimit -and $seen.Count -lt $total)
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            foreach ($row in $customerRows) {
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
