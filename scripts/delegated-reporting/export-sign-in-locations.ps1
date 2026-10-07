#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports where users in GDAP customers sign in from, using Microsoft Entra sign-in logs.
.DESCRIPTION
    Reads the Microsoft Entra sign-in logs of each customer through Microsoft Graph
    (GET /auditLogs/signIns) with the technician's delegated GDAP token. Every sign-in record
    already includes the city, state and country or region, so no third-party IP lookup service
    is needed and no customer data leaves Microsoft.

    By default the script returns one row per user, IP address and location, with sign-in counts,
    the client apps and applications used, and the first and last time seen. Use -Detailed for one
    row per sign-in.

    The list API returns interactive sign-ins only. Reading sign-in logs through Microsoft Graph
    needs a Microsoft Entra ID P1 or P2 licence in the customer. A customer without one is reported
    as Failed with the error Microsoft returns. Sign-in logs are kept for 30 days.

    The original also reported the ISP and the raw user agent string. Microsoft Graph v1.0
    sign-in records don't include either, so the Devices column lists the operating system and
    browser that Microsoft Entra recorded for each sign-in instead.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER Days
    How many days back to read, from 1 to 30.
.PARAMETER Detailed
    Return one row per sign-in instead of one row per user, IP address and location.
.PARAMETER MaxPages
    Stop after this many pages of 1,000 sign-ins per customer. 0 (default) reads every page.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./export-sign-in-locations.ps1 -AllCustomers -Days 7 -OutputPath ./sign-in-locations.csv

    Summarises a week of sign-in locations for every GDAP customer.
.EXAMPLE
    ./export-sign-in-locations.ps1 -TenantId 'contoso.onmicrosoft.com' -Detailed | Where-Object CountryOrRegion -ne 'AU'

    Lists every sign-in from outside Australia in the last 30 days for one customer.
.NOTES
    Replaces the original 2017 method: Search-UnifiedAuditLog over basic authentication remote
    PowerShell (New-PSSession to outlook.office365.com and ?DelegatedOrg=), MSOnline and DAP
    (Get-MsolPartnerContract, Get-MsolCompanyInformation, Get-MsolDomain), and IP lookups sent to
    ip-api.com over plain HTTP.
    Required GDAP roles: Reports Reader, Security Reader or Global Reader (Security Operator and
    Security Administrator also work).
    Required partner app permissions: Microsoft Graph AuditLog.Read.All (delegated).
.LINK
    https://gcit.com.au/knowledge-base/export-list-locations-office-365-users-logging/
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    https://learn.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [ValidateRange(1, 30)]
    [int]$Days = 30,

    [switch]$Detailed,

    [ValidateRange(0, 10000)]
    [int]$MaxPages = 0,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @(
        'Status', 'UserPrincipalName', 'IpAddress', 'City', 'State', 'CountryOrRegion', 'SignInCount',
        'FailedCount', 'ClientApps', 'Applications', 'Devices', 'FirstSeen', 'LastSeen', 'Detail'
    )

    function ConvertTo-ResultRow {
        param(
            [Parameter(Mandatory)][object]$Customer,
            [System.Collections.IDictionary]$Values = @{}
        )
        $row = [ordered]@{
            CustomerTenantId = [string]$Customer.TenantId
            CustomerName     = [string]$Customer.DisplayName
        }
        foreach ($column in $resultColumns) {
            $row[$column] = if ($Values.Contains($column)) { $Values[$column] } else { $null }
        }
        [pscustomobject]$row
    }

    function Get-TargetCustomer {
        param([switch]$All, [string[]]$Requested)
        if ($All) {
            return @(Get-MspCustomer -IncludeGdapStatus)
        }
        foreach ($id in $Requested) {
            $match = $null
            try {
                $match = Get-MspCustomer -TenantId $id -IncludeGdapStatus | Select-Object -First 1
            }
            catch {
                Write-Warning "Could not look up '$id' in your customer list: $($_.Exception.Message)"
            }
            if ($match) { $match } else { [pscustomobject]@{ TenantId = $id; DisplayName = $id } }
        }
    }

    function ConvertTo-UtcText {
        param([object]$Value)
        if ($null -eq $Value -or [string]::IsNullOrEmpty([string]$Value)) { return $null }
        ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    }

    function Format-DeviceDetail {
        # Operating system and browser from deviceDetail, for example 'Windows 10 / Edge 128.0'.
        param([object]$SignIn)
        $parts = @([string]$SignIn.deviceDetail.operatingSystem, [string]$SignIn.deviceDetail.browser) | Where-Object { $_ }
        if ($parts) { $parts -join ' / ' } else { $null }
    }
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $since = [DateTime]::UtcNow.AddDays(-$Days).ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
    $uri = "v1.0/auditLogs/signIns?`$filter=createdDateTime ge $since&`$top=1000"

    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        $customerRows = [System.Collections.Generic.List[object]]::new()
        try {
            Write-Verbose "Reading $Days days of sign-ins for $($customer.DisplayName)"
            $signIns = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri $uri -MaxPages $MaxPages)

            if ($Detailed) {
                foreach ($signIn in $signIns) {
                    $failed = [int]($signIn.status.errorCode -ne 0)
                    $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{
                                Status            = 'Succeeded'
                                UserPrincipalName = [string]$signIn.userPrincipalName
                                IpAddress         = [string]$signIn.ipAddress
                                City              = [string]$signIn.location.city
                                State             = [string]$signIn.location.state
                                CountryOrRegion   = [string]$signIn.location.countryOrRegion
                                SignInCount       = 1
                                FailedCount       = $failed
                                ClientApps        = [string]$signIn.clientAppUsed
                                Devices           = Format-DeviceDetail -SignIn $signIn
                                Applications      = [string]$signIn.appDisplayName
                                FirstSeen         = ConvertTo-UtcText -Value $signIn.createdDateTime
                                LastSeen          = ConvertTo-UtcText -Value $signIn.createdDateTime
                            }))
                }
            }
            else {
                $groups = $signIns | Group-Object -Property { '{0}|{1}|{2}|{3}|{4}' -f $_.userPrincipalName, $_.ipAddress, $_.location.city, $_.location.state, $_.location.countryOrRegion }
                foreach ($group in $groups) {
                    $first = $group.Group[0]
                    $times = @($group.Group | ForEach-Object { ConvertTo-UtcText -Value $_.createdDateTime } | Where-Object { $_ } | Sort-Object)
                    $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{
                                Status            = 'Succeeded'
                                UserPrincipalName = [string]$first.userPrincipalName
                                IpAddress         = [string]$first.ipAddress
                                City              = [string]$first.location.city
                                State             = [string]$first.location.state
                                CountryOrRegion   = [string]$first.location.countryOrRegion
                                SignInCount       = $group.Count
                                FailedCount       = @($group.Group | Where-Object { $_.status.errorCode -ne 0 }).Count
                                ClientApps        = (@($group.Group | ForEach-Object { [string]$_.clientAppUsed } | Where-Object { $_ } | Sort-Object -Unique)) -join ', '
                                Devices           = (@($group.Group | ForEach-Object { Format-DeviceDetail -SignIn $_ } | Where-Object { $_ } | Sort-Object -Unique)) -join ', '
                                Applications      = (@($group.Group | ForEach-Object { [string]$_.appDisplayName } | Where-Object { $_ } | Sort-Object -Unique)) -join ', '
                                FirstSeen         = $times[0]
                                LastSeen          = $times[-1]
                            }))
                }
            }

            if ($customerRows.Count -eq 0) {
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; SignInCount = 0; Detail = "No interactive sign-ins in the last $Days days." }))
            }
        }
        catch {
            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; Detail = $_.Exception.Message }))
        }

        foreach ($row in $customerRows) {
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
        Write-Verbose "Saved $($results.Count) rows to $OutputPath"
    }
}
