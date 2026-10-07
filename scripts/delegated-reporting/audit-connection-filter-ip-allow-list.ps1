#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Audits connection filter IP Allow List entries in GDAP customers and optionally removes chosen entries.
.DESCRIPTION
    The original article added a shared SendGrid sending IP address to the connection filter IP
    Allow List of every customer, so that all mail from that address skipped spam filtering. That
    lets any sender on the same shared IP address, including spammers, bypass filtering in every
    customer. This script does the safe opposite: it finds those entries so they can be removed.

    For each customer it connects to Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline), reads every connection filter policy
    (Get-HostedConnectionFilterPolicy) and reports each IP Allow List entry, plus whether the
    policy uses the Microsoft safe list.

    Without -Apply the script only reports. With -Apply and -RemoveIpAddress it removes those
    entries with Set-HostedConnectionFilterPolicy -IPAllowList @{ Remove = ... } and reads the
    policy back. It never adds entries.

    To stop legitimate bulk mail landing in junk, have the sending service authenticate the sender
    domain (SPF, DKIM with the service's CNAME records, and DMARC alignment) rather than allow-listing
    its IP addresses.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER RemoveIpAddress
    IP addresses or CIDR ranges to remove from the IP Allow List, written exactly as the policy lists
    them. Used only with -Apply.
.PARAMETER Apply
    Remove the -RemoveIpAddress entries. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./audit-connection-filter-ip-allow-list.ps1 -AllCustomers -OutputPath ./ip-allow-list.csv

    Lists every IP Allow List entry in every GDAP customer.
.EXAMPLE
    ./audit-connection-filter-ip-allow-list.ps1 -TenantId 'contoso.onmicrosoft.com' -RemoveIpAddress '192.0.2.10' -Apply -WhatIf

    Shows whether the entry would be removed from one customer, without changing anything.
.NOTES
    Replaces the original 2017 method: MSOnline and DAP (Connect-MsolService, Get-MsolPartnerContract,
    Get-MsolDomain) with basic authentication remote PowerShell (Invoke-Command to
    ps.outlook.com/powershell-liveid?DelegatedOrg=) that added a shared sending IP address to
    Set-HostedConnectionFilterPolicy -IPAllowList in every customer. That unsafe change is not
    replicated. This script only reports and removes.
    Required GDAP roles: Security Reader or Global Reader (report), Security Administrator or
    Exchange Administrator (remove).
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/stop-sendgrid-emails-going-junk-office-365-users/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    https://learn.microsoft.com/en-us/defender-office-365/connection-filter-policies-configure
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant', SupportsShouldProcess, ConfirmImpact = 'Medium')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [ValidatePattern('^[0-9A-Fa-f:.\-/]+$')]
    [string[]]$RemoveIpAddress = @(),

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'PolicyName', 'EnableSafeList', 'IPAllowListEntry', 'MarkedForRemoval', 'Action', 'Detail')

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
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    if ($Apply -and $RemoveIpAddress.Count -eq 0) {
        Write-Warning '-Apply has nothing to do without -RemoveIpAddress. Reporting only.'
    }

    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        $customerRows = [System.Collections.Generic.List[object]]::new()
        try {
            Write-Verbose "Checking connection filter policies for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            foreach ($policy in @(Get-HostedConnectionFilterPolicy)) {
                $policyName = [string]$policy.Identity
                $entries = @($policy.IPAllowList | ForEach-Object { [string]$_ } | Where-Object { $_ })
                $toRemove = @($entries | Where-Object { $RemoveIpAddress -contains $_ })
                $action = 'ReportOnly'
                $detail = $null
                $after = $entries

                if ($Apply -and $toRemove.Count -gt 0) {
                    if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) connection filter policy $policyName", "Remove $($toRemove -join ', ') from the IP Allow List")) {
                        try {
                            Set-HostedConnectionFilterPolicy -Identity $policyName -IPAllowList @{ Remove = $toRemove }
                            $after = @((Get-HostedConnectionFilterPolicy -Identity $policyName).IPAllowList | ForEach-Object { [string]$_ })
                            $action = if (@($toRemove | Where-Object { $after -contains $_ }).Count -eq 0) { 'Removed' } else { 'ChangeNotConfirmed' }
                        }
                        catch {
                            $action = 'Failed'
                            $detail = $_.Exception.Message
                        }
                    }
                    else {
                        $action = 'WhatIf'
                    }
                }

                if ($entries.Count -eq 0) {
                    $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; PolicyName = $policyName; EnableSafeList = [bool]$policy.EnableSafeList; Action = 'None'; Detail = 'IP Allow List is empty.' }))
                    continue
                }
                foreach ($entry in $entries) {
                    $marked = $toRemove -contains $entry
                    $entryAction = if ($marked) { $action } else { 'ReportOnly' }
                    $status = if ($marked -and $entryAction -in 'Failed', 'ChangeNotConfirmed') { 'Failed' } else { 'Succeeded' }
                    $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{
                                Status           = $status
                                PolicyName       = $policyName
                                EnableSafeList   = [bool]$policy.EnableSafeList
                                IPAllowListEntry = $entry
                                MarkedForRemoval = $marked
                                Action           = $entryAction
                                Detail           = if ($marked) { $detail } else { $null }
                            }))
                }
            }
        }
        catch {
            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; Detail = $_.Exception.Message }))
        }
        finally {
            Disconnect-ExchangeOnline -Confirm:$false -WhatIf:$false -ErrorAction SilentlyContinue
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
