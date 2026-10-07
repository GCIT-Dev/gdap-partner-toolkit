#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally turns off, TNEF on the Default remote domain to stop winmail.dat attachments.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline) and reads the TNEFEnabled setting of the Default remote domain.
    Recipients outside the organisation who don't use Outlook receive a winmail.dat attachment when
    TNEF (rich text) formatting is sent to them. Setting TNEFEnabled to $false on the Default remote
    domain stops that for every external domain.

    Without -Apply the script only reports. With -Apply it runs
    Set-RemoteDomain -Identity Default -TNEFEnabled $false and reads the setting back.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER RemoteDomain
    Remote domain to check. Defaults to Default, which applies to every external domain that has no
    remote domain of its own.
.PARAMETER Apply
    Turn TNEF off where it is not already off. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./disable-winmail-dat-tnef.ps1 -AllCustomers -OutputPath ./tnef.csv

    Reports the TNEF setting of every GDAP customer and saves it to CSV.
.EXAMPLE
    ./disable-winmail-dat-tnef.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows whether TNEF would be turned off for one customer, without changing anything.
.NOTES
    Replaces the original 2017 method: MSOnline and DAP (Connect-MsolService, Get-MsolPartnerContract,
    Get-MsolDomain) with basic authentication remote PowerShell (New-PSSession to
    outlook.office365.com/powershell-liveid?DelegatedOrg=).
    Required GDAP roles: Exchange Administrator (or Global Reader for report-only runs).
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/resolve-winmail-dat-attachment-issue-office-365-customer-tenants/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-remotedomain
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

    [ValidateNotNullOrEmpty()]
    [string]$RemoteDomain = 'Default',

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'RemoteDomain', 'TNEFEnabledBefore', 'TNEFEnabledAfter', 'Action', 'Detail')

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

    function Format-TnefValue {
        param([object]$Value)
        if ($null -eq $Value) { 'NotSet (follows user settings)' } else { [string]$Value }
    }
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        $values = @{ RemoteDomain = $RemoteDomain }
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $values.Status = 'Skipped'
            $values.Detail = "No active GDAP relationship (status: $($customer.GdapStatus))."
            $row = ConvertTo-ResultRow -Customer $customer -Values $values
            $results.Add($row)
            $row
            continue
        }

        try {
            Write-Verbose "Checking TNEF for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $domain = Get-RemoteDomain -Identity $RemoteDomain
            $before = $domain.TNEFEnabled
            $values.TNEFEnabledBefore = Format-TnefValue -Value $before
            $values.TNEFEnabledAfter = $values.TNEFEnabledBefore
            $values.Action = 'None'

            if ($before -ne $false) {
                $values.Action = 'ReportOnly'
                if ($Apply) {
                    if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) remote domain $RemoteDomain", 'Set TNEFEnabled to $false')) {
                        Set-RemoteDomain -Identity $RemoteDomain -TNEFEnabled $false
                        $after = (Get-RemoteDomain -Identity $RemoteDomain).TNEFEnabled
                        $values.TNEFEnabledAfter = Format-TnefValue -Value $after
                        $values.Action = if ($after -eq $false) { 'Disabled' } else { 'ChangeNotConfirmed' }
                    }
                    else {
                        $values.Action = 'WhatIf'
                    }
                }
            }
            $values.Status = if ($values.Action -eq 'ChangeNotConfirmed') { 'Failed' } else { 'Succeeded' }
        }
        catch {
            $values.Status = 'Failed'
            $values.Detail = $_.Exception.Message
        }
        finally {
            Disconnect-ExchangeOnline -Confirm:$false -WhatIf:$false -ErrorAction SilentlyContinue
        }

        $row = ConvertTo-ResultRow -Customer $customer -Values $values
        $results.Add($row)
        $row
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
        Write-Verbose "Saved $($results.Count) rows to $OutputPath"
    }
}
