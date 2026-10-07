#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally fixes, the WINMAIL.DAT problem by turning off TNEF on the Default remote domain.

.DESCRIPTION
    When Outlook sends a message in Rich Text Format, Exchange can deliver it as TNEF, and
    recipients whose mail client does not understand TNEF see a blank message with a WINMAIL.DAT
    attachment. Setting TNEFEnabled to $false on the Default remote domain makes Exchange Online
    convert those messages to HTML for every external domain that has no remote domain of its own.

    The script connects to Exchange Online in each customer through GDAP with
    Connect-MspExchangeOnline and reads the TNEFEnabled setting of the Default remote domain. It is
    report-only unless you add -Apply. With -Apply it runs
    Set-RemoteDomain -Identity Default -TNEFEnabled $false where the value is not already $false.
    Use -WhatIf with -Apply to preview.

    TNEFEnabled has three values: $true (always TNEF), $false (never TNEF) and empty (follow the
    sender's Outlook setting, the Exchange Online default).

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER Apply
    Make the change. Without it the script only reports the current setting.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./disable-tnef-remote-domain.ps1 -AllCustomers

    Reports the Default remote domain TNEF setting in every GDAP customer.

.EXAMPLE
    ./disable-tnef-remote-domain.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows the Set-RemoteDomain change without making it. Remove -WhatIf to apply it.

.NOTES
    Replaces the original 2016 method: Set-RemoteDomain Default -TNEFEnabled $false, run after
    connecting with the retired Basic authentication remote PowerShell guide
    (New-PSSession to outlook.office365.com/powershell-liveid).
    Required GDAP roles: Exchange Administrator (Global Reader is enough for report-only runs).
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Message format and transmission in Exchange Online
    https://learn.microsoft.com/en-us/exchange/mail-flow-best-practices/message-format-and-transmission

.LINK
    https://gcit.com.au/knowledge-base/how-to-fix-the-winmail-dat-attachment-issue/

.LINK
    ../../docs/05-exchange-access.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    function Get-TargetCustomer {
        param([string[]]$Tenant, [switch]$All)
        if ($All) {
            foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus -ErrorAction Stop)) {
                $skip = $null
                if ($customer.GdapStatus -ne 'active') { $skip = "No active GDAP relationship (status: $($customer.GdapStatus))." }
                [pscustomobject]@{ TenantId = $customer.TenantId; Name = $customer.DisplayName; Skip = $skip }
            }
            return
        }
        foreach ($entry in $Tenant) {
            $id = $entry
            $name = $entry
            try {
                $customer = @(Get-MspCustomer -TenantId $entry -IncludeGdapStatus -ErrorAction Stop) | Select-Object -First 1
                if ($customer) {
                    $id = $customer.TenantId
                    $name = $customer.DisplayName
                }
            }
            catch {
                Write-Verbose "Could not look up the customer name for $entry. $($_.Exception.Message)"
            }
            [pscustomobject]@{ TenantId = $id; Name = $name; Skip = $null }
        }
    }

    function Close-CustomerSession {
        param([object]$Connection)
        if (-not (Get-Command -Name 'Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue)) { return }
        $disconnect = @{ Confirm = $false; WhatIf = $false; ErrorAction = 'SilentlyContinue' }
        if ($Connection -and $Connection.ConnectionId) { $disconnect['ConnectionId'] = $Connection.ConnectionId }
        Disconnect-ExchangeOnline @disconnect
    }

    function ConvertTo-ResultRow {
        param([object]$Customer, [hashtable]$Data)
        $row = [ordered]@{ CustomerTenantId = $Customer.TenantId; CustomerName = $Customer.Name }
        foreach ($column in 'RemoteDomain', 'TNEFEnabledBefore', 'TNEFEnabledAfter', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    function Format-TnefValue {
        param([object]$Value)
        if ($null -eq $Value) { return 'NotSet' }
        [string]$Value
    }

    $tenantInput = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
}

process {
    foreach ($entry in $TenantId) {
        if ($entry) { $tenantInput.Add($entry) }
    }
}

end {
    $customers = @(Get-TargetCustomer -Tenant $tenantInput -All:$AllCustomers)
    foreach ($customer in $customers) {
        if ($customer.Skip) {
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'Skipped'; Error = $customer.Skip }
            $results.Add($row)
            $row
            continue
        }

        $connection = $null
        $data = @{ RemoteDomain = 'Default' }
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop

            $remoteDomain = Get-RemoteDomain -Identity 'Default' -ErrorAction Stop
            $before = $remoteDomain.TNEFEnabled
            $data['TNEFEnabledBefore'] = Format-TnefValue -Value $before
            $data['TNEFEnabledAfter'] = $data['TNEFEnabledBefore']

            if ($before -eq $false) {
                $data['Status'] = 'AlreadySet'
            }
            elseif (-not $Apply) {
                $data['Status'] = 'WouldChange'
            }
            elseif ($PSCmdlet.ShouldProcess("$($customer.Name): remote domain Default", 'Set TNEFEnabled to $false')) {
                Set-RemoteDomain -Identity 'Default' -TNEFEnabled $false -Confirm:$false -ErrorAction Stop
                $after = (Get-RemoteDomain -Identity 'Default' -ErrorAction Stop).TNEFEnabled
                $data['TNEFEnabledAfter'] = Format-TnefValue -Value $after
                $data['Status'] = if ($after -eq $false) { 'Changed' } else { 'NotConfirmed' }
            }
            else {
                $data['Status'] = 'WhatIf'
            }
        }
        catch {
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $data['Status'] = 'Failed'
            $data['Error'] = $_.Exception.Message
        }
        finally {
            Close-CustomerSession -Connection $connection
        }
        $row = ConvertTo-ResultRow -Customer $customer -Data $data
        $results.Add($row)
        $row
    }

    if ($OutputPath) {
        $folder = Split-Path -Path $OutputPath -Parent
        if ($folder -and -not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false }
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false
        Write-Verbose "Saved $($results.Count) row(s) to $OutputPath."
    }
}
