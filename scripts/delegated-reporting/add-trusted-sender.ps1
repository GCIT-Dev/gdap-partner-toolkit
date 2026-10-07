#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally adds, a trusted sender or domain on every mailbox's junk email settings in GDAP customers.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline) and reads each user and shared mailbox's junk email configuration
    (Get-MailboxJunkEmailConfiguration). It reports the mailboxes whose Trusted Senders and Domains
    list doesn't contain every value in -TrustedSender.

    Without -Apply the script only reports. With -Apply it adds the missing values with
    Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains @{ Add = ... } and reads them back.

    A mailbox trusted sender only affects the Outlook junk email filter. It doesn't bypass
    Exchange Online Protection. The better long-term fix for ticket notifications landing in junk
    is to authenticate the sending domain with SPF, DKIM and DMARC.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER TrustedSender
    Email addresses or domains to add, for example 'support@contoso.com' or 'contoso.com'.
.PARAMETER Apply
    Add missing entries. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./add-trusted-sender.ps1 -AllCustomers -TrustedSender 'support@contoso.com' -OutputPath ./trusted-senders.csv

    Lists every mailbox, across all GDAP customers, that doesn't trust the support address yet.
.EXAMPLE
    ./add-trusted-sender.ps1 -TenantId 'fabrikam.onmicrosoft.com' -TrustedSender 'support@contoso.com' -Apply -WhatIf

    Shows which mailboxes in one customer would get the entry, without changing anything.
.NOTES
    Replaces the original 2017 method: MSOnline and DAP (Connect-MsolService, Get-MsolPartnerContract,
    Get-MsolDomain) with basic authentication remote PowerShell (Invoke-Command to
    ps.outlook.com/powershell-liveid?DelegatedOrg=), and advice to set the execution policy to
    Unrestricted. Neither is needed now.
    Required GDAP roles: Exchange Administrator or Exchange Recipient Administrator (Global Reader for
    report-only runs).
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/add-domain-name-trusted-senders-delegated-office-365-users/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxjunkemailconfiguration
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

    [Parameter(Mandatory)]
    [ValidatePattern('^([^@\s]+@)?[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$')]
    [string[]]$TrustedSender,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'Mailbox', 'MissingEntries', 'Action', 'Detail')

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

    function Get-MissingEntry {
        param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string[]]$Wanted)
        $config = Get-MailboxJunkEmailConfiguration -Identity $Identity
        $present = @($config.TrustedSendersAndDomains | ForEach-Object { ([string]$_).ToLowerInvariant() })
        @($Wanted | Where-Object { $present -notcontains $_.ToLowerInvariant() })
    }
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        $customerRows = [System.Collections.Generic.List[object]]::new()
        try {
            Write-Verbose "Checking trusted senders for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
            foreach ($mailbox in $mailboxes) {
                $address = [string]$mailbox.PrimarySmtpAddress
                $values = @{ Mailbox = $address; Status = 'Succeeded' }
                try {
                    $missing = @(Get-MissingEntry -Identity $address -Wanted $TrustedSender)
                    if ($missing.Count -eq 0) { continue }
                    $values.MissingEntries = $missing -join ', '
                    $values.Action = 'ReportOnly'
                    if ($Apply) {
                        if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) mailbox $address", "Add trusted senders $($values.MissingEntries)")) {
                            Set-MailboxJunkEmailConfiguration -Identity $address -TrustedSendersAndDomains @{ Add = $missing }
                            $stillMissing = @(Get-MissingEntry -Identity $address -Wanted $TrustedSender)
                            $values.Action = if ($stillMissing.Count -eq 0) { 'Added' } else { 'ChangeNotConfirmed' }
                            if ($stillMissing.Count -gt 0) { $values.Status = 'Failed' }
                        }
                        else {
                            $values.Action = 'WhatIf'
                        }
                    }
                }
                catch {
                    $values.Status = 'Failed'
                    $values.Detail = $_.Exception.Message
                }
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values $values))
            }

            if ($customerRows.Count -eq 0) {
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; Action = 'None'; Detail = "All $($mailboxes.Count) mailboxes already trust every entry." }))
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
