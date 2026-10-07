#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally turns off, POP and IMAP for mailbox plans and existing mailboxes in GDAP customers.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline) and finds:
      - mailbox plans (Get-CASMailboxPlan) that still turn POP or IMAP on for new mailboxes, and
      - existing mailboxes (Get-EXOCasMailbox) with POP or IMAP turned on.

    Without -Apply the script only reports. With -Apply it turns POP and IMAP off on each plan
    (Set-CASMailboxPlan) and each mailbox (Set-CASMailbox), then reads each one back.

    Basic authentication for POP and IMAP is already retired in Exchange Online, so any remaining
    POP or IMAP client must use OAuth. Use -ExcludeMailbox for mailboxes that a known application
    still reads over POP or IMAP with OAuth.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER ExcludeMailbox
    Primary SMTP addresses of mailboxes to leave unchanged. They are still reported.
.PARAMETER Apply
    Turn POP and IMAP off on plans and mailboxes. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./disable-pop-imap.ps1 -AllCustomers -OutputPath ./pop-imap.csv

    Lists every plan and mailbox with POP or IMAP still on, for every GDAP customer.
.EXAMPLE
    ./disable-pop-imap.ps1 -TenantId 'contoso.onmicrosoft.com' -ExcludeMailbox 'scanner@contoso.com' -Apply -WhatIf

    Shows what would be turned off in one customer, leaving the scanner mailbox alone.
.NOTES
    Replaces the original 2018 method: MSOnline and DAP (Get-MsolPartnerContract, Get-MsolDomain)
    with basic authentication remote PowerShell (New-PSSession to outlook.office365.com, including
    ?DelegatedOrg=), and advice to allow-list IP addresses to skip MFA.
    Required GDAP roles: Exchange Administrator (or Global Reader for report-only runs).
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/disable-pop-imap-mailboxes-office-365/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/pop3-and-imap4/enable-or-disable-pop3-or-imap4-access
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

    [string[]]$ExcludeMailbox = @(),

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'ObjectType', 'Identity', 'PopEnabled', 'ImapEnabled', 'Action', 'Detail')

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
    $excluded = @($ExcludeMailbox | ForEach-Object { $_.ToLowerInvariant() })

    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        $customerRows = [System.Collections.Generic.List[object]]::new()
        try {
            Write-Verbose "Checking POP and IMAP for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $plans = @(Get-CASMailboxPlan | Where-Object { $_.PopEnabled -or $_.ImapEnabled })
            foreach ($plan in $plans) {
                $values = @{ ObjectType = 'MailboxPlan'; Identity = [string]$plan.Identity; PopEnabled = [bool]$plan.PopEnabled; ImapEnabled = [bool]$plan.ImapEnabled; Action = 'ReportOnly'; Status = 'Succeeded' }
                if ($Apply) {
                    if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) mailbox plan $($plan.Identity)", 'Turn off POP and IMAP')) {
                        try {
                            Set-CASMailboxPlan -Identity $plan.Identity -PopEnabled $false -ImapEnabled $false
                            $check = Get-CASMailboxPlan | Where-Object { [string]$_.Identity -eq [string]$plan.Identity } | Select-Object -First 1
                            $values.PopEnabled = [bool]$check.PopEnabled
                            $values.ImapEnabled = [bool]$check.ImapEnabled
                            $values.Action = if (-not $check.PopEnabled -and -not $check.ImapEnabled) { 'Disabled' } else { 'ChangeNotConfirmed' }
                            if ($values.Action -eq 'ChangeNotConfirmed') { $values.Status = 'Failed' }
                        }
                        catch {
                            $values.Status = 'Failed'
                            $values.Detail = $_.Exception.Message
                        }
                    }
                    else {
                        $values.Action = 'WhatIf'
                    }
                }
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values $values))
            }

            $mailboxes = @(Get-EXOCasMailbox -ResultSize Unlimited -PropertySets Minimum -Filter 'PopEnabled -eq $true -or ImapEnabled -eq $true')
            foreach ($mailbox in $mailboxes) {
                $address = [string]$mailbox.PrimarySmtpAddress
                $values = @{ ObjectType = 'Mailbox'; Identity = $address; PopEnabled = [bool]$mailbox.PopEnabled; ImapEnabled = [bool]$mailbox.ImapEnabled; Action = 'ReportOnly'; Status = 'Succeeded' }
                if ($excluded -contains $address.ToLowerInvariant()) {
                    $values.Action = 'Excluded'
                }
                elseif ($Apply) {
                    if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) mailbox $address", 'Turn off POP and IMAP')) {
                        try {
                            Set-CASMailbox -Identity $address -PopEnabled $false -ImapEnabled $false
                            $check = Get-EXOCasMailbox -Identity $address -PropertySets Minimum
                            $values.PopEnabled = [bool]$check.PopEnabled
                            $values.ImapEnabled = [bool]$check.ImapEnabled
                            $values.Action = if (-not $check.PopEnabled -and -not $check.ImapEnabled) { 'Disabled' } else { 'ChangeNotConfirmed' }
                            if ($values.Action -eq 'ChangeNotConfirmed') { $values.Status = 'Failed' }
                        }
                        catch {
                            $values.Status = 'Failed'
                            $values.Detail = $_.Exception.Message
                        }
                    }
                    else {
                        $values.Action = 'WhatIf'
                    }
                }
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values $values))
            }

            if ($customerRows.Count -eq 0) {
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; ObjectType = 'None'; Action = 'None'; Detail = 'POP and IMAP are already off for every mailbox plan and mailbox.' }))
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
