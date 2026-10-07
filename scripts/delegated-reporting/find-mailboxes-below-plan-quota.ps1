#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds user mailboxes whose quota is lower than their mailbox plan allows, and optionally raises it.
.DESCRIPTION
    The original article raised E3 and E5 mailboxes from 50 GB to 100 GB. Microsoft has given
    those plans a 100 GB mailbox for years, so most mailboxes already match. The ones that don't
    usually had a quota set by hand, or were migrated with a lower quota.

    Instead of guessing from licence names, this script compares each user mailbox's quotas with
    the quotas of the mailbox plan that Exchange Online assigned to it (Get-MailboxPlan). That
    works for every plan, including Business plans, and follows any future Microsoft change.

    For each customer it connects to Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline) and reports user mailboxes whose ProhibitSendReceiveQuota is lower
    than their plan's. Without -Apply it only reports. With -Apply it sets ProhibitSendReceiveQuota,
    ProhibitSendQuota and IssueWarningQuota to the plan's values with Set-Mailbox and reads them back.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER Apply
    Raise the quotas of mailboxes below their plan. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./find-mailboxes-below-plan-quota.ps1 -AllCustomers -OutputPath ./mailbox-quotas.csv

    Lists every user mailbox, in every GDAP customer, whose quota is lower than its plan allows.
.EXAMPLE
    ./find-mailboxes-below-plan-quota.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows which mailboxes would be raised to their plan quota in one customer, without changing anything.
.NOTES
    Replaces the original 2018 method: MSOnline (Connect-MsolService, Get-MsolUser licence matching,
    Get-MsolPartnerContract, Get-MsolDomain) with basic authentication remote PowerShell
    (New-PSSession to outlook.office365.com and ?DelegatedOrg=), and advice to allow-list IP
    addresses to skip MFA.
    Required GDAP roles: Exchange Administrator (or Global Reader for report-only runs).
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/increase-office-365-e3-mailboxes-100-gb-via-powershell/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits
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

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'Mailbox', 'MailboxPlan', 'CurrentQuotaGB', 'PlanQuotaGB', 'Action', 'Detail')

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

    function ConvertTo-ByteCount {
        # Exchange returns quotas as text such as '99 GB (106,300,440,576 bytes)' or 'Unlimited'.
        param([object]$Size)
        if ($null -eq $Size) { return $null }
        $text = [string]$Size
        if ($text -eq 'Unlimited') { return [int64]::MaxValue }
        $match = [regex]::Match($text, '\(([\d,]+) bytes\)')
        if ($match.Success) { return [int64]($match.Groups[1].Value -replace ',', '') }
        $null
    }

    function ConvertTo-QuotaValue {
        # Set-Mailbox treats unqualified quota values as bytes.
        param([object]$Size)
        $bytes = ConvertTo-ByteCount -Size $Size
        if ($null -eq $bytes) { throw "Could not read the plan quota '$Size'." }
        if ($bytes -eq [int64]::MaxValue) { return 'Unlimited' }
        [string]$bytes
    }

    function ConvertTo-Gigabyte {
        param([Nullable[int64]]$Bytes)
        if ($null -eq $Bytes) { return $null }
        if ($Bytes -eq [int64]::MaxValue) { return 'Unlimited' }
        [math]::Round($Bytes / 1GB, 2)
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
            Write-Verbose "Checking mailbox quotas for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $plans = @{}
            foreach ($plan in @(Get-MailboxPlan)) {
                foreach ($key in (@([string]$plan.Name, [string]$plan.Identity, [string]$plan.Alias) | Where-Object { $_ })) {
                    $plans[$key] = $plan
                }
            }

            $properties = 'MailboxPlan', 'ProhibitSendReceiveQuota', 'ProhibitSendQuota', 'IssueWarningQuota'
            $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties $properties)
            foreach ($mailbox in $mailboxes) {
                $address = [string]$mailbox.PrimarySmtpAddress
                $plan = $plans[[string]$mailbox.MailboxPlan]
                if (-not $plan) { continue }

                $current = ConvertTo-ByteCount -Size $mailbox.ProhibitSendReceiveQuota
                $target = ConvertTo-ByteCount -Size $plan.ProhibitSendReceiveQuota
                if ($null -eq $current -or $null -eq $target -or $current -ge $target) { continue }

                $values = @{
                    Status         = 'Succeeded'
                    Mailbox        = $address
                    MailboxPlan    = [string]$mailbox.MailboxPlan
                    CurrentQuotaGB = ConvertTo-Gigabyte -Bytes $current
                    PlanQuotaGB    = ConvertTo-Gigabyte -Bytes $target
                    Action         = 'ReportOnly'
                }
                if ($Apply) {
                    if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) mailbox $address", "Raise quota to $($values.PlanQuotaGB) GB")) {
                        try {
                            $quotaParams = @{
                                Identity                 = $address
                                ProhibitSendReceiveQuota = ConvertTo-QuotaValue -Size $plan.ProhibitSendReceiveQuota
                                ProhibitSendQuota        = ConvertTo-QuotaValue -Size $plan.ProhibitSendQuota
                                IssueWarningQuota        = ConvertTo-QuotaValue -Size $plan.IssueWarningQuota
                            }
                            Set-Mailbox @quotaParams
                            $check = Get-EXOMailbox -Identity $address -Properties ProhibitSendReceiveQuota
                            $after = ConvertTo-ByteCount -Size $check.ProhibitSendReceiveQuota
                            $values.CurrentQuotaGB = ConvertTo-Gigabyte -Bytes $after
                            if ($after -ge $target) {
                                $values.Action = 'Raised'
                            }
                            else {
                                $values.Action = 'ChangeNotConfirmed'
                                $values.Status = 'Failed'
                            }
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
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; Action = 'None'; Detail = "All $($mailboxes.Count) user mailboxes match or exceed their plan quota." }))
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
