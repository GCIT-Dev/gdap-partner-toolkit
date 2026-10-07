#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds mailboxes that forward mail to external addresses in GDAP customers.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline) and checks both mailbox forwarding settings:
      - ForwardingSmtpAddress, which users and admins can set to any SMTP address, and
      - ForwardingAddress, which points at a recipient object such as a mail contact. The original
        script did not check this one. The recipient is resolved with Get-Recipient and reported
        when its external address is outside the accepted domains.

    The script also reports the automatic forwarding mode of the default outbound spam filter
    policy. With the default (Automatic) mode Exchange Online blocks external automatic
    forwarding, so a forward found here may not deliver mail outside the organisation.

    The script only reads. It never changes forwarding.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./find-external-forwarding-mailboxes.ps1 -AllCustomers -OutputPath ./external-forwards.csv

    Checks every GDAP customer and saves external forwards to CSV.
.EXAMPLE
    ./find-external-forwarding-mailboxes.ps1 -TenantId 'contoso.onmicrosoft.com' -Verbose

    Checks one customer and shows progress.
.NOTES
    Replaces the original 2018 method: MSOnline and DAP (Connect-MsolService, Get-MsolPartnerContract,
    Get-MsolDomain) with basic authentication remote PowerShell (New-PSSession to
    outlook.office365.com/powershell-liveid?DelegatedOrg=), and advice to allow-list IP addresses
    to skip MFA.
    Required GDAP roles: Global Reader, or Exchange Administrator.
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/find-external-forwarding-mailboxes-office-365-customer-tenants-powershell/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    https://learn.microsoft.com/en-us/defender-office-365/outbound-spam-policies-external-email-forwarding
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

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'Mailbox', 'MailboxDisplayName', 'ForwardingSetting', 'ExternalRecipient', 'DeliverToMailboxAndForward', 'AutoForwardingMode', 'Detail')

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

    function Test-ExternalAddress {
        param([string]$Address, [string[]]$AcceptedDomain)
        if (-not $Address -or $Address -notmatch '@') { return $false }
        $domain = ($Address -split '@')[-1].ToLowerInvariant()
        $AcceptedDomain -notcontains $domain
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
            Write-Verbose "Checking mailbox forwarding for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $domains = @(Get-AcceptedDomain | ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
            $forwardingMode = $null
            try {
                $forwardingMode = [string](Get-HostedOutboundSpamFilterPolicy -Identity 'Default').AutoForwardingMode
            }
            catch {
                Write-Verbose "Could not read the outbound spam policy: $($_.Exception.Message)"
            }

            $properties = 'ForwardingSmtpAddress', 'ForwardingAddress', 'DeliverToMailboxAndForward'
            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties $properties)
            foreach ($mailbox in $mailboxes) {
                $address = [string]$mailbox.PrimarySmtpAddress
                $keepCopy = [bool]$mailbox.DeliverToMailboxAndForward

                if ($mailbox.ForwardingSmtpAddress) {
                    $target = ([string]$mailbox.ForwardingSmtpAddress) -replace '^(?i)smtp:', ''
                    if (Test-ExternalAddress -Address $target -AcceptedDomain $domains) {
                        $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; Mailbox = $address; MailboxDisplayName = [string]$mailbox.DisplayName; ForwardingSetting = 'ForwardingSmtpAddress'; ExternalRecipient = $target; DeliverToMailboxAndForward = $keepCopy; AutoForwardingMode = $forwardingMode }))
                    }
                }

                if ($mailbox.ForwardingAddress) {
                    try {
                        $recipient = Get-Recipient -Identity ([string]$mailbox.ForwardingAddress)
                        $target = ([string]$recipient.ExternalEmailAddress) -replace '^(?i)smtp:', ''
                        if (-not $target) { $target = [string]$recipient.PrimarySmtpAddress }
                        if (Test-ExternalAddress -Address $target -AcceptedDomain $domains) {
                            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; Mailbox = $address; MailboxDisplayName = [string]$mailbox.DisplayName; ForwardingSetting = "ForwardingAddress ($($recipient.RecipientTypeDetails))"; ExternalRecipient = $target; DeliverToMailboxAndForward = $keepCopy; AutoForwardingMode = $forwardingMode }))
                        }
                    }
                    catch {
                        $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; Mailbox = $address; MailboxDisplayName = [string]$mailbox.DisplayName; ForwardingSetting = 'ForwardingAddress'; ExternalRecipient = [string]$mailbox.ForwardingAddress; Detail = $_.Exception.Message }))
                    }
                }
            }

            if ($customerRows.Count -eq 0) {
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; AutoForwardingMode = $forwardingMode; Detail = "No external forwarding found in $($mailboxes.Count) mailboxes." }))
            }
        }
        catch {
            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; Detail = $_.Exception.Message }))
        }
        finally {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
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
