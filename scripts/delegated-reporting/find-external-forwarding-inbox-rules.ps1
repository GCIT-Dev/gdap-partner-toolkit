#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds inbox rules that forward or redirect mail to external addresses in GDAP customers.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline), reads the accepted domains, then checks the inbox rules of every
    user and shared mailbox (Get-InboxRule). A rule is reported when its ForwardTo,
    ForwardAsAttachmentTo or RedirectTo actions include an SMTP address outside the accepted
    domains.

    The script also reports the automatic forwarding mode of the default outbound spam filter
    policy. Since late 2020 the default (Automatic) blocks external automatic forwarding, so a rule
    found here may not actually deliver mail outside the organisation. It is still worth review,
    because attackers create these rules after taking over a mailbox.

    The script only reads. It never changes or removes a rule.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./find-external-forwarding-inbox-rules.ps1 -AllCustomers -OutputPath ./external-rules.csv

    Checks every GDAP customer and saves the external forwarding rules to CSV.
.EXAMPLE
    'contoso.onmicrosoft.com', 'fabrikam.onmicrosoft.com' | ./find-external-forwarding-inbox-rules.ps1

    Checks two customers passed through the pipeline.
.NOTES
    Replaces the original 2018 method: basic authentication remote PowerShell (New-PSSession to
    outlook.office365.com/powershell-liveid/ and ?DelegatedOrg=), MSOnline and DAP
    (Connect-MsolService, Get-MsolPartnerContract, Get-MsolDomain), and advice to allow-list IP
    addresses to skip MFA.
    Required GDAP roles: Global Reader, or Exchange Administrator.
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/find-inbox-rules-forward-mail-externally-office-365-powershell/
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
    $resultColumns = @('Status', 'Mailbox', 'MailboxDisplayName', 'RuleName', 'RuleIdentity', 'RuleEnabled', 'RuleDescription', 'ForwardType', 'ExternalRecipients', 'AutoForwardingMode', 'Detail')

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

    function Get-ExternalAddress {
        # Rule recipients look like '"Name" [SMTP:someone@example.com]'. Internal recipients
        # often show as [EX:/o=...] and are never external.
        param([object[]]$Recipient, [string[]]$AcceptedDomain)
        foreach ($entry in @($Recipient | Where-Object { $_ })) {
            foreach ($match in [regex]::Matches([string]$entry, '(?i)smtp:([^\]\s]+@([^\]\s]+))')) {
                $domain = $match.Groups[2].Value.ToLowerInvariant()
                if ($AcceptedDomain -notcontains $domain) { $match.Groups[1].Value }
            }
        }
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
            Write-Verbose "Checking inbox rules for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $domains = @(Get-AcceptedDomain | ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
            $forwardingMode = $null
            try {
                $forwardingMode = [string](Get-HostedOutboundSpamFilterPolicy -Identity 'Default').AutoForwardingMode
            }
            catch {
                Write-Verbose "Could not read the outbound spam policy: $($_.Exception.Message)"
            }

            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
            foreach ($mailbox in $mailboxes) {
                $address = [string]$mailbox.PrimarySmtpAddress
                try {
                    $rules = @(Get-InboxRule -Mailbox $address)
                }
                catch {
                    $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; Mailbox = $address; AutoForwardingMode = $forwardingMode; Detail = $_.Exception.Message }))
                    continue
                }
                foreach ($rule in $rules) {
                    $types = [System.Collections.Generic.List[string]]::new()
                    $external = [System.Collections.Generic.List[string]]::new()
                    foreach ($action in 'ForwardTo', 'ForwardAsAttachmentTo', 'RedirectTo') {
                        $found = @(Get-ExternalAddress -Recipient @($rule.$action) -AcceptedDomain $domains)
                        if ($found.Count -gt 0) {
                            $types.Add($action)
                            $found | ForEach-Object { $external.Add($_) }
                        }
                    }
                    if ($external.Count -eq 0) { continue }
                    $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{
                                Status             = 'Succeeded'
                                Mailbox            = $address
                                MailboxDisplayName = [string]$mailbox.DisplayName
                                RuleName           = [string]$rule.Name
                                RuleIdentity       = [string]$rule.Identity
                                RuleEnabled        = [bool]$rule.Enabled
                                RuleDescription    = ([string]$rule.Description).Trim()
                                ForwardType        = $types -join ', '
                                ExternalRecipients = ($external | Sort-Object -Unique) -join ', '
                                AutoForwardingMode = $forwardingMode
                            }))
                }
            }

            if ($customerRows.Count -eq 0) {
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; AutoForwardingMode = $forwardingMode; Detail = "No external forwarding rules found in $($mailboxes.Count) mailboxes." }))
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
