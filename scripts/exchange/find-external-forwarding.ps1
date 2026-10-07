#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds mailboxes that forward mail to addresses outside the customer's accepted domains.

.DESCRIPTION
    Connects to Exchange Online in each customer through GDAP with Connect-MspExchangeOnline,
    then checks every mailbox for external forwarding. It reads only and changes nothing.

    Three kinds of forwarding are checked:
    - ForwardingSmtpAddress, the forward a user or admin sets on the mailbox.
    - ForwardingAddress, a forward to another recipient. A mail contact or mail user with an
      external address counts as external.
    - Inbox rules that forward, forward as attachment or redirect to an external address
      (only with -IncludeInboxRules, because it reads the rules of every mailbox).

    Each finding also shows the AutoForwardingMode of the customer's default outbound spam
    policy. Microsoft blocks automatic external forwarding by default (Automatic means off),
    so a forward found here may not actually deliver. It is still worth removing, because
    attackers set it up after they take over a mailbox.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER IncludeInboxRules
    Also check the inbox rules of every mailbox. Slower, one call per mailbox.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./find-external-forwarding.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeInboxRules

    Lists every external forward in one customer, including inbox rules.

.EXAMPLE
    ./find-external-forwarding.ps1 -AllCustomers -OutputPath 'C:\Reports\ExternalForwarding.csv'

    Checks mailbox forwarding in every GDAP customer and saves the findings to a CSV file.

.NOTES
    Replaces the original 2018 method: a single-tenant script that opened a Basic authentication
    remote PowerShell session (New-PSSession to outlook.office365.com/powershell-liveid) with
    Get-Credential and only checked ForwardingSmtpAddress.
    Required GDAP roles: Exchange Administrator, or Global Reader for the mailbox checks (reading
    inbox rules with -IncludeInboxRules may need Exchange Administrator).
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Configuring and controlling external email forwarding in Microsoft 365
    https://learn.microsoft.com/en-us/defender-office-365/outbound-spam-policies-external-email-forwarding

.LINK
    https://gcit.com.au/knowledge-base/find-external-forwarding-mailboxes-office-365-powershell/

.LINK
    ../../docs/05-exchange-access.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [switch]$IncludeInboxRules,

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
        foreach ($column in 'Mailbox', 'DisplayName', 'ForwardingType', 'RuleName', 'ExternalRecipient', 'DeliverToMailboxAndForward', 'OutboundAutoForwardingMode', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    function Get-SmtpDomain {
        param([string]$Address)
        if (-not $Address) { return $null }
        $clean = $Address -replace '^(?i)smtp:', ''
        if ($clean -notmatch '@') { return $null }
        ($clean -split '@')[-1].Trim().ToLowerInvariant()
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
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop

            $acceptedDomains = @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
            $autoForwardingMode = $null
            try {
                $autoForwardingMode = (Get-HostedOutboundSpamFilterPolicy -Identity 'Default' -ErrorAction Stop).AutoForwardingMode
            }
            catch {
                Write-Verbose "Could not read the default outbound spam policy. $($_.Exception.Message)"
            }

            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward -ErrorAction Stop)
            $found = 0
            foreach ($mailbox in $mailboxes) {
                $base = @{
                    Mailbox                    = [string]$mailbox.PrimarySmtpAddress
                    DisplayName                = $mailbox.DisplayName
                    DeliverToMailboxAndForward = $mailbox.DeliverToMailboxAndForward
                    OutboundAutoForwardingMode = $autoForwardingMode
                    Status                     = 'Found'
                }

                if ($mailbox.ForwardingSmtpAddress) {
                    $address = ([string]$mailbox.ForwardingSmtpAddress) -replace '^(?i)smtp:', ''
                    $domain = Get-SmtpDomain -Address $address
                    if ($domain -and $acceptedDomains -notcontains $domain) {
                        $data = $base.Clone()
                        $data['ForwardingType'] = 'ForwardingSmtpAddress'
                        $data['ExternalRecipient'] = $address
                        $row = ConvertTo-ResultRow -Customer $customer -Data $data
                        $results.Add($row)
                        $row
                        $found++
                    }
                }

                if ($mailbox.ForwardingAddress) {
                    try {
                        $target = Get-Recipient -Identity ([string]$mailbox.ForwardingAddress) -ErrorAction Stop
                        $address = if ($target.ExternalEmailAddress) { [string]$target.ExternalEmailAddress } else { [string]$target.PrimarySmtpAddress }
                        $address = $address -replace '^(?i)smtp:', ''
                        $domain = Get-SmtpDomain -Address $address
                        if ($domain -and $acceptedDomains -notcontains $domain) {
                            $data = $base.Clone()
                            $data['ForwardingType'] = 'ForwardingAddress'
                            $data['ExternalRecipient'] = $address
                            $row = ConvertTo-ResultRow -Customer $customer -Data $data
                            $results.Add($row)
                            $row
                            $found++
                        }
                    }
                    catch {
                        Write-Warning "$($customer.Name): could not resolve the ForwardingAddress of $($mailbox.PrimarySmtpAddress). $($_.Exception.Message)"
                    }
                }

                if ($IncludeInboxRules) {
                    try {
                        $rules = @(Get-InboxRule -Mailbox ([string]$mailbox.PrimarySmtpAddress) -ErrorAction Stop)
                    }
                    catch {
                        Write-Warning "$($customer.Name): could not read inbox rules for $($mailbox.PrimarySmtpAddress). $($_.Exception.Message)"
                        $rules = @()
                    }
                    foreach ($rule in $rules) {
                        $targets = @($rule.ForwardTo) + @($rule.ForwardAsAttachmentTo) + @($rule.RedirectTo) | Where-Object { $_ }
                        foreach ($recipient in $targets) {
                            # Rule recipients look like "Name" [SMTP:user@fabrikam.com]. Internal ones use [EX:/o=...].
                            if ([string]$recipient -notmatch '(?i)SMTP:(?<address>[^\]\s]+)') { continue }
                            $address = $Matches['address']
                            $domain = Get-SmtpDomain -Address $address
                            if ($domain -and $acceptedDomains -notcontains $domain) {
                                $data = $base.Clone()
                                $data['ForwardingType'] = 'InboxRule'
                                $data['RuleName'] = $rule.Name
                                $data['ExternalRecipient'] = $address
                                $data['DeliverToMailboxAndForward'] = $null
                                $row = ConvertTo-ResultRow -Customer $customer -Data $data
                                $results.Add($row)
                                $row
                                $found++
                            }
                        }
                    }
                }
            }

            if ($found -eq 0) {
                $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'NoneFound'; OutboundAutoForwardingMode = $autoForwardingMode }
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'Failed'; Error = $_.Exception.Message }
            $results.Add($row)
            $row
        }
        finally {
            Close-CustomerSession -Connection $connection
        }
    }

    if ($OutputPath) {
        $folder = Split-Path -Path $OutputPath -Parent
        if ($folder -and -not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false }
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false
        Write-Verbose "Saved $($results.Count) row(s) to $OutputPath."
    }
}
