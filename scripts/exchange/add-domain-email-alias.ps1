#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Adds a secondary email address on a new domain to every mailbox whose primary address is on a given domain.

.DESCRIPTION
    Typical use: a business registers a second domain (a shorter name, a common misspelling or a
    rebrand) and wants everyone on the main domain to receive mail at the new one too. For
    jane@contoso.com the script adds jane@contoso.net as a secondary (proxy) address. The primary
    address does not change.

    It connects to Exchange Online in each customer through GDAP with Connect-MspExchangeOnline,
    checks that -AliasDomain is an accepted domain, and finds the mailboxes whose primary SMTP
    address is exactly on one of the -MatchDomain domains (an exact domain match, not a text
    match, so contoso.com does not also match contoso.com.au). The original article's text match
    caught both contoso.com and contoso.com.au on purpose. To keep that, name both domains.

    It is report-only unless you add -Apply. With -Apply it runs Set-Mailbox -EmailAddresses
    @{Add=...} for each mailbox that does not already have the address. An address already used
    by another recipient is reported as a conflict and skipped. Use -WhatIf with -Apply to preview.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped. Only useful when the same domain
    pair applies to several customers, which is rare.

.PARAMETER MatchDomain
    One or more domains of the primary SMTP addresses to match, for example contoso.com or
    'contoso.com', 'contoso.com.au'.

.PARAMETER AliasDomain
    The domain for the new secondary address, for example contoso.net. It must already be an
    accepted domain in the customer tenant.

.PARAMETER UseAlias
    Build the new address from the mailbox Alias instead of the local part of the primary SMTP
    address. The original article used the alias.

.PARAMETER Apply
    Make the change. Without it the script only reports what it would change.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./add-domain-email-alias.ps1 -TenantId 'contoso.onmicrosoft.com' -MatchDomain 'contoso.com' -AliasDomain 'contoso.net'

    Reports which mailboxes would get an @contoso.net address.

.EXAMPLE
    ./add-domain-email-alias.ps1 -TenantId 'contoso.onmicrosoft.com' -MatchDomain 'contoso.com' -AliasDomain 'contoso.net' -Apply -WhatIf

    Shows each Set-Mailbox change without making it. Remove -WhatIf to add the addresses.

.EXAMPLE
    ./add-domain-email-alias.ps1 -TenantId 'contoso.onmicrosoft.com' -MatchDomain 'contoso.com', 'contoso.com.au' -AliasDomain 'contoso.net' -Apply

    Adds an @contoso.net address to everyone whose primary address is on either domain, as the
    original article did.

.NOTES
    Replaces the original 2016 method: a script that opened a Basic authentication remote
    PowerShell session (New-PSSession to outlook.office365.com/powershell-liveid) with
    Get-Credential, matched mailboxes with a regular expression (-match) on the primary address
    and added Alias@newdomain to each one.
    Required GDAP roles: Exchange Recipient Administrator or Exchange Administrator (Global Reader
    is enough for report-only runs).
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Add or remove email addresses for a mailbox in Exchange Online
    https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-user-mailboxes/add-or-remove-email-addresses

.LINK
    https://gcit.com.au/knowledge-base/set-domain-email-alias-users-specific-domain/

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

    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+$')]
    [string[]]$MatchDomain,

    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+$')]
    [string]$AliasDomain,

    [switch]$UseAlias,

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
        foreach ($column in 'Mailbox', 'DisplayName', 'NewAddress', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    $matchSuffixes = @($MatchDomain | ForEach-Object { '@' + $_.ToLowerInvariant() })
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

            $accepted = @(Get-AcceptedDomain -ErrorAction Stop | ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
            if ($accepted -notcontains $AliasDomain.ToLowerInvariant()) {
                throw "$AliasDomain is not an accepted domain in this tenant. Add and verify it first."
            }

            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -Properties EmailAddresses -ErrorAction Stop |
                    Where-Object {
                        $primaryAddress = ([string]$_.PrimarySmtpAddress).ToLowerInvariant()
                        @($matchSuffixes | Where-Object { $primaryAddress.EndsWith($_) }).Count -gt 0
                    })

            foreach ($mailbox in $mailboxes) {
                $primary = [string]$mailbox.PrimarySmtpAddress
                $localPart = if ($UseAlias) { [string]$mailbox.Alias } else { ($primary -split '@')[0] }
                $newAddress = '{0}@{1}' -f $localPart, $AliasDomain.ToLowerInvariant()
                $data = @{ Mailbox = $primary; DisplayName = $mailbox.DisplayName; NewAddress = $newAddress }
                try {
                    $existing = @($mailbox.EmailAddresses | ForEach-Object { ([string]$_ -replace '^(?i)smtp:', '').ToLowerInvariant() })
                    if ($existing -contains $newAddress.ToLowerInvariant()) {
                        $data['Status'] = 'AlreadyPresent'
                    }
                    else {
                        $owner = $null
                        try { $owner = Get-EXORecipient -Identity $newAddress -ErrorAction Stop }
                        catch { $owner = $null }
                        if ($owner) {
                            $data['Status'] = 'Conflict'
                            $data['Error'] = "The address is already used by $($owner.PrimarySmtpAddress)."
                        }
                        elseif (-not $Apply) {
                            $data['Status'] = 'WouldAdd'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $primary", "Add secondary address $newAddress")) {
                            Set-Mailbox -Identity $primary -EmailAddresses @{ Add = "smtp:$newAddress" } -Confirm:$false -ErrorAction Stop
                            $data['Status'] = 'Added'
                        }
                        else {
                            $data['Status'] = 'WhatIf'
                        }
                    }
                }
                catch {
                    $data['Status'] = 'Failed'
                    $data['Error'] = $_.Exception.Message
                }
                $row = ConvertTo-ResultRow -Customer $customer -Data $data
                $results.Add($row)
                $row
            }

            if ($mailboxes.Count -eq 0) {
                $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'NoneFound'; Error = "No mailbox has a primary address on $($MatchDomain -join ', ')." }
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
