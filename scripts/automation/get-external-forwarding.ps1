#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds mailboxes that forward mail outside the organisation in customer tenants, and optionally removes
    mailbox-level forwarding.

.DESCRIPTION
    For each customer the script connects to Exchange Online through MspGdap and returns one row for each
    external forward it finds:

    - mailbox forwarding (ForwardingSmtpAddress, or ForwardingAddress pointing at an external contact),
    - with -IncludeInboxRules, inbox rules that forward or redirect to an external address.

    An address is external when its domain is not one of the tenant's accepted domains. Each row also shows
    the default outbound spam policy's AutoForwardingMode. Automatic and Off block external automatic
    forwarding, so a forward listed here only delivers mail if the policy (or another outbound policy
    scoped to the user) is On.

    By default the script only reports. With -Apply it clears mailbox forwarding on the mailboxes it found
    (inbox rules are reported for review, not changed). -WhatIf shows what -Apply would change.

    Removal is deliberately not exposed through an HTTP endpoint. The ExternalForwardingReport Azure
    Function in the functions folder runs the report on a timer and queues each new forward once (as the
    original only alerted on forwards it had not seen before), and a technician removes forwards
    interactively after review. -Apply clears only the forwarding setting that points outside the
    organisation.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER IncludeInboxRules
    Also check every user and shared mailbox's inbox rules. Slower in large tenants.

.PARAMETER Apply
    Clear mailbox forwarding on the mailboxes found. Without it the script only reports.

.PARAMETER ExchangeAppId
    Application ID of your automation app, to connect app-only instead of as the technician.

.PARAMETER ExchangeCertificate
    The automation app certificate (X509Certificate2) for app-only connections.

.PARAMETER ExchangeCertificateThumbprint
    Thumbprint of the automation app certificate in the local store (Windows) for app-only connections.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-external-forwarding.ps1 -AllCustomers -IncludeInboxRules -OutputPath ./external-forwards.csv

    Lists every external forward across all customers.

.EXAMPLE
    ./get-external-forwarding.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows which mailbox forwards would be removed in one customer.

.NOTES
    Replaces the original 2018 method: an Azure Functions v1 timer function with MSOnline and DAP and an
    AES-encrypted stored password, basic authentication remote PowerShell (?DelegatedOrg=), Azure Table
    storage signed with the storage account key, and an HTTP function protected only by a function key
    that removed forwards when a Microsoft Flow approval was clicked.
    Required GDAP roles: Exchange Administrator (remove forwarding), Global Reader (report only).
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage, Microsoft Graph
    delegated User.Read (organisation name lookup).

.LINK
    https://gcit.com.au/knowledge-base/monitor-external-mailbox-forwards-office-365-customer-tenants/

.LINK
    docs/05-exchange-access.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [switch]$IncludeInboxRules,

    [switch]$Apply,

    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$ExchangeAppId,

    [System.Security.Cryptography.X509Certificates.X509Certificate2]$ExchangeCertificate,

    [ValidatePattern('^[0-9a-fA-F]{40}$')]
    [string]$ExchangeCertificateThumbprint,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    if ($ExchangeAppId -and -not ($ExchangeCertificate -or $ExchangeCertificateThumbprint)) {
        throw '-ExchangeAppId needs -ExchangeCertificate or -ExchangeCertificateThumbprint.'
    }

    function Get-ExchangeConnectParameter {
        param([string]$Tenant)
        $connect = @{ TenantId = $Tenant }
        if ($ExchangeAppId) {
            $connect.AppOnly = $true
            $connect.AppId = $ExchangeAppId
            if ($ExchangeCertificate) { $connect.Certificate = $ExchangeCertificate } else { $connect.CertificateThumbprint = $ExchangeCertificateThumbprint }
        }
        $connect
    }

    function ConvertTo-ResultRow {
        param($Customer, $Mailbox, $Source, $RuleName, $Recipient, $KeepCopy, $Mode, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId           = $Customer.TenantId
            CustomerName               = $Customer.Name
            Mailbox                    = $Mailbox
            Source                     = $Source
            InboxRuleName              = $RuleName
            ExternalRecipient          = $Recipient
            DeliverToMailboxAndForward = $KeepCopy
            OutboundAutoForwardingMode = $Mode
            Action                     = $Action
            Error                      = $ErrorMessage
        }
    }

    function Get-SmtpAddress {
        param([string]$Value)
        if ($Value -match 'SMTP:([^\]\s]+)') { return $Matches[1] }
        if ($Value -match '^[^@\s"]+@[^@\s"]+$') { return $Value }
        $null
    }

    function Test-ExternalAddress {
        param([string]$Address, [string[]]$Domains)
        if (-not $Address -or $Address -notmatch '@') { return $false }
        $domain = ($Address -split '@')[-1].ToLowerInvariant()
        -not ($Domains -contains $domain)
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $connect = Get-ExchangeConnectParameter -Tenant $customer.TenantId
            $null = Connect-MspExchangeOnline @connect
            $customerRows = [System.Collections.Generic.List[object]]::new()
            try {
                $accepted = @(Get-AcceptedDomain | ForEach-Object { ([string]$_.DomainName).ToLowerInvariant() })
                $mode = (Get-HostedOutboundSpamFilterPolicy -Identity 'Default' -ErrorAction SilentlyContinue).AutoForwardingMode

                $forwarding = @(Get-EXOMailbox -ResultSize Unlimited -Filter 'ForwardingSmtpAddress -ne $null -or ForwardingAddress -ne $null' -Properties ForwardingSmtpAddress, ForwardingAddress, DeliverToMailboxAndForward)
                foreach ($mailbox in $forwarding) {
                    # Check both settings separately, so that removal clears only the one that points outside
                    # the organisation (as the original cleared only ForwardingSmtpAddress).
                    $candidates = @()
                    if ($mailbox.ForwardingSmtpAddress) {
                        $candidates += [pscustomobject]@{ Setting = 'ForwardingSmtpAddress'; Address = ([string]$mailbox.ForwardingSmtpAddress) -replace '^smtp:', '' }
                    }
                    if ($mailbox.ForwardingAddress) {
                        $recipient = Get-EXORecipient -Identity ([string]$mailbox.ForwardingAddress) -ErrorAction SilentlyContinue
                        if ($recipient) { $candidates += [pscustomobject]@{ Setting = 'ForwardingAddress'; Address = [string]$recipient.PrimarySmtpAddress } }
                    }

                    foreach ($candidate in $candidates) {
                        $address = $candidate.Address
                        if (-not (Test-ExternalAddress -Address $address -Domains $accepted)) { continue }

                        if (-not $Apply) {
                            $action = 'ReportOnly'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($mailbox.UserPrincipalName) in $($customer.Name)", "Remove forwarding to $address")) {
                            $clear = @{ Identity = $mailbox.UserPrincipalName; ErrorAction = 'Stop' }
                            $clear[$candidate.Setting] = $null
                            # Turn off "keep a copy and forward" only when no other forwarding setting remains.
                            $other = if ($candidate.Setting -eq 'ForwardingSmtpAddress') { $mailbox.ForwardingAddress } else { $mailbox.ForwardingSmtpAddress }
                            if (-not $other) { $clear.DeliverToMailboxAndForward = $false }
                            Set-Mailbox @clear
                            $action = 'Removed'
                        }
                        else {
                            $action = 'WhatIf'
                        }
                        $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Mailbox $mailbox.UserPrincipalName -Source 'MailboxForwarding' -Recipient $address -KeepCopy $mailbox.DeliverToMailboxAndForward -Mode $mode -Action $action))
                    }
                }

                if ($IncludeInboxRules) {
                    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox)
                    foreach ($mailbox in $mailboxes) {
                        $rules = @(Get-InboxRule -Mailbox $mailbox.UserPrincipalName -ErrorAction SilentlyContinue | Where-Object { $_.Enabled })
                        foreach ($rule in $rules) {
                            $targets = @($rule.ForwardTo) + @($rule.ForwardAsAttachmentTo) + @($rule.RedirectTo) | Where-Object { $_ }
                            $external = @($targets | ForEach-Object { Get-SmtpAddress -Value ([string]$_) } | Where-Object { Test-ExternalAddress -Address $_ -Domains $accepted } | Sort-Object -Unique)
                            if ($external.Count -eq 0) { continue }
                            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Mailbox $mailbox.UserPrincipalName -Source 'InboxRule' -RuleName $rule.Name -Recipient ($external -join '; ') -Mode $mode -Action 'ReviewInboxRule'))
                        }
                    }
                }
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            foreach ($row in $customerRows) {
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
