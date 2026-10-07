#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally blocks, automatic forwarding of mail to external recipients in customer tenants.

.DESCRIPTION
    Microsoft now controls external automatic forwarding with the outbound spam filter policy. The default
    AutoForwardingMode of Automatic means Off, so most tenants already block it. For each customer the
    script reports:

    - the default outbound policy's AutoForwardingMode,
    - any other outbound policies with AutoForwardingMode On (deliberate exceptions for some users), and
    - whether the optional mail flow rule exists.

    With -Apply and the default -Method OutboundPolicy, a default policy set to On is changed to Off.
    Custom policies set to On are reported and left alone, because they are usually approved exceptions.
    With -Method TransportRule the script instead creates the mail flow rule the original used, which
    rejects auto-forwarded mail sent from inside the organisation to outside it. The original checked for
    the wrong rule name, so it never found its own rule. This version checks the name it creates.

    -WhatIf shows what -Apply would change. For unattended runs pass -ExchangeAppId with a certificate to
    connect app-only (docs/05). The BlockExternalForwarding Azure Function in the functions folder runs the
    script on a timer.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER Method
    OutboundPolicy (default, Microsoft's recommended control) or TransportRule.

.PARAMETER RuleName
    Name of the mail flow rule for -Method TransportRule.

.PARAMETER RejectMessage
    Rejection text sent back to the user for -Method TransportRule.

.PARAMETER Apply
    Make the change. Without it the script only reports.

.PARAMETER ExchangeAppId
    Application ID of your automation app, to connect app-only instead of as the technician.

.PARAMETER ExchangeCertificate
    The automation app certificate (X509Certificate2) for app-only connections.

.PARAMETER ExchangeCertificateThumbprint
    Thumbprint of the automation app certificate in the local store (Windows) for app-only connections.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./block-external-forwarding.ps1 -AllCustomers -OutputPath ./auto-forwarding.csv

    Reports the external forwarding controls in every customer.

.EXAMPLE
    ./block-external-forwarding.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows whether the default outbound policy in one customer would be changed to block forwarding.

.NOTES
    Replaces the original 2018 method: basic authentication remote PowerShell (New-PSSession, ?DelegatedOrg=
    for DAP customers found with Connect-MsolService and Get-MsolPartnerContract), and an Azure Functions v1
    function with an AES-encrypted stored password and MFA trusted IP exclusions.
    Required GDAP roles: Exchange Administrator (change the outbound policy or create the mail flow rule),
    or Security Administrator for -Method OutboundPolicy only, Global Reader (report only). For app-only
    runs the automation app needs Exchange.ManageAsApp and Exchange Administrator.
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage, Microsoft Graph
    delegated User.Read (organisation name lookup).

.LINK
    https://gcit.com.au/knowledge-base/block-inbox-rules-forwarding-mail-externally-office-365-using-powershell/

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

    [ValidateSet('OutboundPolicy', 'TransportRule')]
    [string]$Method = 'OutboundPolicy',

    [ValidateNotNullOrEmpty()]
    [string]$RuleName = 'Block auto-forwarding to external recipients',

    [ValidateNotNullOrEmpty()]
    [string]$RejectMessage = 'Automatic forwarding to external addresses is turned off to protect your organisation. Contact your IT provider if you need an exception.',

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
        param($Customer, $DefaultMode, [string[]]$OnPolicies, $RuleExists, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId          = $Customer.TenantId
            CustomerName              = $Customer.Name
            Method                    = $Method
            DefaultAutoForwardingMode = $DefaultMode
            PoliciesAllowingForwards  = ($OnPolicies -join '; ')
            TransportRuleExists       = $RuleExists
            Action                    = $Action
            Error                     = $ErrorMessage
        }
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
            try {
                $policies = @(Get-HostedOutboundSpamFilterPolicy)
                $default = $policies | Where-Object { $_.IsDefault -or $_.Name -eq 'Default' } | Select-Object -First 1
                $defaultMode = [string]$default.AutoForwardingMode
                $onPolicies = @($policies | Where-Object { [string]$_.AutoForwardingMode -eq 'On' } | ForEach-Object { $_.Name })
                $ruleExists = [bool](Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue)

                if ($Method -eq 'OutboundPolicy') {
                    if ($defaultMode -ne 'On') {
                        $action = if ($onPolicies.Count -gt 0) { 'DefaultBlocksCustomPolicyAllows' } else { 'AlreadyBlocked' }
                    }
                    elseif (-not $Apply) {
                        $action = 'WouldSetDefaultOff'
                    }
                    elseif ($PSCmdlet.ShouldProcess($customer.Name, "Set AutoForwardingMode Off on outbound policy '$($default.Name)'")) {
                        Set-HostedOutboundSpamFilterPolicy -Identity $default.Name -AutoForwardingMode Off -ErrorAction Stop
                        $defaultMode = [string](Get-HostedOutboundSpamFilterPolicy -Identity $default.Name).AutoForwardingMode
                        $onPolicies = @($onPolicies | Where-Object { $_ -ne $default.Name })
                        $action = if ($defaultMode -eq 'Off') { 'SetDefaultOff' } else { 'SetNotConfirmed' }
                    }
                    else {
                        $action = 'WhatIf'
                    }
                }
                else {
                    if ($ruleExists) {
                        $action = 'RuleExists'
                    }
                    elseif (-not $Apply) {
                        $action = 'WouldCreateRule'
                    }
                    elseif ($PSCmdlet.ShouldProcess($customer.Name, "Create mail flow rule '$RuleName'")) {
                        $ruleParams = @{
                            Name                            = $RuleName
                            Priority                        = 0
                            FromScope                       = 'InOrganization'
                            SentToScope                     = 'NotInOrganization'
                            MessageTypeMatches              = 'AutoForward'
                            RejectMessageEnhancedStatusCode = '5.7.1'
                            RejectMessageReasonText         = $RejectMessage
                            ErrorAction                     = 'Stop'
                        }
                        New-TransportRule @ruleParams | Out-Null
                        $ruleExists = [bool](Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue)
                        $action = if ($ruleExists) { 'CreatedRule' } else { 'RuleNotConfirmed' }
                    }
                    else {
                        $action = 'WhatIf'
                    }
                }
                $row = ConvertTo-ResultRow -Customer $customer -DefaultMode $defaultMode -OnPolicies $onPolicies -RuleExists $ruleExists -Action $action
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            $results.Add($row)
            $row
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
