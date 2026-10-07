#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports native impersonation controls and keeps a mail flow rule that warns users when an external
    sender uses the display name of someone in the organisation.

.DESCRIPTION
    For each customer the script connects to Exchange Online through MspGdap and reports:

    - whether the External tag in Outlook is on (Get-ExternalInOutlook),
    - the default anti-phishing policy's first contact safety tip and impersonation settings
      (Get-AntiPhishPolicy, impersonation settings need Microsoft Defender for Office 365), and
    - whether the display name warning rule exists and is up to date.

    The native controls are Microsoft's recommended defence. The rule is an extra layer for tenants
    without Defender for Office 365. It matches the From header of external mail against the display
    names of the organisation's user mailboxes and prepends an HTML warning. Regular expression special
    characters in each display name (such as brackets, full stops and plus signs) are escaped with a
    backslash, so names like "Jane Citizen (Sales)" match literally. The original passed them unescaped.
    Spaces are left as they are. A mail flow rule is limited to 8 KB in total, including the warning HTML,
    so the script does not create a rule whose patterns are larger than -MaxPatternLength characters and
    reports TooManyNames instead.

    By default the script only reports. With -Apply it creates or updates the rule. -WhatIf shows the change.
    For unattended runs pass -ExchangeAppId with -ExchangeCertificate or -ExchangeCertificateThumbprint to
    connect app-only as your automation app (docs/05). The DisplayNameSpoofRule Azure Function in the
    functions folder runs the script on a timer.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER RuleName
    Name of the mail flow rule.

.PARAMETER WarningHtml
    HTML prepended to matching messages.

.PARAMETER IncludeSharedMailboxes
    Also match the display names of shared mailboxes. Generic names such as "Accounts" cause false positives.

.PARAMETER MaxPatternLength
    Largest total length, in characters, of all display name patterns in the rule. Default 6000, which
    leaves room for the warning HTML within the 8 KB rule limit.

.PARAMETER Apply
    Create or update the rule. Without it the script only reports.

.PARAMETER ExchangeAppId
    Application ID of your automation app, to connect app-only instead of as the technician.

.PARAMETER ExchangeCertificate
    The automation app certificate (X509Certificate2) for app-only connections.

.PARAMETER ExchangeCertificateThumbprint
    Thumbprint of the automation app certificate in the local store (Windows) for app-only connections.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./set-display-name-spoof-rule.ps1 -AllCustomers -OutputPath ./display-name-protection.csv

    Reports native impersonation controls and the rule state for every customer.

.EXAMPLE
    ./set-display-name-spoof-rule.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows whether the rule would be created or updated in one customer.

.NOTES
    Replaces the original 2017 method: basic authentication remote PowerShell (New-PSSession to
    outlook.office365.com/powershell-liveid, with ?DelegatedOrg= for DAP customers found with
    Get-MsolPartnerContract) and an Azure Functions v1 function using an AES-encrypted stored password and
    MFA trusted IP exclusions.
    Required GDAP roles: Exchange Administrator (create or update the rule), Global Reader (report only).
    For app-only runs the automation app needs Exchange.ManageAsApp and Exchange Administrator in the customer.
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage, Microsoft Graph
    delegated User.Read (organisation name lookup).

.LINK
    https://gcit.com.au/knowledge-base/warn-users-external-email-arrives-display-name-someone-organisation/

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

    [ValidateNotNullOrEmpty()]
    [string]$RuleName = 'External senders with matching display names',

    [ValidateNotNullOrEmpty()]
    [string]$WarningHtml = '<div style="border-left:4px solid #910A19;background:#FDF2F4;padding:8px 12px;font-family:Segoe UI,sans-serif;font-size:9pt;color:#212121">This message came from outside the organisation, from a sender whose display name matches someone who works here. Do not click links, open attachments or act on requests unless you have confirmed the sender is genuine.</div><br>',

    [switch]$IncludeSharedMailboxes,

    [ValidateRange(100, 8000)]
    [int]$MaxPatternLength = 6000,

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
        param($Customer, $ExternalTag, $AntiPhish, $NameCount, $PatternLength, $RuleState, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId              = $Customer.TenantId
            CustomerName                  = $Customer.Name
            ExternalTagEnabled            = $ExternalTag
            FirstContactSafetyTips        = $AntiPhish.EnableFirstContactSafetyTips
            MailboxIntelligenceProtection = $AntiPhish.EnableMailboxIntelligenceProtection
            TargetedUserProtection        = $AntiPhish.EnableTargetedUserProtection
            DisplayNameCount              = $NameCount
            PatternLength                 = $PatternLength
            RuleName                      = $RuleName
            RuleState                     = $RuleState
            Action                        = $Action
            Error                         = $ErrorMessage
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
    $mailboxTypes = if ($IncludeSharedMailboxes) { @('UserMailbox', 'SharedMailbox') } else { @('UserMailbox') }

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
                $externalTag = $null
                try { $externalTag = [bool](Get-ExternalInOutlook -ErrorAction Stop | Select-Object -First 1).Enabled }
                catch { Write-Verbose "Get-ExternalInOutlook failed for $($customer.Name): $($_.Exception.Message)" }
                $antiPhish = $null
                try { $antiPhish = Get-AntiPhishPolicy -ErrorAction Stop | Where-Object { $_.IsDefault } | Select-Object -First 1 }
                catch { Write-Verbose "Get-AntiPhishPolicy failed for $($customer.Name): $($_.Exception.Message)" }

                $names = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $mailboxTypes |
                        ForEach-Object { ([string]$_.DisplayName).Trim() } |
                        Where-Object { $_.Length -ge 3 } |
                        Sort-Object -Unique)
                # Escape only regular expression metacharacters. [regex]::Escape would also escape spaces.
                $patterns = @($names | ForEach-Object { $_ -replace '([\\\.\^\$\|\?\*\+\(\)\[\]\{\}])', '\$1' })
                $patternLength = ($patterns | Measure-Object -Property Length -Sum).Sum
                if ($null -eq $patternLength) { $patternLength = 0 }

                $rule = Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue
                $current = (@($rule.HeaderMatchesPatterns | ForEach-Object { [string]$_ } | Sort-Object)) -join "`n"
                $desired = (@($patterns | Sort-Object)) -join "`n"
                $upToDate = [bool]$rule -and $current -ceq $desired
                $ruleState = if (-not $rule) { 'Missing' } elseif ($upToDate) { 'UpToDate' } else { 'OutOfDate' }

                $ruleParams = @{
                    FromScope                         = 'NotInOrganization'
                    HeaderMatchesMessageHeader        = 'From'
                    HeaderMatchesPatterns             = $patterns
                    ApplyHtmlDisclaimerLocation       = 'Prepend'
                    ApplyHtmlDisclaimerText           = $WarningHtml
                    ApplyHtmlDisclaimerFallbackAction = 'Wrap'
                    ErrorAction                       = 'Stop'
                }

                if ($patterns.Count -eq 0) {
                    $action = 'NoMailboxes'
                }
                elseif ($patternLength -gt $MaxPatternLength) {
                    $action = 'TooManyNames'
                }
                elseif ($upToDate) {
                    $action = 'None'
                }
                elseif (-not $Apply) {
                    $action = if ($rule) { 'WouldUpdate' } else { 'WouldCreate' }
                }
                elseif ($PSCmdlet.ShouldProcess($customer.Name, "$(if ($rule) { 'Update' } else { 'Create' }) mail flow rule '$RuleName' with $($patterns.Count) display names")) {
                    if ($rule) {
                        Set-TransportRule -Identity $RuleName @ruleParams
                        $action = 'Updated'
                    }
                    else {
                        New-TransportRule -Name $RuleName -Priority 0 -Comments 'Warns about external senders using internal display names.' @ruleParams | Out-Null
                        $action = 'Created'
                    }
                    $ruleState = if (Get-TransportRule -Identity $RuleName -ErrorAction SilentlyContinue) { 'UpToDate' } else { 'NotConfirmed' }
                }
                else {
                    $action = 'WhatIf'
                }
                $row = ConvertTo-ResultRow -Customer $customer -ExternalTag $externalTag -AntiPhish $antiPhish -NameCount $names.Count -PatternLength $patternLength -RuleState $ruleState -Action $action
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
