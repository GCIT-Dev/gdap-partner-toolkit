#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Mocked smoke tests for the Exchange example scripts in scripts/exchange.
# Every MspGdap call that would reach Microsoft (Get-MspCustomer, Connect-MspExchangeOnline,
# Connect-MspSecurityCompliance, Invoke-MspGraphRequest, Get-MspAuthHeader) is mocked, and every
# Exchange Online cmdlet is a stub function replaced by a mock. Nothing here contacts a tenant.

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    Import-Module (Join-Path $repoRoot 'src/MspGdap/MspGdap.psd1') -Force
    $script:ScriptFolder = Join-Path $repoRoot 'scripts/exchange'
    $script:Tenant = 'contoso.onmicrosoft.com'
    $script:TenantGuid = '33333333-3333-3333-3333-333333333333'

    # Stub functions for the Exchange Online cmdlets the scripts call, so mocks can bind to them.
    $script:ExoStubs = [ordered]@{
        'Disconnect-ExchangeOnline'          = '$ConnectionId'
        'Get-AcceptedDomain'                 = ''
        'Get-HostedOutboundSpamFilterPolicy' = '$Identity'
        'Get-EXOMailbox'                     = '$Identity, $ResultSize, [string[]]$Properties, $RecipientTypeDetails'
        'Get-EXOCasMailbox'                  = '$Identity, $ResultSize, [string[]]$Properties'
        'Get-EXORecipient'                   = '$Identity'
        'Get-EXOMailboxFolderPermission'     = '$Identity, $User'
        'Get-EXOMailboxFolderStatistics'     = '$Identity, $Folderscope'
        'Get-EXOMailboxStatistics'           = '$Identity, [switch]$Archive'
        'Get-Recipient'                      = '$Identity, $RecipientPreviewFilter, $ResultSize'
        'Get-InboxRule'                      = '$Mailbox'
        'Get-OrganizationConfig'             = ''
        'Get-RetentionCompliancePolicy'      = ''
        'Set-MailboxFolderPermission'        = '$Identity, $User, $AccessRights'
        'Get-MobileDevice'                   = '$ResultSize'
        'Set-Mailbox'                        = '$Identity, $EmailAddresses, $RetentionPolicy'
        'Get-DynamicDistributionGroup'       = '$Identity'
        'New-DynamicDistributionGroup'       = '$Name, $RecipientFilter, $PrimarySmtpAddress'
        'Get-RemoteDomain'                   = '$Identity'
        'Set-RemoteDomain'                   = '$Identity, $TNEFEnabled'
        'Get-MobileDeviceMailboxPolicy'      = ''
        'New-MobileDeviceMailboxPolicy'      = '$Name, $AllowCamera, $AllowWiFi, $AllowInternetSharing, $AllowBrowser'
        'Set-CASMailbox'                     = '$Identity, $ActiveSyncMailboxPolicy'
        'Get-RetentionPolicyTag'             = '$Identity'
        'New-RetentionPolicyTag'             = '$Name, $Type, $RetentionEnabled, $AgeLimitForRetention, $RetentionAction'
        'Get-RetentionPolicy'                = '$Identity'
        'New-RetentionPolicy'                = '$Name, $RetentionPolicyTagLinks'
        'Set-RetentionPolicy'                = '$Identity, $RetentionPolicyTagLinks'
        'Enable-OrganizationCustomization'   = ''
        'Enable-Mailbox'                     = '$Identity, [switch]$Archive'
        'Start-ManagedFolderAssistant'       = '$Identity'
        'Get-UnifiedGroup'                   = '$Identity, $ResultSize'
        'Get-UnifiedGroupLinks'              = '$Identity, $LinkType, $ResultSize'
        'Add-UnifiedGroupLinks'              = '$Identity, $LinkType, $Links'
        'New-UnifiedGroup'                   = '$DisplayName, $Alias, $AccessType'
        'Get-DistributionGroup'              = '$Identity, $ResultSize'
        'Get-DistributionGroupMember'        = '$Identity, $ResultSize'
        'Get-MailContact'                    = '$Identity'
        'New-MailContact'                    = '$Name, $ExternalEmailAddress'
        'Add-DistributionGroupMember'        = '$Identity, $Member, [switch]$BypassSecurityGroupManagerCheck'
        'Remove-DistributionGroupMember'     = '$Identity, $Member, [switch]$BypassSecurityGroupManagerCheck'
        'Remove-MailContact'                 = '$Identity'
        'New-DistributionGroup'              = '$Name, $Type, $PrimarySmtpAddress'
        'Get-OwaMailboxPolicy'               = ''
        'Set-OwaMailboxPolicy'               = '$Identity, $ConditionalAccessPolicy'
        'Get-ElevatedAccessApprovalPolicy'   = ''
        'Get-ElevatedAccessRequest'          = ''
    }
    foreach ($name in $script:ExoStubs.Keys) {
        $body = [scriptblock]::Create("[CmdletBinding(SupportsShouldProcess)] param($($script:ExoStubs[$name]))")
        Set-Item -Path "function:global:$name" -Value $body
    }

    function script:Invoke-ExchangeScript {
        param([Parameter(Mandatory)][string]$Name, [hashtable]$Parameters = @{})
        $path = Join-Path $script:ScriptFolder "$Name.ps1"
        @(& $path @Parameters 3>$null)
    }

    function script:Assert-RowShape {
        param([object[]]$Rows, [string[]]$Column)
        $Rows.Count | Should -BeGreaterThan 0
        foreach ($row in $Rows) {
            $row | Should -BeOfType [pscustomobject]
            $names = @($row.PSObject.Properties.Name)
            $names[0] | Should -Be 'CustomerTenantId'
            $names[1] | Should -Be 'CustomerName'
            foreach ($item in $Column + 'Status') { $names | Should -Contain $item }
        }
    }
}

AfterAll {
    foreach ($name in $script:ExoStubs.Keys) { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
    Remove-Variable -Name ExoTest -Scope Global -ErrorAction SilentlyContinue
    Remove-Module -Name MspGdap -ErrorAction SilentlyContinue
}

Describe 'Exchange example scripts' {
    BeforeEach {
        $global:ExoTest = @{ Tnef = $null }
        Mock Get-MspCustomer {
            if ($TenantId -and $TenantId -like 'broken*') { throw 'lookup failed' }
            if ($TenantId) {
                return [pscustomobject]@{ TenantId = '33333333-3333-3333-3333-333333333333'; DisplayName = 'Contoso'; GdapStatus = 'active' }
            }
            [pscustomobject]@{ TenantId = '33333333-3333-3333-3333-333333333333'; DisplayName = 'Contoso'; GdapStatus = 'active' }
            [pscustomobject]@{ TenantId = '55555555-5555-5555-5555-555555555555'; DisplayName = 'Fabrikam'; GdapStatus = 'terminated' }
        }
        Mock Connect-MspExchangeOnline {
            [pscustomobject]@{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; ConnectionId = 'conn-1'; Mode = 'Delegated'; UserPrincipalName = 'tech@contoso-msp.onmicrosoft.com'; TokenExpiryTimeUTC = [datetime]::UtcNow.AddHours(1) }
        }
        Mock Connect-MspSecurityCompliance {
            [pscustomobject]@{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; ConnectionId = 'conn-2'; Mode = 'SecurityComplianceDelegated'; UserPrincipalName = 'tech@contoso-msp.onmicrosoft.com'; Experimental = $true }
        }
        Mock Get-MspAuthHeader { @{ Authorization = 'Bearer ' + 'placeholder' } }
        Mock Invoke-MspGraphRequest { @() }
        Mock Disconnect-ExchangeOnline {}
        Mock Get-AcceptedDomain { [pscustomobject]@{ DomainName = 'contoso.com' }; [pscustomobject]@{ DomainName = 'contoso.net' }; [pscustomobject]@{ DomainName = 'contoso.onmicrosoft.com' } }
        Mock Get-OrganizationConfig { [pscustomobject]@{ DisplayName = 'Contoso'; IsDehydrated = $true; ElevatedAccessControl = 'Enabled'; ElevatedAccessApprovers = @('pam-approvers@contoso.com') } }
    }

    Context 'common behaviour' {
        It 'records a failed customer and carries on with the next one' {
            Mock Connect-MspExchangeOnline { throw 'AADSTS50020 simulated' } -ParameterFilter { $TenantId -eq 'broken.onmicrosoft.com' }
            Mock Get-EXOMailbox { @() }
            Mock Get-HostedOutboundSpamFilterPolicy { [pscustomobject]@{ AutoForwardingMode = 'Automatic' } }
            $rows = Invoke-ExchangeScript -Name 'find-external-forwarding' -Parameters @{ TenantId = @('broken.onmicrosoft.com', $script:Tenant) }
            $rows.Count | Should -Be 2
            $rows[0].Status | Should -Be 'Failed'
            $rows[0].Error | Should -BeLike '*AADSTS50020*'
            $rows[1].Status | Should -Be 'NoneFound'
            $rows[1].CustomerName | Should -Be 'Contoso'
        }

        It 'skips -AllCustomers customers without an active GDAP relationship' {
            Mock Get-RemoteDomain { [pscustomobject]@{ TNEFEnabled = $false } }
            $rows = Invoke-ExchangeScript -Name 'disable-tnef-remote-domain' -Parameters @{ AllCustomers = $true }
            $rows.Count | Should -Be 2
            ($rows | Where-Object CustomerName -eq 'Fabrikam').Status | Should -Be 'Skipped'
            ($rows | Where-Object CustomerName -eq 'Contoso').Status | Should -Be 'AlreadySet'
            Should -Invoke Connect-MspExchangeOnline -Times 1 -Exactly
        }

        It 'accepts tenants from the pipeline and writes a CSV' {
            Mock Get-RemoteDomain { [pscustomobject]@{ TNEFEnabled = $null } }
            $csv = Join-Path $TestDrive 'out/tnef.csv'
            $rows = @($script:Tenant | & (Join-Path $script:ScriptFolder 'disable-tnef-remote-domain.ps1') -OutputPath $csv)
            $rows.Count | Should -Be 1
            Test-Path -LiteralPath $csv | Should -BeTrue
            (Import-Csv -LiteralPath $csv)[0].CustomerTenantId | Should -Be $script:TenantGuid
        }

        It 'never calls a retired or password-based connection method in any script' {
            $forbidden = 'Connect-MsolService', 'Connect-AzureAD', 'Connect-AzAccount', 'Login-AzureRmAccount', 'New-PSSession', 'Import-PSSession', 'ConvertTo-SecureString', 'Get-Credential', 'Connect-ExchangeOnline', 'Connect-IPPSSession'
            foreach ($file in Get-ChildItem -Path $script:ScriptFolder -Filter '*.ps1') {
                $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
                $commands = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() }
                foreach ($name in $forbidden) { $commands | Should -Not -Contain $name -Because "$($file.Name) must authenticate through MspGdap" }
            }
        }

        It 'documents every script with synopsis, two examples, notes and links' {
            foreach ($file in Get-ChildItem -Path $script:ScriptFolder -Filter '*.ps1') {
                $help = Get-Help -Name $file.FullName -Full
                $help.Synopsis | Should -Not -BeNullOrEmpty -Because $file.Name
                @($help.examples.example).Count | Should -BeGreaterOrEqual 2 -Because $file.Name
                $notes = ($help.alertSet.alert | ForEach-Object { $_.Text }) -join ' '
                $notes | Should -BeLike '*Replaces the original 20?? method:*' -Because $file.Name
                $notes | Should -BeLike '*Required GDAP roles:*' -Because $file.Name
                $notes | Should -BeLike '*Required partner app permissions:*' -Because $file.Name
                $links = @($help.relatedLinks.navigationLink | ForEach-Object { $_.uri + $_.linkText })
                ($links -join ' ') | Should -Match 'https://\S+' -Because $file.Name
                ($links -join ' ') | Should -BeLike '*docs/*' -Because $file.Name
                $documented = @($help.parameters.parameter | Where-Object { $_.description } | ForEach-Object { $_.name })
                $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
                foreach ($parameter in $ast.ParamBlock.Parameters) {
                    $documented | Should -Contain $parameter.Name.VariablePath.UserPath -Because "$($file.Name) documents every parameter"
                }
            }
        }
    }

    Context 'connect-exchange-online' {
        It 'connects, checks the organisation and disconnects' {
            $rows = Invoke-ExchangeScript -Name 'connect-exchange-online' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'Organization', 'IsDehydrated', 'KeptOpen'
            $rows[0].Status | Should -Be 'Connected'
            $rows[0].CustomerTenantId | Should -Be $script:TenantGuid
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly -ParameterFilter { $ConnectionId -eq 'conn-1' }
        }
        It 'keeps one session open with -KeepConnected' {
            $rows = Invoke-ExchangeScript -Name 'connect-exchange-online' -Parameters @{ TenantId = $script:Tenant; KeepConnected = $true }
            $rows[0].KeptOpen | Should -BeTrue
            Should -Invoke Disconnect-ExchangeOnline -Times 0 -Exactly
        }
        It 'refuses -KeepConnected with more than one customer' {
            { Invoke-ExchangeScript -Name 'connect-exchange-online' -Parameters @{ TenantId = @($script:Tenant, 'fabrikam.onmicrosoft.com'); KeepConnected = $true } } | Should -Throw '*exactly one customer*'
        }
    }

    Context 'connect-security-compliance' {
        It 'connects through Connect-MspSecurityCompliance and reports the policy count' {
            Mock Get-RetentionCompliancePolicy { [pscustomobject]@{ Name = 'Retain 7 years' } }
            $rows = Invoke-ExchangeScript -Name 'connect-security-compliance' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'Organization', 'RetentionPolicyCount', 'Experimental'
            $rows[0].RetentionPolicyCount | Should -Be 1
            $rows[0].Experimental | Should -BeTrue
            Should -Invoke Connect-MspSecurityCompliance -Times 1 -Exactly
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }
    }

    Context 'find-external-forwarding' {
        BeforeEach {
            Mock Get-HostedOutboundSpamFilterPolicy { [pscustomobject]@{ AutoForwardingMode = 'Automatic' } }
            Mock Get-EXOMailbox {
                [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com'; DisplayName = 'Jane'; ForwardingSmtpAddress = 'smtp:jane.home@fabrikam.example'; ForwardingAddress = $null; DeliverToMailboxAndForward = $true }
                [pscustomobject]@{ PrimarySmtpAddress = 'sam@contoso.com'; DisplayName = 'Sam'; ForwardingSmtpAddress = 'smtp:reception@contoso.com'; ForwardingAddress = 'Partner Contact'; DeliverToMailboxAndForward = $false }
                [pscustomobject]@{ PrimarySmtpAddress = 'kim@contoso.com'; DisplayName = 'Kim'; ForwardingSmtpAddress = $null; ForwardingAddress = $null; DeliverToMailboxAndForward = $false }
            }
            Mock Get-Recipient { [pscustomobject]@{ PrimarySmtpAddress = 'partner@fabrikam.example'; ExternalEmailAddress = 'SMTP:partner@fabrikam.example'; RecipientType = 'MailContact' } }
            Mock Get-InboxRule {
                if ($Mailbox -eq 'kim@contoso.com') {
                    [pscustomobject]@{ Name = 'Send to me'; ForwardTo = @('"Kim Home" [SMTP:kim@fabrikam.example]'); ForwardAsAttachmentTo = $null; RedirectTo = @('"Boss" [EX:/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=boss]') }
                }
            }
        }
        It 'finds mailbox, contact and inbox rule forwards to external domains only' {
            $rows = Invoke-ExchangeScript -Name 'find-external-forwarding' -Parameters @{ TenantId = $script:Tenant; IncludeInboxRules = $true }
            Assert-RowShape -Rows $rows -Column 'Mailbox', 'ForwardingType', 'ExternalRecipient', 'OutboundAutoForwardingMode'
            $rows.Count | Should -Be 3
            ($rows | Where-Object ForwardingType -eq 'ForwardingSmtpAddress').ExternalRecipient | Should -Be 'jane.home@fabrikam.example'
            ($rows | Where-Object ForwardingType -eq 'ForwardingAddress').ExternalRecipient | Should -Be 'partner@fabrikam.example'
            ($rows | Where-Object ForwardingType -eq 'InboxRule').ExternalRecipient | Should -Be 'kim@fabrikam.example'
            $rows | ForEach-Object { $_.OutboundAutoForwardingMode | Should -Be 'Automatic' }
        }
        It 'skips inbox rules unless asked' {
            $rows = Invoke-ExchangeScript -Name 'find-external-forwarding' -Parameters @{ TenantId = $script:Tenant }
            $rows.Count | Should -Be 2
            Should -Invoke Get-InboxRule -Times 0 -Exactly
        }
    }

    Context 'export-mobile-devices' {
        It 'matches devices to mailboxes by distinguished name' {
            Mock Get-EXOMailbox { [pscustomobject]@{ UserPrincipalName = 'jane@contoso.com'; DisplayName = 'Jane'; DistinguishedName = 'CN=Jane,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EXAMPLE,DC=PROD' } }
            Mock Get-MobileDevice { [pscustomobject]@{ FriendlyName = 'iPhone'; DeviceModel = 'iPhone15,2'; DeviceOS = 'iOS 18'; ClientType = 'Outlook'; DeviceAccessState = 'Allowed'; UserDisplayName = 'EXAMPLE\Jane'; DistinguishedName = 'CN=iPhone 123,CN=ExchangeActiveSyncDevices,CN=Jane,OU=contoso.onmicrosoft.com,OU=Microsoft Exchange Hosted Organizations,DC=EXAMPLE,DC=PROD' } }
            $rows = Invoke-ExchangeScript -Name 'export-mobile-devices' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'UserPrincipalName', 'DeviceModel', 'DeviceOS', 'ClientType', 'FirstSyncTime'
            $rows[0].UserPrincipalName | Should -Be 'jane@contoso.com'
            $rows[0].DeviceModel | Should -Be 'iPhone15,2'
        }
    }

    Context 'set-default-calendar-permission' {
        BeforeEach {
            Mock Get-EXOMailbox { [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com'; DisplayName = 'Jane' } }
            Mock Get-EXOMailboxFolderPermission { [pscustomobject]@{ User = 'Default'; AccessRights = @('AvailabilityOnly') } }
            Mock Set-MailboxFolderPermission {}
        }
        It 'reports by default and changes nothing' {
            $rows = Invoke-ExchangeScript -Name 'set-default-calendar-permission' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'Mailbox', 'CalendarFolder', 'CurrentAccessRights', 'DesiredAccessRights'
            $rows[0].Status | Should -Be 'WouldChange'
            $rows[0].CalendarFolder | Should -Be 'jane@contoso.com:\Calendar'
            Should -Invoke Set-MailboxFolderPermission -Times 0 -Exactly
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'set-default-calendar-permission' -Parameters @{ TenantId = $script:Tenant; Apply = $true; WhatIf = $true }
            $rows[0].Status | Should -Be 'WhatIf'
            Should -Invoke Set-MailboxFolderPermission -Times 0 -Exactly
        }
        It 'changes the Default entry with -Apply' {
            $rows = Invoke-ExchangeScript -Name 'set-default-calendar-permission' -Parameters @{ TenantId = $script:Tenant; Apply = $true; Confirm = $false }
            $rows[0].Status | Should -Be 'Changed'
            Should -Invoke Set-MailboxFolderPermission -Times 1 -Exactly -ParameterFilter { $User -eq 'Default' -and $AccessRights -eq 'LimitedDetails' }
        }
        It 'records a mistyped -Mailbox and carries on with the others' {
            Mock Get-EXOMailbox { throw 'not found' } -ParameterFilter { $Identity -eq 'nobody@contoso.com' }
            $rows = Invoke-ExchangeScript -Name 'set-default-calendar-permission' -Parameters @{ TenantId = $script:Tenant; Mailbox = @('nobody@contoso.com', 'jane@contoso.com') }
            $rows.Count | Should -Be 2
            $rows[0].Status | Should -Be 'Failed'
            $rows[0].Mailbox | Should -Be 'nobody@contoso.com'
            $rows[1].Status | Should -Be 'WouldChange'
        }
        It 'finds a calendar folder with a localised name' {
            Mock Get-EXOMailboxFolderPermission { throw 'folder not found' } -ParameterFilter { $Identity -like '*:\Calendar' }
            Mock Get-EXOMailboxFolderPermission { [pscustomobject]@{ User = 'Default'; AccessRights = @('LimitedDetails') } } -ParameterFilter { $Identity -like '*:\Kalender' }
            Mock Get-EXOMailboxFolderStatistics { [pscustomobject]@{ FolderType = 'Calendar'; FolderPath = '/Kalender' } }
            $rows = Invoke-ExchangeScript -Name 'set-default-calendar-permission' -Parameters @{ TenantId = $script:Tenant }
            $rows[0].CalendarFolder | Should -Be 'jane@contoso.com:\Kalender'
            $rows[0].Status | Should -Be 'AlreadySet'
        }
    }

    Context 'add-domain-email-alias' {
        BeforeEach {
            Mock Get-EXOMailbox {
                [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com'; Alias = 'jane'; DisplayName = 'Jane'; EmailAddresses = @('SMTP:jane@contoso.com') }
                [pscustomobject]@{ PrimarySmtpAddress = 'sam@contoso.com'; Alias = 'sam'; DisplayName = 'Sam'; EmailAddresses = @('SMTP:sam@contoso.com', 'smtp:sam@contoso.net') }
                [pscustomobject]@{ PrimarySmtpAddress = 'bob@contoso.com.au'; Alias = 'bob'; DisplayName = 'Bob'; EmailAddresses = @('SMTP:bob@contoso.com.au') }
            }
            Mock Get-EXORecipient { throw 'not found' }
            Mock Set-Mailbox {}
        }
        It 'matches the domain exactly and reports what it would add' {
            $rows = Invoke-ExchangeScript -Name 'add-domain-email-alias' -Parameters @{ TenantId = $script:Tenant; MatchDomain = 'contoso.com'; AliasDomain = 'contoso.net' }
            Assert-RowShape -Rows $rows -Column 'Mailbox', 'NewAddress'
            $rows.Count | Should -Be 2
            ($rows | Where-Object Mailbox -eq 'jane@contoso.com').Status | Should -Be 'WouldAdd'
            ($rows | Where-Object Mailbox -eq 'sam@contoso.com').Status | Should -Be 'AlreadyPresent'
            Should -Invoke Set-Mailbox -Times 0 -Exactly
        }
        It 'makes no change under -Apply -WhatIf' {
            $null = Invoke-ExchangeScript -Name 'add-domain-email-alias' -Parameters @{ TenantId = $script:Tenant; MatchDomain = 'contoso.com'; AliasDomain = 'contoso.net'; Apply = $true; WhatIf = $true }
            Should -Invoke Set-Mailbox -Times 0 -Exactly
        }
        It 'adds the address with -Apply' {
            $null = Invoke-ExchangeScript -Name 'add-domain-email-alias' -Parameters @{ TenantId = $script:Tenant; MatchDomain = 'contoso.com'; AliasDomain = 'contoso.net'; Apply = $true; Confirm = $false }
            Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter { $Identity -eq 'jane@contoso.com' -and $EmailAddresses.Add -eq 'smtp:jane@contoso.net' }
        }
        It 'matches several domains, like the original text match did' {
            $rows = Invoke-ExchangeScript -Name 'add-domain-email-alias' -Parameters @{ TenantId = $script:Tenant; MatchDomain = @('contoso.com', 'contoso.com.au'); AliasDomain = 'contoso.net' }
            $rows.Count | Should -Be 3
            ($rows | Where-Object Mailbox -eq 'bob@contoso.com.au').NewAddress | Should -Be 'bob@contoso.net'
        }
        It 'fails the customer when the alias domain is not accepted' {
            $rows = Invoke-ExchangeScript -Name 'add-domain-email-alias' -Parameters @{ TenantId = $script:Tenant; MatchDomain = 'contoso.com'; AliasDomain = 'contoso.org' }
            $rows[0].Status | Should -Be 'Failed'
        }
    }

    Context 'new-domain-dynamic-distribution-group' {
        BeforeEach {
            Mock Get-Recipient { [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com' }; [pscustomobject]@{ PrimarySmtpAddress = 'sam@contoso.com' } }
            Mock Get-DynamicDistributionGroup { throw 'not found' }
            Mock New-DynamicDistributionGroup { [pscustomobject]@{ Name = $Name } }
        }
        It 'previews members and changes nothing by default' {
            $rows = Invoke-ExchangeScript -Name 'new-domain-dynamic-distribution-group' -Parameters @{ TenantId = $script:Tenant; Domain = 'contoso.com' }
            Assert-RowShape -Rows $rows -Column 'GroupName', 'RecipientFilter', 'PreviewMemberCount'
            $rows[0].Status | Should -Be 'WouldCreate'
            $rows[0].PreviewMemberCount | Should -Be 2
            $rows[0].RecipientFilter | Should -Be "(RecipientTypeDetails -eq 'UserMailbox') -and (WindowsEmailAddress -like '*@contoso.com')"
            Should -Invoke New-DynamicDistributionGroup -Times 0 -Exactly
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'new-domain-dynamic-distribution-group' -Parameters @{ TenantId = $script:Tenant; Domain = 'contoso.com'; Apply = $true; WhatIf = $true }
            $rows[0].Status | Should -Be 'WhatIf'
            Should -Invoke New-DynamicDistributionGroup -Times 0 -Exactly
        }
        It 'creates the group with -Apply' {
            $rows = Invoke-ExchangeScript -Name 'new-domain-dynamic-distribution-group' -Parameters @{ TenantId = $script:Tenant; Domain = 'contoso.com'; Apply = $true; Confirm = $false }
            $rows[0].Status | Should -Be 'Created'
            Should -Invoke New-DynamicDistributionGroup -Times 1 -Exactly -ParameterFilter { $Name -eq 'All Users - contoso.com' }
        }
    }

    Context 'disable-tnef-remote-domain' {
        BeforeEach {
            Mock Get-RemoteDomain { [pscustomobject]@{ Identity = 'Default'; TNEFEnabled = $global:ExoTest.Tnef } }
            Mock Set-RemoteDomain { $global:ExoTest.Tnef = $TNEFEnabled }
        }
        It 'reports by default' {
            $rows = Invoke-ExchangeScript -Name 'disable-tnef-remote-domain' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'RemoteDomain', 'TNEFEnabledBefore', 'TNEFEnabledAfter'
            $rows[0].Status | Should -Be 'WouldChange'
            $rows[0].TNEFEnabledBefore | Should -Be 'NotSet'
            Should -Invoke Set-RemoteDomain -Times 0 -Exactly
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'disable-tnef-remote-domain' -Parameters @{ TenantId = $script:Tenant; Apply = $true; WhatIf = $true }
            $rows[0].Status | Should -Be 'WhatIf'
            Should -Invoke Set-RemoteDomain -Times 0 -Exactly
        }
        It 'sets TNEFEnabled to false and reads it back with -Apply' {
            $rows = Invoke-ExchangeScript -Name 'disable-tnef-remote-domain' -Parameters @{ TenantId = $script:Tenant; Apply = $true; Confirm = $false }
            $rows[0].Status | Should -Be 'Changed'
            $rows[0].TNEFEnabledAfter | Should -Be 'False'
            Should -Invoke Set-RemoteDomain -Times 1 -Exactly -ParameterFilter { $TNEFEnabled -eq $false }
        }
    }

    Context 'set-mobile-device-mailbox-policy' {
        BeforeEach {
            Mock Get-MobileDeviceMailboxPolicy { [pscustomobject]@{ Name = 'Default'; IsDefault = $true; AllowCamera = $true; AllowWiFi = $true; AllowInternetSharing = $true; AllowBrowser = $true } }
            Mock Get-EXOCasMailbox { [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com'; ActiveSyncMailboxPolicy = 'Default' } }
            Mock New-MobileDeviceMailboxPolicy {}
            Mock Set-CASMailbox {}
        }
        It 'reports policies and assignments without -PolicyName' {
            $rows = Invoke-ExchangeScript -Name 'set-mobile-device-mailbox-policy' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'RowType', 'PolicyName', 'Mailbox', 'AllowCamera'
            @($rows | Where-Object RowType -eq 'Policy').Count | Should -Be 1
            ($rows | Where-Object RowType -eq 'Mailbox').PolicyName | Should -Be 'Default'
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'set-mobile-device-mailbox-policy' -Parameters @{ TenantId = $script:Tenant; PolicyName = 'No Camera Policy'; AllowCamera = $false; AssignTo = 'jane@contoso.com'; Apply = $true; WhatIf = $true }
            $rows.Status | Should -Be @('WhatIf', 'WhatIf')
            Should -Invoke New-MobileDeviceMailboxPolicy -Times 0 -Exactly
            Should -Invoke Set-CASMailbox -Times 0 -Exactly
        }
        It 'creates and assigns the policy with -Apply' {
            $rows = Invoke-ExchangeScript -Name 'set-mobile-device-mailbox-policy' -Parameters @{ TenantId = $script:Tenant; PolicyName = 'No Camera Policy'; AllowCamera = $false; AssignTo = 'jane@contoso.com'; Apply = $true; Confirm = $false }
            $rows.Status | Should -Be @('Created', 'Assigned')
            Should -Invoke New-MobileDeviceMailboxPolicy -Times 1 -Exactly -ParameterFilter { $Name -eq 'No Camera Policy' -and $AllowCamera -eq $false }
            Should -Invoke Set-CASMailbox -Times 1 -Exactly
        }
    }

    Context 'enable-mailbox-archive-policy' {
        BeforeEach {
            Mock Get-RetentionPolicyTag { throw 'not found' }
            Mock Get-RetentionPolicy { throw 'not found' }
            Mock Get-EXOMailbox { [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com'; RetentionPolicy = 'Default MRM Policy'; ArchiveStatus = 'None' } }
            Mock Get-EXOMailboxStatistics { [pscustomobject]@{ ItemCount = 10; TotalItemSize = '1 MB' } }
            Mock Enable-OrganizationCustomization {}
            Mock New-RetentionPolicyTag {}
            Mock New-RetentionPolicy {}
            Mock Set-Mailbox {}
            Mock Enable-Mailbox {}
            Mock Start-ManagedFolderAssistant {}
        }
        It 'reports by default' {
            $rows = Invoke-ExchangeScript -Name 'enable-mailbox-archive-policy' -Parameters @{ TenantId = $script:Tenant; Mailbox = 'jane@contoso.com' }
            Assert-RowShape -Rows $rows -Column 'Target', 'CurrentRetentionPolicy', 'ArchiveStatus', 'Actions'
            $rows.Count | Should -Be 2
            $rows[0].Actions | Should -BeLike 'Enable-OrganizationCustomization*'
            $rows[1].Actions | Should -BeLike '*Enable-Mailbox -Archive*'
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'enable-mailbox-archive-policy' -Parameters @{ TenantId = $script:Tenant; Mailbox = 'jane@contoso.com'; Apply = $true; WhatIf = $true }
            $rows.Status | Should -Be @('WhatIf', 'WhatIf')
            foreach ($command in 'Enable-OrganizationCustomization', 'New-RetentionPolicyTag', 'New-RetentionPolicy', 'Set-Mailbox', 'Enable-Mailbox', 'Start-ManagedFolderAssistant') {
                Should -Invoke $command -Times 0 -Exactly
            }
        }
        It 'adds the tag to an existing policy that does not link it' {
            Mock Get-RetentionPolicyTag { [pscustomobject]@{ Name = 'Move to archive after 365 days'; RetentionAction = 'MoveToArchive' } }
            Mock Get-RetentionPolicy { [pscustomobject]@{ Name = 'Archive after 365 days'; RetentionPolicyTagLinks = @('Some other tag') } }
            Mock Set-RetentionPolicy {}
            $rows = Invoke-ExchangeScript -Name 'enable-mailbox-archive-policy' -Parameters @{ TenantId = $script:Tenant; Mailbox = 'jane@contoso.com' }
            $rows[0].Actions | Should -BeLike '*Set-RetentionPolicy*'
            $null = Invoke-ExchangeScript -Name 'enable-mailbox-archive-policy' -Parameters @{ TenantId = $script:Tenant; Mailbox = 'jane@contoso.com'; Apply = $true; WhatIf = $true }
            Should -Invoke Set-RetentionPolicy -Times 0 -Exactly
            $null = Invoke-ExchangeScript -Name 'enable-mailbox-archive-policy' -Parameters @{ TenantId = $script:Tenant; Mailbox = 'jane@contoso.com'; Apply = $true; Confirm = $false }
            Should -Invoke Set-RetentionPolicy -Times 1 -Exactly -ParameterFilter { $RetentionPolicyTagLinks.Add -eq 'Move to archive after 365 days' }
            Should -Invoke New-RetentionPolicy -Times 0 -Exactly
        }
        It 'creates the tag and policy, assigns it and enables the archive with -Apply' {
            $rows = Invoke-ExchangeScript -Name 'enable-mailbox-archive-policy' -Parameters @{ TenantId = $script:Tenant; Mailbox = 'jane@contoso.com'; Apply = $true; Confirm = $false }
            $rows.Status | Should -Be @('Changed', 'Changed')
            Should -Invoke New-RetentionPolicyTag -Times 1 -Exactly -ParameterFilter { $AgeLimitForRetention -eq 365 -and $RetentionAction -eq 'MoveToArchive' }
            Should -Invoke Enable-Mailbox -Times 1 -Exactly -ParameterFilter { $Archive }
            Should -Invoke Start-ManagedFolderAssistant -Times 1 -Exactly
        }
    }

    Context 'manage-microsoft-365-groups' {
        BeforeEach {
            Mock Get-UnifiedGroup { [pscustomobject]@{ DisplayName = 'Contoso Team'; PrimarySmtpAddress = 'team@contoso.com'; AccessType = 'Private'; GroupMemberCount = 1 } }
            Mock Get-UnifiedGroupLinks { [pscustomobject]@{ PrimarySmtpAddress = 'jane@contoso.com' } }
            Mock Add-UnifiedGroupLinks {}
            Mock New-UnifiedGroup {}
        }
        It 'reports groups and members' {
            $rows = Invoke-ExchangeScript -Name 'manage-microsoft-365-groups' -Parameters @{ TenantId = $script:Tenant; IncludeMembers = $true }
            Assert-RowShape -Rows $rows -Column 'GroupName', 'GroupAddress', 'Member'
            $rows[0].Member | Should -Be 'jane@contoso.com'
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'manage-microsoft-365-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'AddMember'; Identity = 'team'; Member = @('jane@contoso.com', 'sam@contoso.com'); Apply = $true; WhatIf = $true }
            $rows.Status | Should -Be @('AlreadyMember', 'WhatIf')
            $null = Invoke-ExchangeScript -Name 'manage-microsoft-365-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'Create'; DisplayName = 'Sales'; Alias = 'sales'; Apply = $true; WhatIf = $true }
            Should -Invoke Add-UnifiedGroupLinks -Times 0 -Exactly
            Should -Invoke New-UnifiedGroup -Times 0 -Exactly
        }
        It 'adds only missing members with -Apply' {
            $null = Invoke-ExchangeScript -Name 'manage-microsoft-365-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'AddMember'; Identity = 'team'; Member = @('jane@contoso.com', 'sam@contoso.com'); Apply = $true; Confirm = $false }
            Should -Invoke Add-UnifiedGroupLinks -Times 1 -Exactly -ParameterFilter { $Links -eq 'sam@contoso.com' }
        }
        It 'creates a public group when asked' {
            Mock Get-UnifiedGroup { throw 'not found' } -ParameterFilter { $Identity -eq 'sales' }
            Mock New-UnifiedGroup { [pscustomobject]@{ PrimarySmtpAddress = 'sales@contoso.com' } }
            $rows = Invoke-ExchangeScript -Name 'manage-microsoft-365-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'Create'; DisplayName = 'Sales'; Alias = 'sales'; AccessType = 'Public'; Apply = $true; Confirm = $false }
            $rows[0].Status | Should -Be 'Created'
            Should -Invoke New-UnifiedGroup -Times 1 -Exactly -ParameterFilter { $AccessType -eq 'Public' }
        }
        It 'refuses changes across -AllCustomers' {
            { Invoke-ExchangeScript -Name 'manage-microsoft-365-groups' -Parameters @{ AllCustomers = $true; Action = 'AddMember'; Identity = 'team'; Member = 'sam@contoso.com' } } | Should -Throw '*one customer at a time*'
        }
    }

    Context 'manage-external-contact-groups' {
        BeforeEach {
            Mock Get-DistributionGroup { [pscustomobject]@{ DisplayName = 'Suppliers'; PrimarySmtpAddress = 'suppliers@contoso.com' } }
            Mock Get-DistributionGroupMember { [pscustomobject]@{ DisplayName = 'Alex'; RecipientType = 'MailContact'; RecipientTypeDetails = 'MailContact'; PrimarySmtpAddress = 'alex@fabrikam.example'; ExternalEmailAddress = 'SMTP:alex@fabrikam.example' } }
            Mock Get-MailContact { [pscustomobject]@{ DisplayName = 'Alex'; Identity = 'Alex' } }
            Mock New-MailContact { [pscustomobject]@{ DisplayName = $Name; Identity = $Name } }
            Mock Add-DistributionGroupMember {}
            Mock Remove-DistributionGroupMember {}
            Mock Remove-MailContact {}
            Mock New-DistributionGroup { [pscustomobject]@{ PrimarySmtpAddress = 'board@contoso.com' } }
        }
        It 'reports groups and external members' {
            $rows = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'Group', 'Member', 'MemberType', 'ExternalAddress'
            $rows[0].ExternalAddress | Should -Be 'alex@fabrikam.example'
        }
        It 'makes no change under -Apply -WhatIf' {
            Mock Get-MailContact { throw 'not found' }
            $null = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'AddContact'; ContactEmail = 'pat@fabrikam.example'; GroupIdentity = 'Suppliers'; Apply = $true; WhatIf = $true }
            Mock Get-MailContact { [pscustomobject]@{ DisplayName = 'Alex'; Identity = 'Alex' } }
            $null = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'RemoveContact'; ContactEmail = 'alex@fabrikam.example'; Apply = $true; WhatIf = $true }
            Mock Get-DistributionGroup { throw 'not found' } -ParameterFilter { $Identity -eq 'Board' }
            $null = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'NewGroup'; GroupName = 'Board'; Apply = $true; WhatIf = $true }
            foreach ($command in 'New-MailContact', 'Add-DistributionGroupMember', 'Remove-DistributionGroupMember', 'New-DistributionGroup') {
                Should -Invoke $command -Times 0 -Exactly
            }
        }
        It 'creates the contact and adds it with -Apply' {
            Mock Get-MailContact { throw 'not found' }
            $rows = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'AddContact'; ContactEmail = 'pat@fabrikam.example'; ContactName = 'Pat'; GroupIdentity = 'Suppliers'; Apply = $true; Confirm = $false }
            $rows.Status | Should -Be @('Created', 'Added')
            Should -Invoke New-MailContact -Times 1 -Exactly -ParameterFilter { $ExternalEmailAddress -eq 'pat@fabrikam.example' }
            Should -Invoke Add-DistributionGroupMember -Times 1 -Exactly
        }
        It 'removes the contact from every group with -Apply' {
            $rows = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'RemoveContact'; ContactEmail = 'alex@fabrikam.example'; Apply = $true; Confirm = $false }
            $rows[0].Status | Should -Be 'Removed'
            Should -Invoke Remove-DistributionGroupMember -Times 1 -Exactly
            Should -Invoke Remove-MailContact -Times 0 -Exactly
        }
        It 'deletes the contact only with -DeleteContact and never under -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'RemoveContact'; ContactEmail = 'alex@fabrikam.example'; DeleteContact = $true; Apply = $true; WhatIf = $true }
            $rows.Status | Should -Be @('WhatIf', 'WhatIf')
            Should -Invoke Remove-MailContact -Times 0 -Exactly
            $rows = Invoke-ExchangeScript -Name 'manage-external-contact-groups' -Parameters @{ TenantId = $script:Tenant; Action = 'RemoveContact'; ContactEmail = 'alex@fabrikam.example'; DeleteContact = $true; Apply = $true; Confirm = $false }
            $rows.Status | Should -Be @('Removed', 'Deleted')
            Should -Invoke Remove-MailContact -Times 1 -Exactly -ParameterFilter { $Identity -eq 'Alex' }
        }
    }

    Context 'set-owa-conditional-access' {
        BeforeEach {
            Mock Get-OwaMailboxPolicy { [pscustomobject]@{ Name = 'OwaMailboxPolicy-Default'; ConditionalAccessPolicy = 'Off' } }
            Mock Set-OwaMailboxPolicy {}
        }
        It 'reports the OWA setting and the Conditional Access policies' {
            Mock Invoke-MspGraphRequest {
                [pscustomobject]@{ displayName = 'OWA limited'; state = 'enabled'; conditions = [pscustomobject]@{ applications = [pscustomobject]@{ includeApplications = @('Office365') } }; sessionControls = [pscustomobject]@{ applicationEnforcedRestrictions = [pscustomobject]@{ isEnabled = $true } } }
                [pscustomobject]@{ displayName = 'Require MFA'; state = 'enabled'; conditions = [pscustomobject]@{ applications = [pscustomobject]@{ includeApplications = @('All') } }; sessionControls = $null }
            } -ParameterFilter { $Method -eq 'GET' }
            $rows = Invoke-ExchangeScript -Name 'set-owa-conditional-access' -Parameters @{ TenantId = $script:Tenant }
            Assert-RowShape -Rows $rows -Column 'Setting', 'Name', 'CurrentValue', 'DesiredValue'
            ($rows | Where-Object Setting -eq 'ConditionalAccessPolicy').Name | Should -Be 'OWA limited'
            ($rows | Where-Object Setting -eq 'OwaMailboxPolicy').Status | Should -Be 'WouldChange'
            Should -Invoke Set-OwaMailboxPolicy -Times 0 -Exactly
        }
        It 'makes no change under -Apply -WhatIf' {
            $rows = Invoke-ExchangeScript -Name 'set-owa-conditional-access' -Parameters @{ TenantId = $script:Tenant; CreateConditionalAccessPolicy = $true; Apply = $true; WhatIf = $true }
            $rows.Status | Should -Be @('WhatIf', 'WhatIf')
            Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -ne 'GET' }
            Should -Invoke Set-OwaMailboxPolicy -Times 0 -Exactly
        }
        It 'creates a report-only Conditional Access policy and sets OWA with -Apply' {
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = 'new'; state = 'enabledForReportingButNotEnforced' } } -ParameterFilter { $Method -eq 'POST' }
            $rows = Invoke-ExchangeScript -Name 'set-owa-conditional-access' -Parameters @{ TenantId = $script:Tenant; CreateConditionalAccessPolicy = $true; Apply = $true; Confirm = $false }
            $rows.Status | Should -Be @('Created', 'Changed')
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter {
                $Method -eq 'POST' -and $Uri -eq 'v1.0/identity/conditionalAccess/policies' -and $Body.state -eq 'enabledForReportingButNotEnforced' -and $Body.sessionControls.applicationEnforcedRestrictions.isEnabled
            }
            Should -Invoke Set-OwaMailboxPolicy -Times 1 -Exactly -ParameterFilter { $ConditionalAccessPolicy -eq 'ReadOnly' }
        }
    }

    Context 'get-privileged-access-management' {
        It 'reports the organisation setting, approval policies and requests' {
            Mock Get-ElevatedAccessApprovalPolicy { [pscustomobject]@{ Task = 'Exchange\New-JournalRule'; ApprovalType = 'Manual'; ApproverGroup = 'pam-approvers@contoso.com' } }
            Mock Get-ElevatedAccessRequest { [pscustomobject]@{ Task = 'Exchange\New-JournalRule'; RequestStatus = 'Approved'; Requestor = 'admin@contoso.com' } }
            $rows = Invoke-ExchangeScript -Name 'get-privileged-access-management' -Parameters @{ TenantId = $script:Tenant; IncludeRequests = $true }
            Assert-RowShape -Rows $rows -Column 'RowType', 'Task', 'Value', 'Detail'
            $rows.RowType | Should -Be @('Organisation', 'ApprovalPolicy', 'Request')
            $rows[0].Value | Should -Be 'Enabled'
            $rows[1].Value | Should -Be 'Manual'
            $rows[1].Detail | Should -BeLike '*ApproverGroup=pam-approvers@contoso.com*'
            $rows[2].Value | Should -Be 'Approved'
        }
    }
}
