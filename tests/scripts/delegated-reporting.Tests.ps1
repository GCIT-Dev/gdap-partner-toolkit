#Requires -Version 7.4
# Offline smoke tests for scripts/delegated-reporting. Every MspGdap call that would reach a tenant,
# every Exchange Online cmdlet and every Security and Compliance cmdlet is mocked or stubbed.
# Nothing here contacts Microsoft or any customer tenant.

BeforeAll {
    $script:RepoRoot = (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../..')).Path
    $script:ScriptRoot = Join-Path -Path $script:RepoRoot -ChildPath 'scripts/delegated-reporting'
    $srcPath = Join-Path -Path $script:RepoRoot -ChildPath 'src'
    if (($env:PSModulePath -split [System.IO.Path]::PathSeparator) -notcontains $srcPath) {
        $env:PSModulePath = $srcPath + [System.IO.Path]::PathSeparator + $env:PSModulePath
    }
    Import-Module (Join-Path -Path $srcPath -ChildPath 'MspGdap/MspGdap.psd1') -Force

    $script:CustomerId = '22222222-2222-4222-8222-222222222222'
    $script:Tenant = 'contoso.onmicrosoft.com'

    # Scripts owned by this test file. Other script groups in the same folder have their own tests.
    $script:OwnScripts = @(
        'add-trusted-sender.ps1'
        'audit-connection-filter-ip-allow-list.ps1'
        'disable-pop-imap.ps1'
        'disable-winmail-dat-tnef.ps1'
        'elevation-of-privilege-alert-policy.ps1'
        'enable-unified-audit-log.ps1'
        'export-sign-in-locations.ps1'
        'find-external-forwarding-inbox-rules.ps1'
        'find-external-forwarding-mailboxes.ps1'
        'find-licensed-shared-mailboxes.ps1'
        'find-mailboxes-below-plan-quota.ps1'
        'grant-calendar-access.ps1'
    )

    function Get-ScriptPath {
        param([Parameter(Mandatory)][string]$Name)
        Join-Path -Path $script:ScriptRoot -ChildPath $Name
    }

    # Stubs for Exchange Online and Security and Compliance cmdlets, so the tests run without
    # ExchangeOnlineManagement and so Mock has a command to bind to.
    function Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param() }
    function Get-AdminAuditLogConfig { [CmdletBinding()] param() }
    function Set-AdminAuditLogConfig { [CmdletBinding()] param($UnifiedAuditLogIngestionEnabled) }
    function Get-OrganizationConfig { [CmdletBinding()] param() }
    function Enable-OrganizationCustomization { [CmdletBinding()] param() }
    function Get-RemoteDomain { [CmdletBinding()] param($Identity) }
    function Set-RemoteDomain { [CmdletBinding()] param($Identity, $TNEFEnabled) }
    function Get-CASMailboxPlan { [CmdletBinding()] param($Filter) }
    function Set-CASMailboxPlan { [CmdletBinding()] param($Identity, $ImapEnabled, $PopEnabled) }
    function Get-EXOCasMailbox { [CmdletBinding()] param($Filter, $ResultSize, $PropertySets, $Properties) }
    function Set-CASMailbox { [CmdletBinding()] param($Identity, $ImapEnabled, $PopEnabled) }
    function Get-EXOMailbox { [CmdletBinding()] param($Identity, $ResultSize, $RecipientTypeDetails, $Properties, $PropertySets, $Filter) }
    function Get-EXOMailboxStatistics { [CmdletBinding()] param($Identity, $ExternalDirectoryObjectId, $Properties, $PropertySets) }
    function Get-EXOMailboxFolderStatistics { [CmdletBinding()] param($Identity, $FolderScope) }
    function Get-EXOMailboxFolderPermission { [CmdletBinding()] param($Identity, $User) }
    function Add-MailboxFolderPermission { [CmdletBinding()] param($Identity, $User, $AccessRights) }
    function Set-MailboxFolderPermission { [CmdletBinding()] param($Identity, $User, $AccessRights) }
    function Get-MailboxJunkEmailConfiguration { [CmdletBinding()] param($Identity) }
    function Set-MailboxJunkEmailConfiguration { [CmdletBinding()] param($Identity, $TrustedSendersAndDomains) }
    function Get-AcceptedDomain { [CmdletBinding()] param() }
    function Get-InboxRule { [CmdletBinding()] param($Mailbox, $IncludeHidden) }
    function Get-Recipient { [CmdletBinding()] param($Identity) }
    function Get-HostedOutboundSpamFilterPolicy { [CmdletBinding()] param($Identity) }
    function Get-HostedConnectionFilterPolicy { [CmdletBinding()] param($Identity) }
    function Set-HostedConnectionFilterPolicy { [CmdletBinding()] param($Identity, $IPAllowList) }
    function Get-MailboxPlan { [CmdletBinding()] param($Identity) }
    function Set-Mailbox { [CmdletBinding()] param($Identity, $ProhibitSendReceiveQuota, $ProhibitSendQuota, $IssueWarningQuota) }
    function Get-ProtectionAlert { [CmdletBinding()] param($Identity) }
    function New-ProtectionAlert { [CmdletBinding()] param($Name, $Category, $NotifyUser, $ThreatType, $Operation, $AggregationType, $Severity, $Description) }

    function Assert-CustomerShape {
        param([Parameter(Mandatory)][object[]]$Rows)
        $Rows.Count | Should -BeGreaterThan 0
        foreach ($row in $Rows) {
            $row | Should -BeOfType [pscustomobject]
            $row.PSObject.Properties.Name | Should -Contain 'CustomerTenantId'
            $row.PSObject.Properties.Name | Should -Contain 'CustomerName'
            $row.PSObject.Properties.Name | Should -Contain 'Status'
            $row.CustomerTenantId | Should -Be $script:CustomerId
            $row.CustomerName | Should -Be 'Contoso'
        }
    }
}

Describe 'delegated-reporting scripts (mocked)' {
    BeforeEach {
        Mock Get-MspCustomer {
            [pscustomobject]@{ TenantId = '22222222-2222-4222-8222-222222222222'; DisplayName = 'Contoso'; DefaultDomainName = 'contoso.onmicrosoft.com'; GdapStatus = 'active' }
        }
        Mock Connect-MspExchangeOnline { }
        Mock Connect-MspSecurityCompliance { }
        Mock Disconnect-ExchangeOnline { }
        Mock Get-MspAuthHeader { throw 'Get-MspAuthHeader must not be called directly in these tests.' }
        Mock Invoke-MspGraphRequest { }
        Mock Invoke-RestMethod { throw 'Network calls are not allowed in these tests.' }
        Mock Invoke-WebRequest { throw 'Network calls are not allowed in these tests.' }
    }

    Context 'enable-unified-audit-log.ps1' {
        BeforeEach {
            Mock Get-AdminAuditLogConfig { [pscustomobject]@{ UnifiedAuditLogIngestionEnabled = $false } }
            Mock Get-OrganizationConfig { [pscustomobject]@{ IsDehydrated = $true } }
            Mock Enable-OrganizationCustomization { }
            Mock Set-AdminAuditLogConfig { }
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ id = 'u1'; userPrincipalName = 'auditadmin1@contoso.onmicrosoft.com' } }
        }

        It 'reports only by default' {
            $rows = @(& (Get-ScriptPath 'enable-unified-audit-log.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'ReportOnly'
            Should -Invoke Set-AdminAuditLogConfig -Times 0 -Exactly
            Should -Invoke Connect-MspExchangeOnline -Times 1 -Exactly
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'enable-unified-audit-log.ps1') -TenantId $script:Tenant -Apply -LegacyAdminPrefix 'auditadmin' -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].LegacyAdminAccounts | Should -Be 'auditadmin1@contoso.onmicrosoft.com'
            Should -Invoke Set-AdminAuditLogConfig -Times 0 -Exactly
            Should -Invoke Enable-OrganizationCustomization -Times 0 -Exactly
        }

        It 'turns ingestion on with -Apply' {
            $rows = @(& (Get-ScriptPath 'enable-unified-audit-log.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            Assert-CustomerShape -Rows $rows
            Should -Invoke Enable-OrganizationCustomization -Times 1 -Exactly
            Should -Invoke Set-AdminAuditLogConfig -Times 1 -Exactly
        }

        It 'records a failure and continues to the next customer' {
            Mock Connect-MspExchangeOnline { throw 'AADSTS50020 simulated' } -ParameterFilter { $TenantId -eq '22222222-2222-4222-8222-222222222222' }
            $rows = @(& (Get-ScriptPath 'enable-unified-audit-log.ps1') -TenantId $script:Tenant, $script:Tenant)
            $rows.Count | Should -Be 2
            $rows | ForEach-Object { $_.Status | Should -Be 'Failed' }
        }

        It 'writes a CSV with -OutputPath' {
            $csv = Join-Path -Path $TestDrive -ChildPath 'ual.csv'
            $null = & (Get-ScriptPath 'enable-unified-audit-log.ps1') -AllCustomers -OutputPath $csv
            Test-Path -LiteralPath $csv | Should -BeTrue
            (Import-Csv -LiteralPath $csv)[0].CustomerName | Should -Be 'Contoso'
        }
    }

    Context 'disable-winmail-dat-tnef.ps1' {
        BeforeEach {
            Mock Get-RemoteDomain { [pscustomobject]@{ Identity = 'Default'; TNEFEnabled = $null } }
            Mock Set-RemoteDomain { }
        }

        It 'reports only by default' {
            $rows = @(& (Get-ScriptPath 'disable-winmail-dat-tnef.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'ReportOnly'
            Should -Invoke Set-RemoteDomain -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'disable-winmail-dat-tnef.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Set-RemoteDomain -Times 0 -Exactly
        }

        It 'calls Set-RemoteDomain with -Apply' {
            $null = & (Get-ScriptPath 'disable-winmail-dat-tnef.ps1') -TenantId $script:Tenant -Apply -Confirm:$false
            Should -Invoke Set-RemoteDomain -Times 1 -Exactly -ParameterFilter { $TNEFEnabled -eq $false -and $Identity -eq 'Default' }
        }
    }

    Context 'disable-pop-imap.ps1' {
        BeforeEach {
            Mock Get-CASMailboxPlan { [pscustomobject]@{ Identity = 'ExchangeOnlineEnterprise-1'; PopEnabled = $true; ImapEnabled = $true } }
            Mock Get-EXOCasMailbox {
                [pscustomobject]@{ PrimarySmtpAddress = 'user@contoso.com'; PopEnabled = $false; ImapEnabled = $true }
                [pscustomobject]@{ PrimarySmtpAddress = 'scanner@contoso.com'; PopEnabled = $true; ImapEnabled = $false }
            }
            Mock Set-CASMailboxPlan { }
            Mock Set-CASMailbox { }
        }

        It 'reports plans and mailboxes without changing them' {
            $rows = @(& (Get-ScriptPath 'disable-pop-imap.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 3
            @($rows | Where-Object ObjectType -eq 'MailboxPlan').Count | Should -Be 1
            Should -Invoke Set-CASMailboxPlan -Times 0 -Exactly
            Should -Invoke Set-CASMailbox -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'disable-pop-imap.ps1') -TenantId $script:Tenant -ExcludeMailbox 'scanner@contoso.com' -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            ($rows | Where-Object Identity -eq 'scanner@contoso.com').Action | Should -Be 'Excluded'
            ($rows | Where-Object Identity -eq 'user@contoso.com').Action | Should -Be 'WhatIf'
            Should -Invoke Set-CASMailboxPlan -Times 0 -Exactly
            Should -Invoke Set-CASMailbox -Times 0 -Exactly
        }

        It 'changes plans and non-excluded mailboxes with -Apply' {
            $null = & (Get-ScriptPath 'disable-pop-imap.ps1') -TenantId $script:Tenant -ExcludeMailbox 'scanner@contoso.com' -Apply -Confirm:$false
            Should -Invoke Set-CASMailboxPlan -Times 1 -Exactly
            Should -Invoke Set-CASMailbox -Times 1 -Exactly -ParameterFilter { $Identity -eq 'user@contoso.com' }
        }
    }

    Context 'find-licensed-shared-mailboxes.ps1' {
        BeforeEach {
            Mock Invoke-MspGraphRequest { [pscustomobject]@{ skuId = 'sku-1'; skuPartNumber = 'O365_BUSINESS_PREMIUM' } } -ParameterFilter { $Uri -like 'v1.0/subscribedSkus*' }
            Mock Invoke-MspGraphRequest {
                [pscustomobject]@{
                    id                      = 'small-id'
                    userPrincipalName       = 'reception@contoso.com'
                    assignedLicenses        = @([pscustomobject]@{ skuId = 'sku-1' })
                    licenseAssignmentStates = @([pscustomobject]@{ skuId = 'sku-1'; assignedByGroup = $null; state = 'Active' })
                }
            } -ParameterFilter { $Method -eq 'GET' -and $Uri -like 'v1.0/users/small-id*' }
            Mock Invoke-MspGraphRequest {
                [pscustomobject]@{
                    id                      = 'held-id'
                    userPrincipalName       = 'records@contoso.com'
                    assignedLicenses        = @([pscustomobject]@{ skuId = 'sku-1' })
                    licenseAssignmentStates = @([pscustomobject]@{ skuId = 'sku-1'; assignedByGroup = $null; state = 'Active' })
                }
            } -ParameterFilter { $Method -eq 'GET' -and $Uri -like 'v1.0/users/held-id*' }
            Mock Invoke-MspGraphRequest { } -ParameterFilter { $Method -eq 'POST' }
            Mock Get-EXOMailbox {
                [pscustomobject]@{ DisplayName = 'Reception'; PrimarySmtpAddress = 'reception@contoso.com'; ExternalDirectoryObjectId = 'small-id'; ArchiveStatus = 'None'; LitigationHoldEnabled = $false; InPlaceHolds = @(); RetentionHoldEnabled = $false }
                [pscustomobject]@{ DisplayName = 'Records'; PrimarySmtpAddress = 'records@contoso.com'; ExternalDirectoryObjectId = 'held-id'; ArchiveStatus = 'Active'; LitigationHoldEnabled = $true; InPlaceHolds = @(); RetentionHoldEnabled = $false }
            }
            Mock Get-EXOMailboxStatistics { [pscustomobject]@{ TotalItemSize = '1.5 GB (1,610,612,736 bytes)' } }
        }

        It 'reports licences and the reasons a licence must stay' {
            $rows = @(& (Get-ScriptPath 'find-licensed-shared-mailboxes.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            $small = $rows | Where-Object PrimarySmtpAddress -eq 'reception@contoso.com'
            $small.SafeToRemove | Should -BeTrue
            $small.Licences | Should -Be 'O365_BUSINESS_PREMIUM'
            $small.MailboxSizeGB | Should -Be 1.5
            $held = $rows | Where-Object PrimarySmtpAddress -eq 'records@contoso.com'
            $held.SafeToRemove | Should -BeFalse
            $held.BlockingReasons | Should -Match 'Litigation hold'
            Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'POST' }
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'find-licensed-shared-mailboxes.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            ($rows | Where-Object PrimarySmtpAddress -eq 'reception@contoso.com').Action | Should -Be 'WhatIf'
            ($rows | Where-Object PrimarySmtpAddress -eq 'records@contoso.com').Action | Should -Be 'KeepLicence'
            Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'POST' }
        }

        It 'removes only the safe licence with -Apply' {
            $null = & (Get-ScriptPath 'find-licensed-shared-mailboxes.ps1') -TenantId $script:Tenant -Apply -Confirm:$false
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'POST' -and $Uri -eq 'v1.0/users/small-id/assignLicense' }
        }
    }

    Context 'grant-calendar-access.ps1' {
        BeforeEach {
            Mock Get-EXOMailbox { [pscustomobject]@{ PrimarySmtpAddress = 'reception@contoso.com' } } -ParameterFilter { $Identity }
            Mock Get-EXOMailbox {
                [pscustomobject]@{ PrimarySmtpAddress = 'reception@contoso.com' }
                [pscustomobject]@{ PrimarySmtpAddress = 'ceo@contoso.com' }
            } -ParameterFilter { -not $Identity }
            Mock Get-EXOMailboxFolderStatistics { [pscustomobject]@{ FolderType = 'Calendar'; Name = 'Calendar'; FolderPath = '/Calendar' } }
            Mock Get-EXOMailboxFolderPermission { throw 'There is no existing permission entry found for user.' }
            Mock Add-MailboxFolderPermission { }
            Mock Set-MailboxFolderPermission { }
        }

        It 'reports only by default and skips the user''s own mailbox' {
            $rows = @(& (Get-ScriptPath 'grant-calendar-access.ps1') -TenantId $script:Tenant -User 'reception@contoso.com' -AccessRight Reviewer)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].Mailbox | Should -Be 'ceo@contoso.com'
            $rows[0].CalendarFolder | Should -Be 'ceo@contoso.com:\Calendar'
            $rows[0].Action | Should -Be 'ReportOnly'
            Should -Invoke Add-MailboxFolderPermission -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'grant-calendar-access.ps1') -TenantId $script:Tenant -User 'reception@contoso.com' -AccessRight Reviewer -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Add-MailboxFolderPermission -Times 0 -Exactly
            Should -Invoke Set-MailboxFolderPermission -Times 0 -Exactly
        }

        It 'adds the permission with -Apply' {
            $null = & (Get-ScriptPath 'grant-calendar-access.ps1') -TenantId $script:Tenant -User 'reception@contoso.com' -AccessRight Reviewer -Apply -Confirm:$false
            Should -Invoke Add-MailboxFolderPermission -Times 1 -Exactly -ParameterFilter { $AccessRights -eq 'Reviewer' -and $Identity -eq 'ceo@contoso.com:\Calendar' }
        }
    }

    Context 'elevation-of-privilege-alert-policy.ps1' {
        BeforeEach {
            Mock Get-ProtectionAlert { [pscustomobject]@{ Name = 'Elevation of Exchange admin privilege'; Disabled = $false; NotifyUser = @('TenantAdmins') } }
            Mock New-ProtectionAlert { }
        }

        It 'reports the default policy and uses Security and Compliance PowerShell' {
            $rows = @(& (Get-ScriptPath 'elevation-of-privilege-alert-policy.ps1') -TenantId $script:Tenant -NotifyUser 'alerts@contoso.com')
            Assert-CustomerShape -Rows $rows
            $rows[0].DefaultPolicyPresent | Should -BeTrue
            $rows[0].CustomPolicyPresent | Should -BeFalse
            Should -Invoke Connect-MspSecurityCompliance -Times 1 -Exactly
            Should -Invoke New-ProtectionAlert -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'elevation-of-privilege-alert-policy.ps1') -TenantId $script:Tenant -NotifyUser 'alerts@contoso.com' -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke New-ProtectionAlert -Times 0 -Exactly
        }

        It 'creates the policy with -Apply' {
            $null = & (Get-ScriptPath 'elevation-of-privilege-alert-policy.ps1') -TenantId $script:Tenant -NotifyUser 'alerts@contoso.com' -Apply -Confirm:$false
            Should -Invoke New-ProtectionAlert -Times 1 -Exactly -ParameterFilter { $ThreatType -eq 'Activity' -and $AggregationType -eq 'None' -and $Operation -eq 'Add-RoleGroupMember' }
        }
    }

    Context 'add-trusted-sender.ps1' {
        BeforeEach {
            Mock Get-EXOMailbox { [pscustomobject]@{ PrimarySmtpAddress = 'user@contoso.com' } }
            Mock Get-MailboxJunkEmailConfiguration { [pscustomobject]@{ TrustedSendersAndDomains = @('partner@example.com') } }
            Mock Set-MailboxJunkEmailConfiguration { }
        }

        It 'reports missing entries' {
            $rows = @(& (Get-ScriptPath 'add-trusted-sender.ps1') -TenantId $script:Tenant -TrustedSender 'support@fabrikam.com')
            Assert-CustomerShape -Rows $rows
            $rows[0].MissingEntries | Should -Be 'support@fabrikam.com'
            Should -Invoke Set-MailboxJunkEmailConfiguration -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'add-trusted-sender.ps1') -TenantId $script:Tenant -TrustedSender 'support@fabrikam.com' -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Set-MailboxJunkEmailConfiguration -Times 0 -Exactly
        }

        It 'adds the entry with -Apply' {
            $null = & (Get-ScriptPath 'add-trusted-sender.ps1') -TenantId $script:Tenant -TrustedSender 'support@fabrikam.com' -Apply -Confirm:$false
            Should -Invoke Set-MailboxJunkEmailConfiguration -Times 1 -Exactly
        }
    }

    Context 'find-external-forwarding-inbox-rules.ps1' {
        BeforeEach {
            Mock Get-AcceptedDomain { [pscustomobject]@{ DomainName = 'contoso.com' } }
            Mock Get-HostedOutboundSpamFilterPolicy { [pscustomobject]@{ AutoForwardingMode = 'Automatic' } }
            Mock Get-EXOMailbox { [pscustomobject]@{ PrimarySmtpAddress = 'user@contoso.com'; DisplayName = 'Contoso User' } }
            Mock Get-InboxRule {
                [pscustomobject]@{ Name = 'fwd'; Identity = 'user\1'; Enabled = $true; Description = "If the message:`r`n`tForward the message to 'Outside'  "; ForwardTo = @('"Outside" [SMTP:someone@fabrikam.example]', '"Inside" [SMTP:colleague@contoso.com]'); ForwardAsAttachmentTo = $null; RedirectTo = $null }
                [pscustomobject]@{ Name = 'internal'; Identity = 'user\2'; Enabled = $true; ForwardTo = @('"Inside" [EX:/o=ExchangeLabs/cn=Recipients/cn=abc]'); ForwardAsAttachmentTo = $null; RedirectTo = $null }
            }
        }

        It 'reports only rules with external recipients' {
            $rows = @(& (Get-ScriptPath 'find-external-forwarding-inbox-rules.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].ExternalRecipients | Should -Be 'someone@fabrikam.example'
            $rows[0].ForwardType | Should -Be 'ForwardTo'
            $rows[0].AutoForwardingMode | Should -Be 'Automatic'
            $rows[0].MailboxDisplayName | Should -Be 'Contoso User'
            $rows[0].RuleDescription | Should -BeLike 'If the message:*Outside*'
        }
    }

    Context 'find-external-forwarding-mailboxes.ps1' {
        BeforeEach {
            Mock Get-AcceptedDomain { [pscustomobject]@{ DomainName = 'contoso.com' } }
            Mock Get-HostedOutboundSpamFilterPolicy { [pscustomobject]@{ AutoForwardingMode = 'On' } }
            Mock Get-EXOMailbox {
                [pscustomobject]@{ PrimarySmtpAddress = 'a@contoso.com'; DisplayName = 'User A'; ForwardingSmtpAddress = 'smtp:a@fabrikam.example'; ForwardingAddress = $null; DeliverToMailboxAndForward = $true }
                [pscustomobject]@{ PrimarySmtpAddress = 'b@contoso.com'; ForwardingSmtpAddress = $null; ForwardingAddress = 'External Contact'; DeliverToMailboxAndForward = $false }
                [pscustomobject]@{ PrimarySmtpAddress = 'c@contoso.com'; ForwardingSmtpAddress = 'smtp:c2@contoso.com'; ForwardingAddress = $null; DeliverToMailboxAndForward = $false }
            }
            Mock Get-Recipient { [pscustomobject]@{ ExternalEmailAddress = 'SMTP:b@fabrikam.example'; PrimarySmtpAddress = 'b@fabrikam.example'; RecipientTypeDetails = 'MailContact' } }
        }

        It 'reports both kinds of external forwarding and ignores internal forwards' {
            $rows = @(& (Get-ScriptPath 'find-external-forwarding-mailboxes.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            ($rows | Where-Object Mailbox -eq 'a@contoso.com').ExternalRecipient | Should -Be 'a@fabrikam.example'
            ($rows | Where-Object Mailbox -eq 'a@contoso.com').MailboxDisplayName | Should -Be 'User A'
            ($rows | Where-Object Mailbox -eq 'b@contoso.com').ExternalRecipient | Should -Be 'b@fabrikam.example'
        }
    }

    Context 'audit-connection-filter-ip-allow-list.ps1' {
        BeforeEach {
            Mock Get-HostedConnectionFilterPolicy { [pscustomobject]@{ Identity = 'Default'; EnableSafeList = $false; IPAllowList = @('192.0.2.10', '198.51.100.0/24') } }
            Mock Set-HostedConnectionFilterPolicy { }
        }

        It 'lists every entry' {
            $rows = @(& (Get-ScriptPath 'audit-connection-filter-ip-allow-list.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            Should -Invoke Set-HostedConnectionFilterPolicy -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'audit-connection-filter-ip-allow-list.ps1') -TenantId $script:Tenant -RemoveIpAddress '192.0.2.10' -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            ($rows | Where-Object IPAllowListEntry -eq '192.0.2.10').Action | Should -Be 'WhatIf'
            Should -Invoke Set-HostedConnectionFilterPolicy -Times 0 -Exactly
        }

        It 'removes only the chosen entry with -Apply' {
            $null = & (Get-ScriptPath 'audit-connection-filter-ip-allow-list.ps1') -TenantId $script:Tenant -RemoveIpAddress '192.0.2.10' -Apply -Confirm:$false
            Should -Invoke Set-HostedConnectionFilterPolicy -Times 1 -Exactly -ParameterFilter { @($IPAllowList.Remove) -contains '192.0.2.10' -and @($IPAllowList.Remove).Count -eq 1 }
        }
    }

    Context 'export-sign-in-locations.ps1' {
        BeforeEach {
            Mock Invoke-MspGraphRequest {
                [pscustomobject]@{ createdDateTime = '2026-10-01T01:00:00Z'; userPrincipalName = 'user@contoso.com'; ipAddress = '203.0.113.5'; clientAppUsed = 'Browser'; appDisplayName = 'Office 365'; deviceDetail = [pscustomobject]@{ operatingSystem = 'Windows 10'; browser = 'Edge 128.0' }; location = [pscustomobject]@{ city = 'Gold Coast'; state = 'Queensland'; countryOrRegion = 'AU' }; status = [pscustomobject]@{ errorCode = 0 } }
                [pscustomobject]@{ createdDateTime = '2026-10-02T01:00:00Z'; userPrincipalName = 'user@contoso.com'; ipAddress = '203.0.113.5'; clientAppUsed = 'Mobile Apps and Desktop clients'; appDisplayName = 'Outlook'; location = [pscustomobject]@{ city = 'Gold Coast'; state = 'Queensland'; countryOrRegion = 'AU' }; status = [pscustomobject]@{ errorCode = 50126 } }
            } -ParameterFilter { $Uri -like 'v1.0/auditLogs/signIns*' }
        }

        It 'summarises sign-ins per user, IP address and location' {
            $rows = @(& (Get-ScriptPath 'export-sign-in-locations.ps1') -TenantId $script:Tenant -Days 7)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].SignInCount | Should -Be 2
            $rows[0].FailedCount | Should -Be 1
            $rows[0].City | Should -Be 'Gold Coast'
            $rows[0].Devices | Should -Be 'Windows 10 / Edge 128.0'
            $rows[0].FirstSeen | Should -Be '2026-10-01T01:00:00Z'
            $rows[0].LastSeen | Should -Be '2026-10-02T01:00:00Z'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'GET' -and $Uri -match 'createdDateTime ge ' }
        }

        It 'returns one row per sign-in with -Detailed' {
            $rows = @(& (Get-ScriptPath 'export-sign-in-locations.ps1') -TenantId $script:Tenant -Detailed)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
        }
    }

    Context 'find-mailboxes-below-plan-quota.ps1' {
        BeforeEach {
            Mock Get-MailboxPlan {
                [pscustomobject]@{ Name = 'ExchangeOnlineEnterprise-1'; Identity = 'ExchangeOnlineEnterprise-1'; Alias = 'ExchangeOnlineEnterprise-1'; ProhibitSendReceiveQuota = '100 GB (107,374,182,400 bytes)'; ProhibitSendQuota = '99 GB (106,300,440,576 bytes)'; IssueWarningQuota = '98 GB (105,226,698,752 bytes)' }
            }
            Mock Get-EXOMailbox {
                [pscustomobject]@{ PrimarySmtpAddress = 'small@contoso.com'; MailboxPlan = 'ExchangeOnlineEnterprise-1'; ProhibitSendReceiveQuota = '50 GB (53,687,091,200 bytes)' }
                [pscustomobject]@{ PrimarySmtpAddress = 'ok@contoso.com'; MailboxPlan = 'ExchangeOnlineEnterprise-1'; ProhibitSendReceiveQuota = '100 GB (107,374,182,400 bytes)' }
            } -ParameterFilter { -not $Identity }
            Mock Get-EXOMailbox { [pscustomobject]@{ ProhibitSendReceiveQuota = '100 GB (107,374,182,400 bytes)' } } -ParameterFilter { $Identity }
            Mock Set-Mailbox { }
        }

        It 'reports only mailboxes below their plan quota' {
            $rows = @(& (Get-ScriptPath 'find-mailboxes-below-plan-quota.ps1') -TenantId $script:Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].Mailbox | Should -Be 'small@contoso.com'
            $rows[0].CurrentQuotaGB | Should -Be 50
            $rows[0].PlanQuotaGB | Should -Be 100
            Should -Invoke Set-Mailbox -Times 0 -Exactly
        }

        It 'makes no change under -Apply -WhatIf' {
            $rows = @(& (Get-ScriptPath 'find-mailboxes-below-plan-quota.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Set-Mailbox -Times 0 -Exactly
        }

        It 'sets the plan quotas in bytes with -Apply' {
            $rows = @(& (Get-ScriptPath 'find-mailboxes-below-plan-quota.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            $rows[0].Action | Should -Be 'Raised'
            Should -Invoke Set-Mailbox -Times 1 -Exactly -ParameterFilter { $ProhibitSendReceiveQuota -eq '107374182400' -and $IssueWarningQuota -eq '105226698752' }
        }
    }

    Context 'Skipped customers and hygiene' {
        It 'skips customers without an active GDAP relationship' {
            Mock Get-MspCustomer { [pscustomobject]@{ TenantId = '22222222-2222-4222-8222-222222222222'; DisplayName = 'Contoso'; GdapStatus = 'expired' } }
            $rows = @(& (Get-ScriptPath 'disable-winmail-dat-tnef.ps1') -AllCustomers)
            $rows[0].Status | Should -Be 'Skipped'
            Should -Invoke Connect-MspExchangeOnline -Times 0 -Exactly
        }

        It 'contains no retired or unsafe connection methods outside help text' {
            $pattern = 'Connect-MsolService|Get-MsolPartnerContract|Connect-AzureAD|AzureRM|New-PSSession|-Authentication\s+Basic|ConvertTo-SecureString|powershell-liveid|graph\.windows\.net|oauth2/token\b|New-MsolUser|Add-MsolRoleMember'
            foreach ($file in $script:OwnScripts | ForEach-Object { Get-Item -LiteralPath (Get-ScriptPath $_) }) {
                $code = (Get-Content -LiteralPath $file.FullName -Raw) -replace '(?s)<#.*?#>', ''
                $code | Should -Not -Match $pattern -Because $file.Name
            }
        }

        It 'has the required header and help in every script' {
            foreach ($file in $script:OwnScripts | ForEach-Object { Get-Item -LiteralPath (Get-ScriptPath $_) }) {
                $text = Get-Content -LiteralPath $file.FullName -Raw
                $text | Should -Match '#Requires -Version 7\.4' -Because $file.Name
                $text | Should -Match '#Requires -Modules MspGdap' -Because $file.Name
                $help = Get-Help -Name $file.FullName -Full
                $help.Synopsis | Should -Not -BeNullOrEmpty -Because $file.Name
                @($help.Examples.Example).Count | Should -BeGreaterOrEqual 2 -Because $file.Name
                $text | Should -Match 'Replaces the original \d{4} method:' -Because $file.Name
                $text | Should -Match 'Required GDAP roles:' -Because $file.Name
                $text | Should -Match 'Required partner app permissions:' -Because $file.Name
                $text | Should -Match '\.LINK\s+https://gcit\.com\.au/' -Because $file.Name
                $text | Should -Not -Match '[\u2013\u2014]' -Because "$($file.Name) must not contain en or em dashes"
            }
        }
    }
}
