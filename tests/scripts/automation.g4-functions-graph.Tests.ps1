#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Mocked smoke tests for the scripts in scripts/automation that replace the retired Azure Functions v1,
# MSOnline and stored-password articles (see scripts/MAPPING.json). Nothing here contacts a
# tenant: MspGdap commands, Microsoft Graph and Exchange Online cmdlets are all mocked.

BeforeAll {
    $repoRoot = (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath '..')).ProviderPath
    Import-Module (Join-Path -Path $repoRoot -ChildPath 'src' -AdditionalChildPath 'MspGdap', 'MspGdap.psd1') -Force
    $script:ScriptDir = Join-Path -Path $repoRoot -ChildPath 'scripts' -AdditionalChildPath 'automation'
    $script:Tenant = 'contoso.onmicrosoft.com'
    $script:TenantGuid = '11111111-1111-1111-1111-111111111111'
    $global:G4AutoGraphCalls = [System.Collections.Generic.List[object]]::new()

    # Exchange Online and Security and Compliance cmdlets are imported at connect time, so stub them here.
    $stubNames = @(
        'Get-EXOMailbox', 'Get-EXORecipient', 'Get-AcceptedDomain', 'Get-HostedOutboundSpamFilterPolicy',
        'Set-HostedOutboundSpamFilterPolicy', 'Get-TransportRule', 'New-TransportRule', 'Set-TransportRule',
        'Set-Mailbox', 'Get-InboxRule', 'Get-ExternalInOutlook', 'Get-AntiPhishPolicy', 'Disconnect-ExchangeOnline',
        'Get-RetentionCompliancePolicy', 'New-RetentionCompliancePolicy', 'New-RetentionComplianceRule', 'Get-Secret',
        'Push-OutputBinding'
    )
    # Leave a real Get-Secret (SecretManagement installed) in place. Shadowing it with a global stub and removing
    # the stub later confuses Pester's mock call history for test files that run afterwards and mock Get-Secret.
    $stubNames = @($stubNames | Where-Object { $_ -ne 'Get-Secret' -or -not (Get-Command -Name 'Get-Secret' -CommandType Cmdlet -ErrorAction SilentlyContinue) })
    foreach ($name in $stubNames) {
        $null = New-Item -Path "Function:\global:$name" -Value { [CmdletBinding()] param([Parameter(ValueFromRemainingArguments)][object[]]$Rest) } -Force
    }
    $script:StubNames = $stubNames

    function global:Get-G4AutomationGraphFixture {
        param([string]$Method, [string]$Uri)
        if ($Method -ne 'GET') { return $null }
        switch -Regex ($Uri) {
            '^organization' {
                return [pscustomobject]@{
                    id              = '11111111-1111-1111-1111-111111111111'
                    displayName     = 'Contoso'
                    verifiedDomains = @(
                        [pscustomobject]@{ name = 'contoso.onmicrosoft.com'; isInitial = $true; isDefault = $false }
                        [pscustomobject]@{ name = 'contoso.com'; isInitial = $false; isDefault = $true }
                    )
                }
            }
            '^domains$' {
                return @(
                    [pscustomobject]@{ id = 'contoso.com'; isVerified = $true; passwordValidityPeriodInDays = 90 }
                    [pscustomobject]@{ id = 'contoso.onmicrosoft.com'; isVerified = $true; passwordValidityPeriodInDays = 2147483647 }
                    [pscustomobject]@{ id = 'unverified.contoso.com'; isVerified = $false; passwordValidityPeriodInDays = 90 }
                )
            }
            '^domains/' { return [pscustomobject]@{ id = 'contoso.com'; passwordValidityPeriodInDays = 2147483647 } }
            '^subscribedSkus' {
                return @(
                    [pscustomobject]@{ skuId = 'a'; skuPartNumber = 'SPE_E3'; capabilityStatus = 'Enabled'; consumedUnits = 7; prepaidUnits = [pscustomobject]@{ enabled = 10; warning = 0; suspended = 0 } }
                    [pscustomobject]@{ skuId = 'b'; skuPartNumber = 'FLOW_FREE'; capabilityStatus = 'Enabled'; consumedUnits = 2; prepaidUnits = [pscustomobject]@{ enabled = 10000; warning = 0; suspended = 0 } }
                    [pscustomobject]@{ skuId = 'c'; skuPartNumber = 'O365_BUSINESS_PREMIUM'; capabilityStatus = 'Enabled'; consumedUnits = 5; prepaidUnits = [pscustomobject]@{ enabled = 5; warning = 0; suspended = 0 } }
                )
            }
            '^auditLogs/directoryAudits' {
                return [pscustomobject]@{
                    activityDisplayName = 'Add member to role'
                    activityDateTime    = '2026-10-06T01:00:00Z'
                    result              = 'success'
                    initiatedBy         = [pscustomobject]@{ user = [pscustomobject]@{ userPrincipalName = 'admin@contoso.com' }; app = $null }
                    targetResources     = @(
                        [pscustomobject]@{
                            type               = 'User'
                            userPrincipalName  = 'new.admin@contoso.com'
                            displayName        = 'New Admin'
                            modifiedProperties = @([pscustomobject]@{ displayName = 'Role.DisplayName'; oldValue = $null; newValue = '"Global Administrator"' })
                        }
                        [pscustomobject]@{ type = 'Role'; displayName = $null; modifiedProperties = @() }
                    )
                }
            }
            '^roleManagement/directory/roleDefinitions' { return [pscustomobject]@{ id = 'r1'; displayName = 'Global Administrator' } }
            '^roleManagement/directory/roleAssignments' {
                return [pscustomobject]@{
                    roleDefinitionId = 'r1'
                    roleDefinition   = [pscustomobject]@{ displayName = 'Exchange Administrator' }
                    principal        = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; userPrincipalName = 'admin@contoso.com' }
                }
            }
            '^reports/authenticationMethods' { return [pscustomobject]@{ id = 'u1'; isMfaRegistered = $true; methodsRegistered = @('microsoftAuthenticatorPush'); isAdmin = $false } }
            '^users\?\$filter=startswith' { return [pscustomobject]@{ id = 'legacy1'; userPrincipalName = 'msp-reports@contoso.com'; accountEnabled = $true } }
            '^users\?' {
                return @(
                    [pscustomobject]@{ id = 'u1'; displayName = 'Jane Citizen'; userPrincipalName = 'jane@contoso.com'; accountEnabled = $true; assignedLicenses = @([pscustomobject]@{ skuId = 'a' }); proxyAddresses = @('SMTP:jane@contoso.com', 'smtp:j.citizen@contoso.com') }
                    [pscustomobject]@{ id = 'u2'; displayName = 'Unlicensed'; userPrincipalName = 'room@contoso.com'; accountEnabled = $true; assignedLicenses = @(); proxyAddresses = @() }
                )
            }
            '^users/' { return [pscustomobject]@{ accountEnabled = $false } }
            default { return $null }
        }
    }

    function Assert-G4CustomerRow {
        param([object[]]$Rows, [string[]]$Property)
        $Rows.Count | Should -BeGreaterThan 0
        foreach ($row in $Rows) {
            $row.CustomerTenantId | Should -Be '11111111-1111-1111-1111-111111111111'
            $row.CustomerName | Should -Be 'Contoso'
            $row.Error | Should -BeNullOrEmpty
            foreach ($name in $Property) { $row.PSObject.Properties.Name | Should -Contain $name }
        }
    }
}

AfterAll {
    foreach ($name in $script:StubNames) { Remove-Item -Path "Function:\global:$name" -ErrorAction SilentlyContinue }
    Remove-Item -Path 'Function:\global:Get-G4AutomationGraphFixture' -ErrorAction SilentlyContinue
    Remove-Variable -Name G4AutoGraphCalls, G4AutoPatterns, G4AutoSetMailbox, G4AutoBindings -Scope Global -ErrorAction SilentlyContinue
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'scripts/automation (g4 functions and Graph rewrites)' {
    BeforeEach {
        $global:G4AutoGraphCalls.Clear()
        Mock Invoke-MspGraphRequest {
            $verb = if ($Method) { $Method } else { 'GET' }
            $global:G4AutoGraphCalls.Add([pscustomobject]@{ Method = $verb; Uri = $Uri; TenantId = $TenantId })
            Get-G4AutomationGraphFixture -Method $verb -Uri $Uri
        }
        Mock Get-MspCustomer { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; DisplayName = 'Contoso'; DefaultDomainName = 'contoso.com' } }
        Mock Connect-MspExchangeOnline { [pscustomobject]@{ TenantId = $TenantId; Mode = 'Delegated' } }
        Mock Connect-MspSecurityCompliance { [pscustomobject]@{ TenantId = $TenantId } }
        Mock Disconnect-ExchangeOnline {}
        Mock Get-EXOMailbox {
            @(
                [pscustomobject]@{ UserPrincipalName = 'jane@contoso.com'; DisplayName = 'Jane Citizen (Sales)'; ForwardingSmtpAddress = 'smtp:jane@gmail.example'; ForwardingAddress = $null; DeliverToMailboxAndForward = $true }
                [pscustomobject]@{ UserPrincipalName = 'joe@contoso.com'; DisplayName = 'Joe Bloggs'; ForwardingSmtpAddress = 'smtp:joe@contoso.com'; ForwardingAddress = $null; DeliverToMailboxAndForward = $false }
            )
        }
        Mock Get-EXORecipient {}
        Mock Get-AcceptedDomain { @([pscustomobject]@{ DomainName = 'contoso.com' }, [pscustomobject]@{ DomainName = 'contoso.onmicrosoft.com' }) }
        Mock Get-HostedOutboundSpamFilterPolicy { [pscustomobject]@{ Name = 'Default'; IsDefault = $true; AutoForwardingMode = 'On' } }
        Mock Set-HostedOutboundSpamFilterPolicy {}
        Mock Get-TransportRule {}
        Mock New-TransportRule {}
        Mock Set-TransportRule {}
        Mock Set-Mailbox {}
        Mock Get-InboxRule { [pscustomobject]@{ Name = 'Fwd'; Enabled = $true; ForwardTo = @('"Ext" [SMTP:someone@fabrikam.example]'); ForwardAsAttachmentTo = $null; RedirectTo = $null } }
        Mock Get-ExternalInOutlook { [pscustomobject]@{ Enabled = $true } }
        Mock Get-AntiPhishPolicy { [pscustomobject]@{ IsDefault = $true; EnableFirstContactSafetyTips = $true; EnableMailboxIntelligenceProtection = $false; EnableTargetedUserProtection = $false } }
        Mock Get-RetentionCompliancePolicy {}
        Mock New-RetentionCompliancePolicy {}
        Mock New-RetentionComplianceRule {}
        Mock Test-MspExchangeAppAccess { [pscustomobject]@{ Success = $false; Outcome = 'Failed'; Steps = @() } }
        Mock Enable-MspExchangeAppAccess { [pscustomobject]@{ Success = $true; Outcome = 'Succeeded'; Steps = @() } }
    }

    Context 'set-password-never-expires.ps1' {
        It 'reports verified domains and writes nothing under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'set-password-never-expires.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'Domain', 'PasswordValidityDays', 'PasswordsExpire', 'Action'
            $rows.Count | Should -Be 2
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'WhatIf'
            ($rows | Where-Object Domain -eq 'contoso.onmicrosoft.com').Action | Should -Be 'AlreadyNeverExpires'
            @($global:G4AutoGraphCalls | Where-Object Method -ne 'GET').Count | Should -Be 0
        }

        It 'is report-only without -Apply' {
            $rows = @(& (Join-Path $script:ScriptDir 'set-password-never-expires.ps1') -AllCustomers)
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'WouldSetNeverExpire'
            @($global:G4AutoGraphCalls | Where-Object Method -ne 'GET').Count | Should -Be 0
        }

        It 'patches the domain with -Apply and confirms by readback' {
            $rows = @(& (Join-Path $script:ScriptDir 'set-password-never-expires.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'SetNeverExpire'
            @($global:G4AutoGraphCalls | Where-Object { $_.Method -eq 'PATCH' -and $_.Uri -eq 'domains/contoso.com' }).Count | Should -Be 1
        }

        It 'records a failed customer and carries on' {
            Mock Invoke-MspGraphRequest -ParameterFilter { $TenantId -eq 'broken.onmicrosoft.com' } -MockWith { throw 'AADSTS65001: consent missing' }
            $rows = @(& (Join-Path $script:ScriptDir 'set-password-never-expires.ps1') -TenantId 'broken.onmicrosoft.com', $script:Tenant -WarningAction SilentlyContinue)
            ($rows | Where-Object Action -eq 'Failed').Error | Should -Match 'AADSTS65001'
            @($rows | Where-Object CustomerName -eq 'Contoso').Count | Should -Be 2
        }

        It 'records a failed domain and carries on with the other domains' {
            Mock Invoke-MspGraphRequest -ParameterFilter { $Method -eq 'PATCH' } -MockWith { throw 'Federated domains cannot be updated' }
            $rows = @(& (Join-Path $script:ScriptDir 'set-password-never-expires.ps1') -TenantId $script:Tenant -Apply -Confirm:$false -WarningAction SilentlyContinue)
            $rows.Count | Should -Be 2
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'Failed'
            ($rows | Where-Object Domain -eq 'contoso.com').Error | Should -Match 'Federated'
            ($rows | Where-Object Domain -eq 'contoso.onmicrosoft.com').Action | Should -Be 'AlreadyNeverExpires'
        }
    }

    Context 'get-unused-licences.ps1' {
        It 'returns only paid SKUs with unused units' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-unused-licences.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'SkuPartNumber', 'Enabled', 'Consumed', 'Unused'
            $rows.Count | Should -Be 1
            $rows[0].SkuPartNumber | Should -Be 'SPE_E3'
            $rows[0].Unused | Should -Be 3
            $rows[0].Message | Should -Be 'Contoso has 3 unused SPE_E3 licence(s)'
        }

        It 'accepts Get-MspCustomer output on the pipeline' {
            $rows = @(Get-MspCustomer | & (Join-Path $script:ScriptDir 'get-unused-licences.ps1') -MinimumUnused 0)
            $rows.Count | Should -Be 2
        }
    }

    Context 'set-default-retention-policy.ps1' {
        It 'reports a customer with no policy and creates nothing under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'set-default-retention-policy.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'ExistingPolicyCount', 'PolicyName', 'Action'
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke New-RetentionCompliancePolicy -Times 0 -Exactly
            Should -Invoke New-RetentionComplianceRule -Times 0 -Exactly
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }

        It 'adds SharePoint, OneDrive and group locations only when the tenant has SharePoint Online' {
            $rows = @(& (Join-Path $script:ScriptDir 'set-default-retention-policy.ps1') -TenantId $script:Tenant)
            $rows[0].Action | Should -Be 'WouldCreate'
            $rows[0].Locations | Should -Be 'Exchange, PublicFolder'

            Mock Invoke-MspGraphRequest -ParameterFilter { $Uri -eq 'subscribedSkus' } -MockWith {
                [pscustomobject]@{ capabilityStatus = 'Enabled'; servicePlans = @([pscustomobject]@{ servicePlanName = 'SHAREPOINTSTANDARD'; provisioningStatus = 'Success' }) }
            }
            $rows = @(& (Join-Path $script:ScriptDir 'set-default-retention-policy.ps1') -TenantId $script:Tenant)
            $rows[0].Locations | Should -Be 'Exchange, SharePoint, OneDrive, ModernGroup, PublicFolder'
        }

        It 'skips customers that already have a policy' {
            Mock Get-RetentionCompliancePolicy { [pscustomobject]@{ Name = 'Customer policy' } }
            $rows = @(& (Join-Path $script:ScriptDir 'set-default-retention-policy.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            $rows[0].Action | Should -Be 'SkippedOtherPoliciesExist'
            Should -Invoke New-RetentionCompliancePolicy -Times 0 -Exactly
        }
    }

    Context 'get-admin-role-changes.ps1' {
        It 'returns role additions with the role name and current assignments' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-admin-role-changes.ps1') -TenantId $script:Tenant -IncludeCurrentAssignments)
            Assert-G4CustomerRow -Rows $rows -Property 'Activity', 'RoleName', 'Target', 'InitiatedBy'
            $change = $rows | Where-Object Activity -eq 'Add member to role'
            $change.RoleName | Should -Be 'Global Administrator'
            $change.Target | Should -Be 'new.admin@contoso.com'
            $change.InitiatedBy | Should -Be 'admin@contoso.com'
            ($rows | Where-Object Activity -eq 'CurrentAssignment').RoleName | Should -Be 'Global Administrator'
            ($global:G4AutoGraphCalls | Where-Object Uri -like 'auditLogs/*').Uri | Should -Match "category eq 'RoleManagement' and activityDateTime ge "
        }

        It 'reads the role name of a removal from oldValue' {
            Mock Invoke-MspGraphRequest -ParameterFilter { $Uri -like 'auditLogs/*' } -MockWith {
                [pscustomobject]@{
                    activityDisplayName = 'Remove member from role'
                    activityDateTime    = '2026-10-06T02:00:00Z'
                    result              = 'success'
                    initiatedBy         = [pscustomobject]@{ user = [pscustomobject]@{ userPrincipalName = 'admin@contoso.com' }; app = $null }
                    targetResources     = @(
                        [pscustomobject]@{
                            type               = 'User'
                            userPrincipalName  = 'old.admin@contoso.com'
                            modifiedProperties = @([pscustomobject]@{ displayName = 'Role.DisplayName'; oldValue = '"Exchange Administrator"'; newValue = '""' })
                        }
                    )
                }
            }
            $rows = @(& (Join-Path $script:ScriptDir 'get-admin-role-changes.ps1') -TenantId $script:Tenant)
            $rows.Count | Should -Be 1
            $rows[0].Activity | Should -Be 'Remove member from role'
            $rows[0].RoleName | Should -Be 'Exchange Administrator'
        }
    }

    Context 'get-user-details.ps1' {
        It 'returns flat user rows with licence names' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-user-details.ps1') -TenantId $script:Tenant -LicensedOnly -IncludeAuthMethods)
            Assert-G4CustomerRow -Rows $rows -Property 'UserPrincipalName', 'Licences', 'IsMfaRegistered', 'Aliases'
            $rows.Count | Should -Be 1
            $rows[0].Licences | Should -Be 'SPE_E3'
            $rows[0].Aliases | Should -Be 'j.citizen@contoso.com'
            $rows[0].IsMfaRegistered | Should -BeTrue
        }

        It 'adds the mailbox type from Exchange Online with -IncludeMailboxType' {
            Mock Get-EXOMailbox { [pscustomobject]@{ ExternalDirectoryObjectId = 'u1'; RecipientTypeDetails = 'UserMailbox' } }
            $rows = @(& (Join-Path $script:ScriptDir 'get-user-details.ps1') -TenantId $script:Tenant -IncludeMailboxType)
            ($rows | Where-Object UserId -eq 'u1').MailboxType | Should -Be 'UserMailbox'
            ($rows | Where-Object UserId -eq 'u2').MailboxType | Should -BeNullOrEmpty
            Should -Invoke Connect-MspExchangeOnline -Times 1 -Exactly
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }
    }

    Context 'enable-exchange-app-access.ps1' {
        It 'reports app access and legacy accounts and changes nothing under -WhatIf' {
            $params = @{ TenantId = $script:Tenant; AppId = '22222222-2222-2222-2222-222222222222'; LegacyAccountUpnPrefix = 'msp-reports'; DisableLegacyAccounts = $true; Apply = $true; WhatIf = $true }
            $rows = @(& (Join-Path $script:ScriptDir 'enable-exchange-app-access.ps1') @params)
            Assert-G4CustomerRow -Rows $rows -Property 'AppAccessReady', 'Action', 'LegacyAccounts', 'LegacyAccountsAction'
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].LegacyAccounts | Should -Match 'msp-reports@contoso.com \(enabled, roles: Exchange Administrator\)'
            Should -Invoke Enable-MspExchangeAppAccess -Times 0 -Exactly
            @($global:G4AutoGraphCalls | Where-Object Method -ne 'GET').Count | Should -Be 0
        }

        It 'enables access and blocks legacy accounts with -Apply' {
            $params = @{ TenantId = $script:Tenant; AppId = '22222222-2222-2222-2222-222222222222'; LegacyAccountUpnPrefix = 'msp-reports'; DisableLegacyAccounts = $true; Apply = $true; Confirm = $false }
            $rows = @(& (Join-Path $script:ScriptDir 'enable-exchange-app-access.ps1') @params)
            $rows[0].Action | Should -Be 'Enabled'
            $rows[0].LegacyAccountsAction | Should -Be 'Disabled'
            Should -Invoke Enable-MspExchangeAppAccess -Times 1 -Exactly
        }
    }

    Context 'set-display-name-spoof-rule.ps1' {
        It 'escapes display names, reports native controls and writes no rule under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'set-display-name-spoof-rule.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'ExternalTagEnabled', 'FirstContactSafetyTips', 'RuleState', 'Action'
            $rows[0].RuleState | Should -Be 'Missing'
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[0].ExternalTagEnabled | Should -BeTrue
            Should -Invoke New-TransportRule -Times 0 -Exactly
            Should -Invoke Set-TransportRule -Times 0 -Exactly
        }

        It 'creates the rule with escaped patterns when -Apply is used' {
            $global:G4AutoPatterns = $null
            Mock New-TransportRule { $global:G4AutoPatterns = $Rest }
            $rows = @(& (Join-Path $script:ScriptDir 'set-display-name-spoof-rule.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            $rows[0].Action | Should -Be 'Created'
            ($global:G4AutoPatterns | Out-String) | Should -Match ([regex]::Escape('Jane Citizen \(Sales\)'))
        }

        It 'refuses app-only without a certificate' {
            { & (Join-Path $script:ScriptDir 'set-display-name-spoof-rule.ps1') -TenantId $script:Tenant -ExchangeAppId '22222222-2222-2222-2222-222222222222' } | Should -Throw '*ExchangeCertificate*'
        }
    }

    Context 'get-external-forwarding.ps1' {
        It 'returns external forwards only and removes nothing under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'get-external-forwarding.ps1') -TenantId $script:Tenant -IncludeInboxRules -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'Mailbox', 'Source', 'ExternalRecipient', 'OutboundAutoForwardingMode', 'Action'
            $mailboxRow = $rows | Where-Object Source -eq 'MailboxForwarding'
            @($mailboxRow).Count | Should -Be 1
            $mailboxRow.ExternalRecipient | Should -Be 'jane@gmail.example'
            $mailboxRow.Action | Should -Be 'WhatIf'
            ($rows | Where-Object Source -eq 'InboxRule')[0].ExternalRecipient | Should -Be 'someone@fabrikam.example'
            Should -Invoke Set-Mailbox -Times 0 -Exactly
        }

        It 'clears only the external forwarding setting with -Apply' {
            $global:G4AutoSetMailbox = $null
            Mock Set-Mailbox { $global:G4AutoSetMailbox = ($Rest | Out-String) }
            $rows = @(& (Join-Path $script:ScriptDir 'get-external-forwarding.ps1') -TenantId $script:Tenant -Apply -Confirm:$false)
            ($rows | Where-Object Source -eq 'MailboxForwarding').Action | Should -Be 'Removed'
            Should -Invoke Set-Mailbox -Times 1 -Exactly
            $global:G4AutoSetMailbox | Should -Match 'ForwardingSmtpAddress'
            $global:G4AutoSetMailbox | Should -Not -Match '-ForwardingAddress:'
            $global:G4AutoSetMailbox | Should -Match 'DeliverToMailboxAndForward'
        }

        It 'queues each forward once from the ExternalForwardingReport function' {
            $app = Join-Path $TestDrive 'app'
            $null = New-Item -ItemType Directory -Path (Join-Path $app 'scripts'), (Join-Path $app 'ExternalForwardingReport') -Force
            Copy-Item -Path (Join-Path $script:ScriptDir 'get-external-forwarding.ps1') -Destination (Join-Path $app 'scripts')
            Copy-Item -Path (Join-Path $script:ScriptDir 'functions' -AdditionalChildPath 'ExternalForwardingReport', 'run.ps1') -Destination (Join-Path $app 'ExternalForwardingReport')
            $global:G4AutoBindings = @{}
            Mock Push-OutputBinding { $global:G4AutoBindings[$Rest[1]] = $Rest[3] }
            $timer = [pscustomobject]@{ IsPastDue = $false }
            $env:MSPGDAP_TENANT_IDS = $script:Tenant
            try {
                & (Join-Path $app 'ExternalForwardingReport' -AdditionalChildPath 'run.ps1') -Timer $timer -previousState $null
                @($global:G4AutoBindings['alerts']).Count | Should -Be 1
                $state = $global:G4AutoBindings['state']
                $state | Should -Match 'jane@contoso.com'

                $global:G4AutoBindings = @{}
                & (Join-Path $app 'ExternalForwardingReport' -AdditionalChildPath 'run.ps1') -Timer $timer -previousState $state
                $global:G4AutoBindings.ContainsKey('alerts') | Should -BeFalse
                $global:G4AutoBindings['state'] | Should -Be $state
            }
            finally {
                Remove-Item -Path Env:\MSPGDAP_TENANT_IDS -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'block-external-forwarding.ps1' {
        It 'reports the outbound policy and changes nothing under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'block-external-forwarding.ps1') -TenantId $script:Tenant -Apply -WhatIf)
            Assert-G4CustomerRow -Rows $rows -Property 'DefaultAutoForwardingMode', 'PoliciesAllowingForwards', 'Action'
            $rows[0].DefaultAutoForwardingMode | Should -Be 'On'
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Set-HostedOutboundSpamFilterPolicy -Times 0 -Exactly
        }

        It 'does not create the transport rule under -WhatIf' {
            $rows = @(& (Join-Path $script:ScriptDir 'block-external-forwarding.ps1') -TenantId $script:Tenant -Method TransportRule -Apply -WhatIf)
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke New-TransportRule -Times 0 -Exactly
        }
    }

    Context 'sync-tenant-info-itglue.ps1' {
        It 'collects tenant info without IT Glue' {
            $rows = @(& (Join-Path $script:ScriptDir 'sync-tenant-info-itglue.ps1') -TenantId $script:Tenant)
            Assert-G4CustomerRow -Rows $rows -Property 'InitialDomain', 'VerifiedDomains', 'LicenceCount', 'LicensedUserCount', 'Action'
            $rows[0].InitialDomain | Should -Be 'contoso.onmicrosoft.com'
            $rows[0].LicensedUserCount | Should -Be 1
            $rows[0].Action | Should -Be 'Collected'
        }

        It 'reads IT Glue but does not write under -WhatIf' {
            $mapPath = Join-Path $TestDrive 'map.csv'
            "TenantId,OrganizationId`n11111111-1111-1111-1111-111111111111,42" | Set-Content -Path $mapPath
            Mock Get-Secret { 'not-a-real-key' }
            Mock Invoke-RestMethod { [pscustomobject]@{ data = @(); links = $null } }
            $rows = @(& (Join-Path $script:ScriptDir 'sync-tenant-info-itglue.ps1') -TenantId $script:Tenant -ITGlueFlexibleAssetTypeId 7 -ITGlueOrganizationMapPath $mapPath -Apply -WhatIf)
            $rows[0].ITGlueOrganizationId | Should -Be '42'
            $rows[0].Action | Should -Be 'WhatIf'
            Should -Invoke Invoke-RestMethod -ParameterFilter { $Method -ne 'GET' } -Times 0 -Exactly
            Should -Invoke Invoke-RestMethod -ParameterFilter { "$Uri" -like 'https://api.itglue.com/flexible_assets*' } -Times 1 -Exactly
        }
    }
}
