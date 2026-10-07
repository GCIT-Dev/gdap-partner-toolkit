#Requires -Version 7.4
# Offline smoke tests for the licensing, admin and user-report scripts in scripts/delegated-reporting
# (see scripts/MAPPING.json). Every MspGdap call that would reach a tenant, the Have I Been Pwned API and
# SecretManagement are mocked. Nothing here contacts Microsoft, Have I Been Pwned or any customer tenant.

BeforeAll {
    $script:RepoRoot = (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../..')).Path
    $script:ScriptRoot = Join-Path -Path $script:RepoRoot -ChildPath 'scripts/delegated-reporting'
    $srcPath = Join-Path -Path $script:RepoRoot -ChildPath 'src'
    if (($env:PSModulePath -split [System.IO.Path]::PathSeparator) -notcontains $srcPath) {
        $env:PSModulePath = $srcPath + [System.IO.Path]::PathSeparator + $env:PSModulePath
    }
    Import-Module (Join-Path -Path $srcPath -ChildPath 'MspGdap/MspGdap.psd1') -Force
    Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction SilentlyContinue

    # Mock bodies run in the scope of the script under test, so shared test data lives in one
    # global hashtable that AfterAll removes.
    $global:G2Reports = @{
        CustomerId = '33333333-3333-4333-8333-333333333333'
        Tenant     = 'contoso.onmicrosoft.com'
        SkuSpb     = 'aaaaaaaa-0000-4000-8000-000000000001'
        SkuFree    = 'aaaaaaaa-0000-4000-8000-000000000009'
        GroupId    = 'bbbbbbbb-0000-4000-8000-000000000001'
        AdminOne   = 'cccccccc-0000-4000-8000-000000000001'
        AdminTwo   = 'cccccccc-0000-4000-8000-000000000002'
    }

    function Get-ScriptPath {
        param([Parameter(Mandatory)][string]$Name)
        Join-Path -Path $script:ScriptRoot -ChildPath $Name
    }

    # check-hibp-breaches.ps1 has #Requires -Modules Microsoft.PowerShell.SecretManagement. On a runner without
    # that module (CI installs only Pester and PSScriptAnalyzer), load a stub module of the same name so the
    # script can start. Get-Secret is mocked in every test that runs it.
    $script:G2StubSecretModule = $false
    if (-not (Get-Module -Name 'Microsoft.PowerShell.SecretManagement') -and -not (Get-Module -ListAvailable -Name 'Microsoft.PowerShell.SecretManagement')) {
        $stub = New-Module -Name 'Microsoft.PowerShell.SecretManagement' -ScriptBlock {
            function Get-Secret { [CmdletBinding()] param($Name, $Vault, [switch]$AsPlainText) }
        }
        $stub | Import-Module -Global
        $script:G2StubSecretModule = $true
    }
    if (-not (Get-Command -Name Get-Secret -ErrorAction SilentlyContinue)) {
        function Get-Secret { [CmdletBinding()] param($Name, $Vault, [switch]$AsPlainText) }
    }

    function Invoke-FakeGraph {
        param([string]$Uri, [string]$Method)
        $data = $global:G2Reports
        if ($Method -and $Method -ne 'GET') {
            if ($Method -eq 'POST') { return [pscustomobject]@{ id = 'dddddddd-0000-4000-8000-000000000001' } }
            return $null
        }
        switch -Regex ($Uri) {
            '^v1\.0/organization' { return [pscustomobject]@{ displayName = 'Contoso' } }
            '^v1\.0/subscribedSkus' {
                return @(
                    [pscustomobject]@{ skuId = $data.SkuSpb; skuPartNumber = 'SPB'; capabilityStatus = 'Enabled'; appliesTo = 'User'; consumedUnits = 3; prepaidUnits = [pscustomobject]@{ enabled = 5; warning = 0; suspended = 0 } }
                    [pscustomobject]@{ skuId = $data.SkuFree; skuPartNumber = 'FLOW_FREE'; capabilityStatus = 'Enabled'; appliesTo = 'User'; consumedUnits = 1; prepaidUnits = [pscustomobject]@{ enabled = 10000; warning = 0; suspended = 0 } }
                )
            }
            '^v1\.0/roleManagement/directory/roleDefinitions' {
                if ($Uri -match "displayName eq 'Exchange Administrator'") { return [pscustomobject]@{ id = '29232cdf-9323-42fd-ade2-1d097af3e4de'; displayName = 'Exchange Administrator' } }
                return @()
            }
            '^v1\.0/roleManagement/directory/roleAssignments' {
                return @(
                    [pscustomobject]@{ principalId = $data.AdminOne; principal = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = $data.AdminOne; displayName = 'Default Admin' } }
                    [pscustomobject]@{ principalId = $data.AdminTwo; principal = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.user'; id = $data.AdminTwo; displayName = 'Named Admin' } }
                    [pscustomobject]@{ principalId = 'eeeeeeee-0000-4000-8000-000000000001'; principal = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.group'; id = 'eeeeeeee-0000-4000-8000-000000000001'; displayName = 'Role group' } }
                )
            }
            '^v1\.0/users/cccccccc-0000-4000-8000-000000000001' {
                return [pscustomobject]@{ id = $data.AdminOne; displayName = 'Default Admin'; userPrincipalName = 'admin@contoso.onmicrosoft.com'; accountEnabled = $true; assignedLicenses = @(); signInActivity = [pscustomobject]@{ lastSignInDateTime = '2025-01-01T00:00:00Z' } }
            }
            '^v1\.0/users/cccccccc-0000-4000-8000-000000000002' {
                return [pscustomobject]@{ id = $data.AdminTwo; displayName = 'Named Admin'; userPrincipalName = 'it.admin@contoso.com'; accountEnabled = $true; assignedLicenses = @([pscustomobject]@{ skuId = $data.SkuSpb }) }
            }
            '^v1\.0/users[?]' {
                return @(
                    [pscustomobject]@{
                        id                         = 'cccccccc-0000-4000-8000-000000000003'
                        displayName                = 'Adele Vance'
                        givenName                  = 'Adele'
                        surname                    = 'Vance'
                        mail                       = 'adele.vance@contoso.com'
                        userPrincipalName          = 'adele.vance@contoso.com'
                        userType                   = 'Member'
                        accountEnabled             = $true
                        usageLocation              = 'AU'
                        proxyAddresses             = @('SMTP:adele.vance@contoso.com', 'smtp:adele@contoso.onmicrosoft.com')
                        lastPasswordChangeDateTime = '2026-01-01T00:00:00Z'
                        assignedLicenses           = @([pscustomobject]@{ skuId = $data.SkuSpb })
                        licenseAssignmentStates    = @([pscustomobject]@{ skuId = $data.SkuSpb; assignedByGroup = $data.GroupId })
                    }
                    [pscustomobject]@{
                        id                         = 'cccccccc-0000-4000-8000-000000000004'
                        displayName                = 'Guest User'
                        givenName                  = 'Guest'
                        surname                    = 'User'
                        mail                       = 'guest@fabrikam.com'
                        userPrincipalName          = 'guest_fabrikam.com#EXT#@contoso.onmicrosoft.com'
                        userType                   = 'Guest'
                        accountEnabled             = $true
                        usageLocation              = $null
                        proxyAddresses             = @('SMTP:guest@fabrikam.com')
                        lastPasswordChangeDateTime = $null
                        assignedLicenses           = @([pscustomobject]@{ skuId = $data.SkuSpb })
                        licenseAssignmentStates    = @()
                    }
                    [pscustomobject]@{
                        id                         = 'cccccccc-0000-4000-8000-000000000005'
                        displayName                = 'Room 1'
                        givenName                  = $null
                        surname                    = $null
                        mail                       = 'room1@contoso.com'
                        userPrincipalName          = 'room1@contoso.com'
                        userType                   = 'Member'
                        accountEnabled             = $false
                        usageLocation              = $null
                        proxyAddresses             = @()
                        lastPasswordChangeDateTime = $null
                        assignedLicenses           = @()
                        licenseAssignmentStates    = @()
                    }
                )
            }
            '^v1\.0/policies/identitySecurityDefaultsEnforcementPolicy' { return [pscustomobject]@{ isEnabled = $false } }
            '^v1\.0/identity/conditionalAccess/policies' { return @() }
            '^v1\.0/reports/authenticationMethods/userRegistrationDetails' {
                return @(
                    [pscustomobject]@{ userPrincipalName = 'admin@contoso.onmicrosoft.com'; userDisplayName = 'Default Admin'; isAdmin = $true; isMfaRegistered = $false; isMfaCapable = $false; methodsRegistered = @() }
                    [pscustomobject]@{ userPrincipalName = 'adele.vance@contoso.com'; userDisplayName = 'Adele Vance'; isAdmin = $false; isMfaRegistered = $true; isMfaCapable = $true; methodsRegistered = @('microsoftAuthenticatorPush') }
                )
            }
            default { throw "Unexpected Graph request in test: $Method $Uri" }
        }
    }

    function Assert-CustomerShape {
        param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)
        $Rows.Count | Should -BeGreaterThan 0
        foreach ($row in $Rows) {
            $row | Should -BeOfType [pscustomobject]
            $row.PSObject.Properties.Name | Should -Contain 'CustomerTenantId'
            $row.PSObject.Properties.Name | Should -Contain 'CustomerName'
            $row.PSObject.Properties.Name | Should -Contain 'Status'
            $row.CustomerTenantId | Should -Be $global:G2Reports.CustomerId
        }
    }

    function Assert-NoGraphWrite {
        Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -and $Method -ne 'GET' }
    }
}

AfterAll {
    Remove-Variable -Name G2Reports -Scope Global -ErrorAction SilentlyContinue
    if ($script:G2StubSecretModule) { Remove-Module -Name 'Microsoft.PowerShell.SecretManagement' -Force -ErrorAction SilentlyContinue }
}

Describe 'delegated-reporting scripts (users and licensing group)' {
    BeforeEach {
        Mock Resolve-MspTenantId { $global:G2Reports.CustomerId }
        Mock Invoke-MspGraphRequest { Invoke-FakeGraph -Uri $Uri -Method $Method }
        Mock Get-MspAuthHeader { throw 'Get-MspAuthHeader must not be called in tests.' }
        Mock Get-MspCustomer {
            [pscustomobject]@{ TenantId = $global:G2Reports.CustomerId; DisplayName = 'Contoso'; DefaultDomainName = 'contoso.com'; GdapStatus = 'active' }
        }
    }

    Context 'export-unused-licences.ps1' {
        It 'calculates unused units' {
            $rows = @(& (Get-ScriptPath 'export-unused-licences.ps1') -TenantId $global:G2Reports.Tenant)
            Assert-CustomerShape -Rows $rows
            ($rows | Where-Object SkuPartNumber -eq 'SPB').UnusedUnits | Should -Be 2
            Assert-NoGraphWrite
        }

        It 'filters with -OnlyUnused and -ExcludeSkuPartNumber' {
            $rows = @(& (Get-ScriptPath 'export-unused-licences.ps1') -AllCustomers -OnlyUnused -ExcludeSkuPartNumber 'FLOW_FREE')
            $rows.Count | Should -Be 1
            $rows[0].SkuPartNumber | Should -Be 'SPB'
            $rows[0].DefaultDomain | Should -Be 'contoso.com'
        }
    }

    Context 'export-user-licences.ps1' {
        It 'returns licensed users with group assignments shown' {
            $csv = Join-Path -Path $TestDrive -ChildPath 'licences.csv'
            $rows = @(& (Get-ScriptPath 'export-user-licences.ps1') -TenantId $global:G2Reports.Tenant -OutputPath $csv)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            ($rows | Where-Object UserPrincipalName -eq 'adele.vance@contoso.com').GroupAssignedLicences | Should -Be 'SPB'
            Test-Path -Path $csv | Should -BeTrue
            Assert-NoGraphWrite
        }
    }

    Context 'export-customers-by-licensed-user-count.ps1' {
        It 'leaves out customers below the threshold' {
            $rows = @(& (Get-ScriptPath 'export-customers-by-licensed-user-count.ps1') -TenantId $global:G2Reports.Tenant)
            $rows.Count | Should -Be 0
        }

        It 'marks the threshold with -IncludeAll' {
            $rows = @(& (Get-ScriptPath 'export-customers-by-licensed-user-count.ps1') -TenantId $global:G2Reports.Tenant -MinimumLicensedUsers 2 -IncludeAll)
            Assert-CustomerShape -Rows $rows
            $rows[0].LicensedUserCount | Should -Be 2
            $rows[0].MeetsThreshold | Should -BeTrue
            $rows[0].Licences | Should -Match 'SPB \(3\)'
            Assert-NoGraphWrite
        }
    }

    Context 'export-global-admins.ps1' {
        It 'lists Global Administrators, including groups' {
            $rows = @(& (Get-ScriptPath 'export-global-admins.ps1') -TenantId $global:G2Reports.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 3
            ($rows | Where-Object UserId -eq $global:G2Reports.AdminOne).LastSignInDateTime | Should -Be '2025-01-01T00:00:00Z'
            ($rows | Where-Object PrincipalType -eq 'group').DisplayName | Should -Be 'Role group'
            Assert-NoGraphWrite
        }

        It 'reports another role with -RoleName' {
            $rows = @(& (Get-ScriptPath 'export-global-admins.ps1') -TenantId $global:G2Reports.Tenant -RoleName 'Exchange Administrator')
            Assert-CustomerShape -Rows $rows
            $rows[0].RoleName | Should -Be 'Exchange Administrator'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Uri -match "roleDefinitionId eq '29232cdf-9323-42fd-ade2-1d097af3e4de'" }
            Assert-NoGraphWrite
        }

        It 'records a role that does not exist as a failed customer' {
            $rows = @(& (Get-ScriptPath 'export-global-admins.ps1') -TenantId $global:G2Reports.Tenant -RoleName 'No Such Role' -WarningAction SilentlyContinue)
            $rows[0].Status | Should -Be 'Failed'
            $rows[0].Error | Should -Match 'No role named'
        }

        It 'returns only unlicensed users with -UnlicensedOnly' {
            $rows = @(& (Get-ScriptPath 'export-global-admins.ps1') -AllCustomers -UnlicensedOnly)
            $rows.Count | Should -Be 1
            $rows[0].UserPrincipalName | Should -Be 'admin@contoso.onmicrosoft.com'
        }

        It 'blocks nothing under -WhatIf and never blocks the last enabled admin' {
            $csv = Join-Path -Path $TestDrive -ChildPath 'reviewed.csv'
            @(
                [pscustomobject]@{ CustomerTenantId = $global:G2Reports.CustomerId; UserId = $global:G2Reports.AdminOne }
                [pscustomobject]@{ CustomerTenantId = $global:G2Reports.CustomerId; UserId = $global:G2Reports.AdminTwo }
            ) | Export-Csv -Path $csv -NoTypeInformation
            $rows = @(& (Get-ScriptPath 'export-global-admins.ps1') -BlockFromCsv $csv -Apply -WhatIf)
            Assert-CustomerShape -Rows $rows
            $rows[0].Action | Should -Be 'WhatIf'
            $rows[1].Action | Should -Be 'SkippedLastEnabledAdmin'
            Assert-NoGraphWrite
        }

        It 'blocks a reviewed account when applied' {
            $csv = Join-Path -Path $TestDrive -ChildPath 'reviewed-one.csv'
            [pscustomobject]@{ CustomerTenantId = $global:G2Reports.CustomerId; UserId = $global:G2Reports.AdminOne } | Export-Csv -Path $csv -NoTypeInformation
            $null = & (Get-ScriptPath 'export-global-admins.ps1') -BlockFromCsv $csv -Apply -Confirm:$false -WarningAction SilentlyContinue
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'PATCH' -and $Body.accountEnabled -eq $false -and $Uri -eq "v1.0/users/$($global:G2Reports.AdminOne)" }
        }
    }

    Context 'export-admin-mfa-status.ps1' {
        It 'reports admins and that nothing enforces MFA' {
            $rows = @(& (Get-ScriptPath 'export-admin-mfa-status.ps1') -TenantId $global:G2Reports.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].IsMfaRegistered | Should -BeFalse
            $rows[0].AdminMfaEnforcedBy | Should -Be 'Nothing'
            $rows[0].Status | Should -Be 'Attention'
            Assert-NoGraphWrite
        }

        It 'still reports when Conditional Access cannot be read, and creates no policy there' {
            Mock Invoke-MspGraphRequest -ParameterFilter { $Uri -like 'v1.0/identity/conditionalAccess/policies*' -and (-not $Method -or $Method -eq 'GET') } { throw 'Forbidden: tenant is not licensed for Conditional Access.' }
            $rows = @(& (Get-ScriptPath 'export-admin-mfa-status.ps1') -TenantId $global:G2Reports.Tenant -CreateConditionalAccessPolicy -Apply -Confirm:$false)
            Assert-CustomerShape -Rows $rows
            $rows[0].AdminMfaEnforcedBy | Should -Be 'Unknown'
            $rows[0].PolicyAction | Should -Be 'SkippedConditionalAccessNotReadable'
            $rows[0].Error | Should -Match 'Conditional Access not readable'
            Assert-NoGraphWrite
        }

        It 'creates no policy under -WhatIf' {
            $rows = @(& (Get-ScriptPath 'export-admin-mfa-status.ps1') -TenantId $global:G2Reports.Tenant -CreateConditionalAccessPolicy -Apply -WhatIf)
            $rows[0].PolicyAction | Should -Be 'WhatIf'
            Assert-NoGraphWrite
        }

        It 'creates a report-only policy when applied' {
            $rows = @(& (Get-ScriptPath 'export-admin-mfa-status.ps1') -TenantId $global:G2Reports.Tenant -CreateConditionalAccessPolicy -ExcludeUserId 'ffffffff-0000-4000-8000-000000000001' -Apply -Confirm:$false)
            $rows[0].PolicyAction | Should -Match '^CreatedReportOnly'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter {
                $Method -eq 'POST' -and $Body.state -eq 'enabledForReportingButNotEnforced' -and $Body.conditions.users.includeRoles.Count -eq 14 -and $Body.conditions.users.excludeUsers -contains 'ffffffff-0000-4000-8000-000000000001'
            }
        }
    }

    Context 'check-hibp-breaches.ps1' {
        BeforeEach {
            # Calls are counted in the mock body, not with Should -Invoke. When an earlier test file has shadowed
            # Get-Secret with a global stub, Pester's call history for Get-Secret can come back empty.
            $global:G2Reports.SecretNames = [System.Collections.Generic.List[string]]::new()
            Mock Get-Secret { $global:G2Reports.SecretNames.Add([string]$PesterBoundParameters['Name']); 'fake-hibp-key-for-tests' }
            Mock Invoke-RestMethod {
                @([pscustomobject]@{ Name = 'Example'; Title = 'Example breach'; BreachDate = '2025-06-01'; AddedDate = '2025-07-01T00:00:00Z'; DataClasses = @('Email addresses', 'Passwords'); IsVerified = $true; IsSensitive = $false; IsSpamList = $false })
            }
        }

        It 'checks member addresses only and reports breaches' {
            $rows = @(& (Get-ScriptPath 'check-hibp-breaches.ps1') -TenantId $global:G2Reports.Tenant -DelayMilliseconds 0)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].Email | Should -Be 'adele.vance@contoso.com'
            $rows[0].Status | Should -Be 'Breached'
            $rows[0].PasswordChangedSinceBreach | Should -BeTrue
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -ParameterFilter { $Uri -like 'https://haveibeenpwned.com/api/v3/breachedaccount/adele.vance%40contoso.com*' -and $Headers['hibp-api-key'] -eq 'fake-hibp-key-for-tests' }
            # One secret read. Its name is checked when the mock sees it: with the global Get-Secret stubs that other
            # test files create on runners without SecretManagement, Pester can't always bind -Name in the mock.
            $secretNames = @($global:G2Reports.SecretNames)
            $secretNames.Count | Should -Be 1
            if ($secretNames[0]) { $secretNames[0] | Should -Be 'HibpApiKey' }
            Assert-NoGraphWrite
        }

        It 'returns a NoBreachesFound row for a clean customer' {
            Mock Invoke-RestMethod {
                $notFound = [System.Exception]::new('Response status code does not indicate success: 404 (Not Found).')
                $notFound | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = 404 })
                throw $notFound
            }
            $rows = @(& (Get-ScriptPath 'check-hibp-breaches.ps1') -TenantId $global:G2Reports.Tenant -DelayMilliseconds 0)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].Status | Should -Be 'NoBreachesFound'
            $rows[0].Email | Should -Be '1 address(es) checked'
        }
    }

    Context 'export-licensed-user-contacts.ps1' {
        It 'exports enabled licensed members with a segment and no consent assumed' {
            $rows = @(& (Get-ScriptPath 'export-licensed-user-contacts.ps1') -TenantId $global:G2Reports.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].Email | Should -Be 'adele.vance@contoso.com'
            $rows[0].Segment | Should -Be 'Business'
            $rows[0].ConsentStatus | Should -Be 'NotRequested'
            Assert-NoGraphWrite
        }
    }
}
