#Requires -Version 7.4
# Offline smoke tests for scripts/users-and-licensing. Every MspGdap call that would reach a tenant
# and every Exchange Online cmdlet is mocked or stubbed. Nothing here contacts Microsoft or any
# customer tenant.

BeforeAll {
    $script:RepoRoot = (Resolve-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath '../..')).Path
    $script:ScriptRoot = Join-Path -Path $script:RepoRoot -ChildPath 'scripts/users-and-licensing'
    $srcPath = Join-Path -Path $script:RepoRoot -ChildPath 'src'
    if (($env:PSModulePath -split [System.IO.Path]::PathSeparator) -notcontains $srcPath) {
        $env:PSModulePath = $srcPath + [System.IO.Path]::PathSeparator + $env:PSModulePath
    }
    Import-Module (Join-Path -Path $srcPath -ChildPath 'MspGdap/MspGdap.psd1') -Force

    $global:G2Users = @{}
    $global:G2Users.CustomerId = '11111111-1111-4111-8111-111111111111'
    $global:G2Users.Tenant = 'contoso.onmicrosoft.com'
    $global:G2Users.SkuSpb = 'aaaaaaaa-0000-4000-8000-000000000001'
    $global:G2Users.SkuEms = 'aaaaaaaa-0000-4000-8000-000000000002'
    $global:G2Users.GroupId = 'bbbbbbbb-0000-4000-8000-000000000001'
    $global:G2Users.State = @{ ManagerSet = $false; DomainChanged = $false; AuthenticatorChanged = $false }
    # Deleted 7 days ago, so 23 days remain before the 30-day automatic purge.
    $global:G2Users.DeletedAt = [datetimeoffset]::UtcNow.AddDays(-7).ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)

    function Get-ScriptPath {
        param([Parameter(Mandatory)][string]$Name)
        Join-Path -Path $script:ScriptRoot -ChildPath $Name
    }

    # Stubs for Exchange Online cmdlets, so the tests run without ExchangeOnlineManagement and so
    # Mock has a command to bind to.
    function Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param() }
    function Get-EXOMailbox { [CmdletBinding()] param($Identity, $ResultSize, $RecipientTypeDetails, $Properties, $PropertySets, $Filter) }
    function Get-EXOMailboxStatistics { [CmdletBinding()] param($Identity, $ExchangeGuid, $Properties, $PropertySets) }
    function Get-EXORecipient { [CmdletBinding()] param($Identity, $Filter, $Properties, $ResultSize) }
    function Set-Mailbox { [CmdletBinding()] param($Identity, $EmailAddresses) }

    # A fake Microsoft Graph that answers by URI. Writes are recorded by the Should -Invoke checks.
    function Invoke-FakeGraph {
        param([string]$Uri, [string]$Method)
        if ($Method -and $Method -ne 'GET') {
            if ($Uri -match '/manager/\$ref$') { $global:G2Users.State.ManagerSet = $true }
            if ($Uri -match '^v1\.0/domains/') { $global:G2Users.State.DomainChanged = $true }
            if ($Uri -match 'microsoftAuthenticator$') { $global:G2Users.State.AuthenticatorChanged = $true }
            return $null
        }
        switch -Regex ($Uri) {
            '^v1\.0/organization' {
                return [pscustomobject]@{
                    id              = $global:G2Users.CustomerId
                    displayName     = 'Contoso'
                    verifiedDomains = @(
                        [pscustomobject]@{ name = 'contoso.onmicrosoft.com'; isInitial = $true; isDefault = $false }
                        [pscustomobject]@{ name = 'contoso.com'; isInitial = $false; isDefault = $true }
                    )
                }
            }
            '^v1\.0/subscribedSkus' {
                return @(
                    [pscustomobject]@{ skuId = $global:G2Users.SkuSpb; skuPartNumber = 'SPB'; consumedUnits = 2; prepaidUnits = [pscustomobject]@{ enabled = 5; warning = 0; suspended = 0 } }
                    [pscustomobject]@{ skuId = $global:G2Users.SkuEms; skuPartNumber = 'EMS'; consumedUnits = 1; prepaidUnits = [pscustomobject]@{ enabled = 1; warning = 0; suspended = 0 } }
                )
            }
            '^v1\.0/directory/deletedItems/microsoft\.graph\.user' {
                return @(
                    [pscustomobject]@{ id = 'dddddddd-0000-4000-8000-000000000001'; displayName = 'Old Account'; userPrincipalName = 'dddddddd000040008000000000000001old@contoso.com'; mail = 'old@contoso.com'; deletedDateTime = $global:G2Users.DeletedAt }
                    [pscustomobject]@{ id = 'dddddddd-0000-4000-8000-000000000002'; displayName = 'Bold Account'; userPrincipalName = 'dddddddd000040008000000000000002bold@contoso.com'; mail = $null; deletedDateTime = $global:G2Users.DeletedAt }
                )
            }
            '^v1\.0/users/[^/?]+/manager\?' {
                if ($global:G2Users.State.ManagerSet) { return [pscustomobject]@{ id = 'cccccccc-0000-4000-8000-000000000002'; userPrincipalName = 'megan@contoso.com' } }
                throw 'Resource manager does not exist (404).'
            }
            '^v1\.0/users/leaver' { return [pscustomobject]@{ id = 'cccccccc-0000-4000-8000-000000000001'; displayName = 'Leaver'; userPrincipalName = 'leaver@contoso.com'; accountEnabled = $true } }
            '^v1\.0/users/megan' { return [pscustomobject]@{ id = 'cccccccc-0000-4000-8000-000000000002'; displayName = 'Megan'; userPrincipalName = 'megan@contoso.com'; accountEnabled = $true } }
            '^v1\.0/users/(adele|cccccccc-0000-4000-8000-000000000003)' {
                return [pscustomobject]@{ id = 'cccccccc-0000-4000-8000-000000000003'; displayName = 'Adele Vance'; userPrincipalName = 'adele.vance@contoso.com'; accountEnabled = $true; userType = 'Member'; createdDateTime = '2024-01-01T00:00:00Z'; assignedLicenses = @([pscustomobject]@{ skuId = $global:G2Users.SkuSpb; disabledPlans = @() }); licenseAssignmentStates = @([pscustomobject]@{ skuId = $global:G2Users.SkuSpb; assignedByGroup = $null }) }
            }
            '^v1\.0/users/(shared|eeeeeeee-0000-4000-8000-00000000000[12])' {
                return [pscustomobject]@{ id = 'eeeeeeee-0000-4000-8000-000000000001'; userPrincipalName = 'shared@contoso.com'; assignedLicenses = @([pscustomobject]@{ skuId = $global:G2Users.SkuSpb; disabledPlans = @() }); licenseAssignmentStates = @([pscustomobject]@{ skuId = $global:G2Users.SkuSpb; assignedByGroup = $null }) }
            }
            '^v1\.0/users[?]' {
                $filter = if ($Uri -match 'assignedLicenses/any') { 'sku' } else { 'all' }
                $licensed = [pscustomobject]@{
                    id                         = 'cccccccc-0000-4000-8000-000000000003'
                    displayName                = 'Adele Vance'
                    userPrincipalName          = 'adele.vance@contoso.com'
                    accountEnabled             = $true
                    userType                   = 'Member'
                    createdDateTime            = '2024-01-01T00:00:00Z'
                    assignedLicenses           = @(
                        [pscustomobject]@{ skuId = $global:G2Users.SkuSpb; disabledPlans = @('ffffffff-0000-4000-8000-000000000001') }
                        [pscustomobject]@{ skuId = $global:G2Users.SkuEms; disabledPlans = @() }
                    )
                    licenseAssignmentStates    = @(
                        [pscustomobject]@{ skuId = $global:G2Users.SkuSpb; assignedByGroup = $null }
                        [pscustomobject]@{ skuId = $global:G2Users.SkuEms; assignedByGroup = $global:G2Users.GroupId }
                    )
                }
                if ($filter -eq 'sku') { return $licensed }
                $unlicensed = [pscustomobject]@{
                    id                      = 'cccccccc-0000-4000-8000-000000000004'
                    displayName             = 'Room 1'
                    userPrincipalName       = 'room1@contoso.com'
                    accountEnabled          = $false
                    userType                = 'Member'
                    createdDateTime         = '2024-01-01T00:00:00Z'
                    assignedLicenses        = @()
                    licenseAssignmentStates = @()
                }
                return @($licensed, $unlicensed)
            }
            '^v1\.0/domains\?' {
                return @(
                    [pscustomobject]@{ id = 'contoso.com'; isVerified = $true; authenticationType = 'Managed'; passwordValidityPeriodInDays = 90; passwordNotificationWindowInDays = 14 }
                    [pscustomobject]@{ id = 'contoso.onmicrosoft.com'; isVerified = $true; authenticationType = 'Managed'; passwordValidityPeriodInDays = 2147483647; passwordNotificationWindowInDays = 14 }
                    [pscustomobject]@{ id = 'pending.contoso.com'; isVerified = $false; authenticationType = 'Managed'; passwordValidityPeriodInDays = 90; passwordNotificationWindowInDays = 14 }
                )
            }
            '^v1\.0/domains/[^?]+\?' {
                $value = if ($global:G2Users.State.DomainChanged) { 2147483647 } else { 90 }
                return [pscustomobject]@{ id = 'contoso.com'; isDefault = $false; isVerified = $true; passwordValidityPeriodInDays = $value }
            }
            'microsoftAuthenticator$' {
                if ($global:G2Users.State.AuthenticatorChanged) {
                    return [pscustomobject]@{ state = 'enabled'; includeTargets = @([pscustomobject]@{ id = 'all_users'; targetType = 'group'; authenticationMode = 'any' }) }
                }
                return [pscustomobject]@{ state = 'enabled'; includeTargets = @([pscustomobject]@{ id = 'all_users'; targetType = 'group'; authenticationMode = 'push'; isRegistrationRequired = $false }) }
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
        }
    }

    function Assert-NoGraphWrite {
        Should -Invoke Invoke-MspGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -and $Method -ne 'GET' }
    }
}

AfterAll {
    Remove-Variable -Name G2Users -Scope Global -ErrorAction SilentlyContinue
}

Describe 'users-and-licensing scripts' {
    BeforeEach {
        $global:G2Users.State.ManagerSet = $false
        $global:G2Users.State.DomainChanged = $false
        $global:G2Users.State.AuthenticatorChanged = $false
        Mock Resolve-MspTenantId {
            if ($Tenant -like 'fabrikam*') { throw 'Tenant not found.' }
            $global:G2Users.CustomerId
        }
        Mock Invoke-MspGraphRequest { Invoke-FakeGraph -Uri $Uri -Method $Method }
        Mock Get-MspAuthHeader { throw 'Get-MspAuthHeader must not be called in tests.' }
        Mock Get-MspCustomer {
            [pscustomobject]@{ TenantId = $global:G2Users.CustomerId; DisplayName = 'Contoso (partner view)'; DefaultDomainName = 'contoso.com'; GdapStatus = 'active' }
            [pscustomobject]@{ TenantId = '99999999-9999-4999-8999-999999999999'; DisplayName = 'Expired Customer'; DefaultDomainName = 'expired.com'; GdapStatus = 'expired' }
        }
        Mock Connect-MspExchangeOnline { [pscustomobject]@{ TenantId = $TenantId } }
        Mock Disconnect-ExchangeOnline { }
        Mock Set-Mailbox { }
    }

    Context 'export-licensed-users.ps1' {
        It 'returns licensed users with the customer columns' {
            $rows = @(& (Get-ScriptPath 'export-licensed-users.ps1') -TenantId $global:G2Users.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].CustomerTenantId | Should -Be $global:G2Users.CustomerId
            $rows[0].CustomerName | Should -Be 'Contoso'
            $rows[0].Licences | Should -Be 'EMS, SPB'
            Assert-NoGraphWrite
        }

        It 'includes unlicensed users and writes a CSV' {
            $csv = Join-Path -Path $TestDrive -ChildPath 'users.csv'
            $rows = @(& (Get-ScriptPath 'export-licensed-users.ps1') -TenantId $global:G2Users.Tenant -IncludeUnlicensed -OutputPath $csv)
            $rows.Count | Should -Be 2
            @(Import-Csv -Path $csv).Count | Should -Be 2
        }

        It 'records a failed customer and carries on with the next one' {
            $rows = @(& (Get-ScriptPath 'export-licensed-users.ps1') -TenantId 'fabrikam.onmicrosoft.com', $global:G2Users.Tenant -WarningAction SilentlyContinue)
            $rows[0].Status | Should -Be 'Failed'
            $rows[0].Error | Should -Match 'not found'
            $rows[1].Status | Should -Be 'OK'
        }

        It 'uses only active GDAP customers with -AllCustomers' {
            $rows = @(& (Get-ScriptPath 'export-licensed-users.ps1') -AllCustomers)
            $rows.Count | Should -Be 1
            $rows[0].CustomerName | Should -Be 'Contoso (partner view)'
        }

        It 'accepts tenant IDs from the pipeline' {
            $rows = @($global:G2Users.Tenant | & (Get-ScriptPath 'export-licensed-users.ps1'))
            $rows.Count | Should -Be 1
        }
    }

    Context 'test-customer-graph-access.ps1' {
        It 'reports each check and closes the Exchange session' {
            Mock Get-EXOMailbox { [pscustomobject]@{ UserPrincipalName = 'adele.vance@contoso.com' } }
            $rows = @(& (Get-ScriptPath 'test-customer-graph-access.ps1') -TenantId $global:G2Users.Tenant -IncludeExchange)
            Assert-CustomerShape -Rows $rows
            $rows[0].GraphOrganisationRead | Should -Be 'Passed'
            $rows[0].GraphUserRead | Should -Be 'Passed'
            $rows[0].ExchangeConnect | Should -Be 'Passed'
            $rows[0].InitialDomain | Should -Be 'contoso.onmicrosoft.com'
            $rows[0].Status | Should -Be 'OK'
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
            Assert-NoGraphWrite
        }
    }

    Context 'get-customer-users.ps1' {
        It 'finds the customer by name and lists users' {
            $rows = @(& (Get-ScriptPath 'get-customer-users.ps1') -CustomerName 'contoso')
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            ($rows | Where-Object UserPrincipalName -eq 'adele.vance@contoso.com').IsLicensed | Should -BeTrue
            ($rows | Where-Object UserPrincipalName -eq 'room1@contoso.com').IsLicensed | Should -BeFalse
        }

        It 'returns one user with -UserPrincipalName' {
            $rows = @(& (Get-ScriptPath 'get-customer-users.ps1') -TenantId $global:G2Users.Tenant -UserPrincipalName 'adele.vance@contoso.com')
            $rows.Count | Should -Be 1
            $rows[0].UserPrincipalName | Should -Be 'adele.vance@contoso.com'
            Assert-NoGraphWrite
        }
    }

    Context 'get-users-by-licence.ps1' {
        It 'lists users with the SKU' {
            $rows = @(& (Get-ScriptPath 'get-users-by-licence.ps1') -TenantId $global:G2Users.Tenant -SkuPartNumber 'spb')
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 1
            $rows[0].SkuPartNumber | Should -Be 'SPB'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Uri -match 'assignedLicenses/any\(x:x/skuId eq aaaaaaaa-0000-4000-8000-000000000001\)' }
        }

        It 'lists SKUs with -ListSkus' {
            $rows = @(& (Get-ScriptPath 'get-users-by-licence.ps1') -TenantId $global:G2Users.Tenant -ListSkus)
            $rows.Count | Should -Be 2
            ($rows | Where-Object SkuPartNumber -eq 'SPB').EnabledUnits | Should -Be 5
        }

        It 'accepts a wildcard, as the original -match did' {
            $rows = @(& (Get-ScriptPath 'get-users-by-licence.ps1') -TenantId $global:G2Users.Tenant -SkuPartNumber 'S*')
            $rows.Count | Should -Be 1
            $rows[0].SkuPartNumber | Should -Be 'SPB'
        }

        It 'reports a SKU the tenant does not have' {
            $rows = @(& (Get-ScriptPath 'get-users-by-licence.ps1') -TenantId $global:G2Users.Tenant -SkuPartNumber 'ENTERPRISEPREMIUM')
            $rows[0].Status | Should -Be 'SkuNotFound'
        }
    }

    Context 'export-licence-restore-commands.ps1' {
        It 'creates restore commands for direct licences only' {
            $rows = @(& (Get-ScriptPath 'export-licence-restore-commands.ps1') -TenantId $global:G2Users.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            $direct = $rows | Where-Object SkuPartNumber -eq 'SPB'
            $direct.AssignedBy | Should -Be 'Direct'
            $direct.RestoreCommand | Should -Match "assignLicense"
            $direct.RestoreCommand | Should -Match 'ffffffff-0000-4000-8000-000000000001'
            $tokens = $null
            $errors = $null
            $null = [System.Management.Automation.Language.Parser]::ParseInput($direct.RestoreCommand, [ref]$tokens, [ref]$errors)
            $errors.Count | Should -Be 0
            $group = $rows | Where-Object SkuPartNumber -eq 'EMS'
            $group.AssignedBy | Should -Be "Group:$($global:G2Users.GroupId)"
            $group.RestoreCommand | Should -BeNullOrEmpty
            Assert-NoGraphWrite
        }
    }

    Context 'remove-deleted-user.ps1' {
        It 'lists deleted users with the purge date and days left, without changing anything' {
            $rows = @(& (Get-ScriptPath 'remove-deleted-user.ps1') -TenantId $global:G2Users.Tenant 6>$null)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            $rows[0].Mail | Should -Be 'old@contoso.com'
            $rows[0].OriginalUserPrincipalName | Should -Be 'old@contoso.com'
            $rows[0].Status | Should -Be 'OK'
            $rows[0].DaysUntilPurge | Should -Be 23
            $expectedPurge = ([datetimeoffset]::Parse($global:G2Users.DeletedAt, [System.Globalization.CultureInfo]::InvariantCulture)).AddDays(30).UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
            $rows[0].PurgeDateTime | Should -Be $expectedPurge
            $rows[0].Guidance | Should -Match 'Microsoft Entra admin center'
            $rows[0].Guidance | Should -Match 'Delete permanently'
            Assert-NoGraphWrite
        }

        It 'writes the manual permanent-delete guidance to the information stream' {
            $info = & (Get-ScriptPath 'remove-deleted-user.ps1') -TenantId $global:G2Users.Tenant 6>&1 | Where-Object { $_ -is [System.Management.Automation.InformationRecord] }
            @($info).Count | Should -Be 1
            [string]$info[0].MessageData | Should -Match 'least privilege'
        }

        It 'matches the deleted name exactly, so old@ never selects bold@, and reports names it cannot find' {
            $rows = @(& (Get-ScriptPath 'remove-deleted-user.ps1') -TenantId $global:G2Users.Tenant -UserPrincipalName 'ld@contoso.com', 'old@contoso.com' 6>$null)
            ($rows | Where-Object Status -eq 'NotFound').UserPrincipalName | Should -Be 'ld@contoso.com'
            @($rows | Where-Object Status -eq 'OK').Count | Should -Be 1
            ($rows | Where-Object Status -eq 'OK').DeletedObjectId | Should -Be 'dddddddd-0000-4000-8000-000000000001'
            Assert-NoGraphWrite
        }

        It 'has no -Apply switch and makes no Graph write, because it never permanently deletes' {
            $command = Get-Command -Name (Get-ScriptPath 'remove-deleted-user.ps1')
            $command.Parameters.Keys | Should -Not -Contain 'Apply'
            $command.Parameters.Keys | Should -Not -Contain 'WhatIf'
            Get-Content -LiteralPath (Get-ScriptPath 'remove-deleted-user.ps1') -Raw | Should -Not -Match '-Method\s+DELETE'
        }
    }

    Context 'set-onedrive-successor.ps1' {
        It 'reports the planned manager change' {
            $rows = @(& (Get-ScriptPath 'set-onedrive-successor.ps1') -TenantId $global:G2Users.Tenant -DepartingUserPrincipalName 'leaver@contoso.com' -SuccessorUserPrincipalName 'megan@contoso.com')
            Assert-CustomerShape -Rows $rows
            $rows[0].ManagerAction | Should -Be 'WouldSet'
            Assert-NoGraphWrite
        }

        It 'makes no change under -WhatIf' {
            $rows = @(& (Get-ScriptPath 'set-onedrive-successor.ps1') -TenantId $global:G2Users.Tenant -DepartingUserPrincipalName 'leaver@contoso.com' -SuccessorUserPrincipalName 'megan@contoso.com' -Apply -RemoveDepartingUser -WhatIf -WarningAction SilentlyContinue)
            $rows[0].ManagerAction | Should -Be 'WhatIf'
            $rows[0].RemoveAction | Should -Be 'WhatIf'
            Assert-NoGraphWrite
        }

        It 'sets the manager and deletes the user when applied' {
            $rows = @(& (Get-ScriptPath 'set-onedrive-successor.ps1') -TenantId $global:G2Users.Tenant -DepartingUserPrincipalName 'leaver@contoso.com' -SuccessorUserPrincipalName 'megan@contoso.com' -Apply -RemoveDepartingUser -Confirm:$false)
            $rows[0].ManagerAction | Should -Be 'Set'
            $rows[0].RemoveAction | Should -Be 'Deleted'
            $rows[0].NextStep | Should -BeOfType [string]
            $rows[0].NextStep | Should -Match 'emailed a link'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'PUT' -and $Uri -match 'manager/\$ref$' }
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'DELETE' }
        }
    }

    Context 'fix-proxy-address-conflict.ps1' {
        BeforeEach {
            Mock Get-EXORecipient {
                [pscustomobject]@{ PrimarySmtpAddress = 'accounts@contoso.com'; RecipientTypeDetails = 'SharedMailbox'; EmailAddresses = @('SMTP:accounts@contoso.com', 'smtp:accounts@contoso.onmicrosoft.com'); ExternalDirectoryObjectId = 'eeeeeeee-0000-4000-8000-000000000002' }
            }
            Mock Get-EXOMailbox {
                [pscustomobject]@{ UserPrincipalName = 'accounts@fabrikam.com'; PrimarySmtpAddress = 'accounts@contoso.com'; ExchangeGuid = '12345678-0000-4000-8000-000000000001'; ExternalDirectoryObjectId = 'eeeeeeee-0000-4000-8000-000000000002'; EmailAddresses = @('SMTP:accounts@contoso.com', 'smtp:accounts@fabrikam.com') }
            }
        }

        It 'reports who holds the address' {
            $rows = @(& (Get-ScriptPath 'fix-proxy-address-conflict.ps1') -TenantId $global:G2Users.Tenant -EmailAddress 'accounts@contoso.com')
            Assert-CustomerShape -Rows $rows
            $rows[0].HolderCount | Should -Be 1
            $rows[0].Holders | Should -Match 'primary'
            Should -Invoke Set-Mailbox -Times 0 -Exactly
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
        }

        It 'stops when the address is primary and no new primary is given' {
            $rows = @(& (Get-ScriptPath 'fix-proxy-address-conflict.ps1') -TenantId $global:G2Users.Tenant -EmailAddress 'accounts@contoso.com' -RemoveFrom 'accounts@fabrikam.com' -Apply -Confirm:$false)
            $rows[0].PrimaryAction | Should -Be 'NeedsNewPrimarySmtpAddress'
            $rows[0].RemoveAction | Should -Be 'Blocked'
            Should -Invoke Set-Mailbox -Times 0 -Exactly
        }

        It 'makes no change under -WhatIf' {
            $rows = @(& (Get-ScriptPath 'fix-proxy-address-conflict.ps1') -TenantId $global:G2Users.Tenant -EmailAddress 'accounts@contoso.com' -RemoveFrom 'accounts@fabrikam.com' -NewPrimarySmtpAddress 'accounts@fabrikam.com' -NewUserPrincipalName 'accounts2@fabrikam.com' -DefaultDomain 'contoso.com' -Apply -WhatIf)
            $rows[0].UpnAction | Should -Be 'WhatIf'
            $rows[0].PrimaryAction | Should -Be 'WhatIf'
            $rows[0].DefaultDomainAction | Should -Be 'WhatIf'
            Should -Invoke Set-Mailbox -Times 0 -Exactly
            Assert-NoGraphWrite
        }

        It 'adds the new primary before removing the old address' {
            $global:G2Users.SetMailboxCalls = [System.Collections.Generic.List[string]]::new()
            Mock Set-Mailbox { $global:G2Users.SetMailboxCalls.Add(($EmailAddresses.Keys | Select-Object -First 1)) }
            $null = & (Get-ScriptPath 'fix-proxy-address-conflict.ps1') -TenantId $global:G2Users.Tenant -EmailAddress 'accounts@contoso.com' -RemoveFrom 'accounts@fabrikam.com' -NewPrimarySmtpAddress 'accounts@fabrikam.com' -Apply -Confirm:$false -WarningAction SilentlyContinue
            $global:G2Users.SetMailboxCalls | Should -Be @('Add', 'Remove')
        }
    }

    Context 'remove-shared-mailbox-licences.ps1' {
        BeforeEach {
            Mock Get-EXOMailbox {
                [pscustomobject]@{ DisplayName = 'Shared Small'; PrimarySmtpAddress = 'shared@contoso.com'; ExchangeGuid = '12345678-0000-4000-8000-000000000009'; ExternalDirectoryObjectId = 'eeeeeeee-0000-4000-8000-000000000001'; ArchiveStatus = 'None'; LitigationHoldEnabled = $false; InPlaceHolds = @() }
                [pscustomobject]@{ DisplayName = 'Shared On Hold'; PrimarySmtpAddress = 'hold@contoso.com'; ExternalDirectoryObjectId = 'eeeeeeee-0000-4000-8000-000000000002'; ArchiveStatus = 'None'; LitigationHoldEnabled = $true; InPlaceHolds = @() }
            }
            Mock Get-EXOMailboxStatistics { [pscustomobject]@{ TotalItemSize = '1.5 GB (1,610,612,736 bytes)' } }
        }

        It 'flags mailboxes that must keep a licence and makes no change by default' {
            $rows = @(& (Get-ScriptPath 'remove-shared-mailbox-licences.ps1') -TenantId $global:G2Users.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            ($rows | Where-Object PrimarySmtpAddress -eq 'shared@contoso.com').Action | Should -Be 'WouldRemove'
            ($rows | Where-Object PrimarySmtpAddress -eq 'shared@contoso.com').SizeGB | Should -Be 1.5
            ($rows | Where-Object PrimarySmtpAddress -eq 'hold@contoso.com').Action | Should -Be 'KeepLicence'
            ($rows | Where-Object PrimarySmtpAddress -eq 'hold@contoso.com').KeepLicenceReason | Should -Match 'Litigation hold'
            Should -Invoke Get-EXOMailboxStatistics -Times 1 -Exactly -ParameterFilter { $ExchangeGuid -eq '12345678-0000-4000-8000-000000000009' }
            Should -Invoke Disconnect-ExchangeOnline -Times 1 -Exactly
            Assert-NoGraphWrite
        }

        It 'removes nothing under -WhatIf' {
            $rows = @(& (Get-ScriptPath 'remove-shared-mailbox-licences.ps1') -TenantId $global:G2Users.Tenant -Apply -WhatIf)
            ($rows | Where-Object PrimarySmtpAddress -eq 'shared@contoso.com').Action | Should -Be 'WhatIf'
            Assert-NoGraphWrite
        }
    }

    Context 'set-password-expiration-policy.ps1' {
        It 'reports verified domains and makes no change by default' {
            $rows = @(& (Get-ScriptPath 'set-password-expiration-policy.ps1') -TenantId $global:G2Users.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows.Count | Should -Be 2
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'WouldSetNeverExpires'
            ($rows | Where-Object Domain -eq 'contoso.onmicrosoft.com').Action | Should -Be 'AlreadyNeverExpires'
            Assert-NoGraphWrite
        }

        It 'adds per-user rows with -IncludeUsers' {
            $rows = @(& (Get-ScriptPath 'set-password-expiration-policy.ps1') -TenantId $global:G2Users.Tenant -IncludeUsers)
            Assert-CustomerShape -Rows $rows
            @($rows | Where-Object RowType -eq 'User').Count | Should -Be 2
            @($rows | Where-Object RowType -eq 'Domain').Count | Should -Be 2
            Assert-NoGraphWrite
        }

        It 'makes no change under -WhatIf' {
            $rows = @(& (Get-ScriptPath 'set-password-expiration-policy.ps1') -TenantId $global:G2Users.Tenant -Apply -WhatIf)
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'WhatIf'
            Assert-NoGraphWrite
        }

        It 'sets and confirms the policy when applied' {
            $rows = @(& (Get-ScriptPath 'set-password-expiration-policy.ps1') -TenantId $global:G2Users.Tenant -Apply -Confirm:$false)
            ($rows | Where-Object Domain -eq 'contoso.com').Action | Should -Be 'Changed'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter { $Method -eq 'PATCH' -and $Body.passwordValidityPeriodInDays -eq 2147483647 }
        }
    }

    Context 'enable-authenticator-passwordless.ps1' {
        It 'reports push-only targets as not allowing passwordless' {
            $rows = @(& (Get-ScriptPath 'enable-authenticator-passwordless.ps1') -TenantId $global:G2Users.Tenant)
            Assert-CustomerShape -Rows $rows
            $rows[0].PasswordlessAllowed | Should -BeFalse
            $rows[0].Targets | Should -Be 'all_users:push'
            $rows[0].Action | Should -Be 'WouldChange'
            Assert-NoGraphWrite
        }

        It 'makes no change under -WhatIf' {
            $rows = @(& (Get-ScriptPath 'enable-authenticator-passwordless.ps1') -TenantId $global:G2Users.Tenant -Apply -WhatIf)
            $rows[0].Action | Should -Be 'WhatIf'
            Assert-NoGraphWrite
        }

        It 'keeps existing targets and switches them to any when applied' {
            $rows = @(& (Get-ScriptPath 'enable-authenticator-passwordless.ps1') -TenantId $global:G2Users.Tenant -Apply -Confirm:$false)
            $rows[0].Action | Should -Be 'Changed'
            Should -Invoke Invoke-MspGraphRequest -Times 1 -Exactly -ParameterFilter {
                $Method -eq 'PATCH' -and $Body.'@odata.type' -eq '#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration' -and $Body.includeTargets[0].authenticationMode -eq 'any' -and $Body.includeTargets[0].isRegistrationRequired -eq $false -and -not $Body.ContainsKey('state')
            }
        }
    }
}
