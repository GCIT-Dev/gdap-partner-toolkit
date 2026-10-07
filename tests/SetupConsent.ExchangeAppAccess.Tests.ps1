#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Support/SetupConsent.TestSupport.ps1')
    Mock Start-Sleep {}
    $customer = $global:MspTest.CustomerTenantId
    $automation = $global:MspTest.AutomationAppId
    $manageAsApp = 'dc50a0fb-09a3-484d-be87-e023b12c6440'
    $exchangeAdmin = '29232cdf-9323-42fd-ade2-1d097af3e4de'

    function Reset-ExchangeFake {
        param([bool]$ExoPresent = $true, [bool]$AppSpPresent = $true, [bool]$HasAppRole = $false, [string[]]$Roles = @(), [string]$RolePost = 'apply', [string]$SpPost = 'apply', [bool]$AppRolePostApplies = $true, [string]$Owner = $global:MspTest.PartnerTenantId, [bool]$RegisteredInPartner = $true)
        $global:MspTest.State = @{ ExoPresent = $ExoPresent; AppSpPresent = $AppSpPresent; HasAppRole = $HasAppRole; Roles = [System.Collections.Generic.List[string]]$Roles; RolePost = $RolePost; SpPost = $SpPost; AppRolePostApplies = $AppRolePostApplies; Owner = $Owner; Registered = $RegisteredInPartner }
        Set-FakeGraphRoute @(
            New-FakeRoute -Pattern "servicePrincipals\?\`$filter=appId eq '00000002-0000-0ff1-ce00-000000000000'" -Response {
                if ($global:MspTest.State.ExoPresent) { [pscustomobject]@{ id = 'exo-sp'; appId = '00000002-0000-0ff1-ce00-000000000000'; appRoles = @([pscustomobject]@{ id = $manageAsApp; value = 'Exchange.ManageAsApp' }) } }
            }
            New-FakeRoute -Pattern "servicePrincipals\?\`$filter=appId eq '$automation'" -Response {
                if ($global:MspTest.State.AppSpPresent) { [pscustomobject]@{ id = 'auto-sp'; appId = $automation; displayName = 'Automation'; appOwnerOrganizationId = $global:MspTest.State.Owner } }
            }
            New-FakeRoute -Pattern "applications\?\`$filter=appId eq '$automation'" -Response {
                param($Body, $Uri, $TenantId)
                if ($TenantId) { throw 'The partner app registration must be read in the partner tenant.' }
                if ($global:MspTest.State.Registered) { [pscustomobject]@{ id = 'auto-app'; appId = $automation; displayName = 'Automation' } }
            }
            New-FakeRoute -Method POST -Pattern 'servicePrincipals$' -Response {
                if ($global:MspTest.State.SpPost -eq 'forbidden') { throw (New-FakeHttpError -StatusCode 403 -Message 'Insufficient privileges') }
                $global:MspTest.State.AppSpPresent = $true
                [pscustomobject]@{ id = 'auto-sp' }
            }
            New-FakeRoute -Pattern 'servicePrincipals/auto-sp/appRoleAssignments' -Response {
                if ($global:MspTest.State.HasAppRole) { [pscustomobject]@{ resourceId = 'exo-sp'; appRoleId = $manageAsApp } }
            }
            New-FakeRoute -Method POST -Pattern 'servicePrincipals/exo-sp/appRoleAssignedTo' -Response {
                if ($global:MspTest.State.AppRolePostApplies) { $global:MspTest.State.HasAppRole = $true }
                [pscustomobject]@{ id = 'ara' }
            }
            New-FakeRoute -Pattern 'roleManagement/directory/roleAssignments\?' -Response {
                param($Body, $Uri)
                $role = [regex]::Match($Uri, "roleDefinitionId eq '([^']+)'").Groups[1].Value
                if ($global:MspTest.State.Roles.Contains($role)) { [pscustomobject]@{ id = "ra-$role"; principalId = 'auto-sp'; roleDefinitionId = $role; directoryScopeId = '/' } }
            }
            New-FakeRoute -Method POST -Pattern 'roleManagement/directory/roleAssignments$' -Response {
                param($Body)
                switch ($global:MspTest.State.RolePost) {
                    'forbidden' { throw (New-FakeHttpError -StatusCode 403 -Message 'Authorization_RequestDenied') }
                    'noeffect' { [pscustomobject]@{ id = 'ra' } }
                    default { $global:MspTest.State.Roles.Add($Body.roleDefinitionId); [pscustomobject]@{ id = 'ra' } }
                }
            }
            New-FakeRoute -Pattern 'organization' -Response { [pscustomobject]@{ id = $customer; displayName = 'Fabrikam'; verifiedDomains = @([pscustomobject]@{ name = 'fabrikam.onmicrosoft.com'; isInitial = $true }) } }
        )
    }
}

AfterAll { Remove-Variable -Name MspTest -Scope Global -ErrorAction SilentlyContinue }

Describe 'Enable-MspExchangeAppAccess' {
    It 'grants Exchange.ManageAsApp and assigns Exchange Administrator with unified RBAC, confirmed by readback' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId 'fabrikam.onmicrosoft.com' -AppId $automation -Confirm:$false
        $r.Success | Should -BeTrue
        $r.InitialDomain | Should -Be 'fabrikam.onmicrosoft.com'
        $post = Get-FakeGraphCall -Method POST -Pattern 'roleAssignments$'
        $post[0].Body.roleDefinitionId | Should -Be $exchangeAdmin
        $post[0].Body.directoryScopeId | Should -Be '/'
        $post[0].Body.principalId | Should -Be 'auto-sp'
        (Get-FakeGraphCall -Method POST -Pattern 'appRoleAssignedTo')[0].Body.appRoleId | Should -Be $manageAsApp
    }
    It 'never uses the legacy directoryRoles API' {
        Reset-ExchangeFake
        $null = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false
        @(Get-FakeGraphCall -Pattern 'directoryRoles').Count | Should -Be 0
    }
    It 'fails, and does not pass, when the role assignment is refused' {
        Reset-ExchangeFake -RolePost 'forbidden'
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Assign role: Exchange Administrator').Detail | Should -Match 'Privileged Role Administrator'
    }
    It 'fails when the role assignment is accepted but not visible on readback' {
        Reset-ExchangeFake -RolePost 'noeffect'
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false -ReadbackTimeoutSeconds 10
        $r.Success | Should -BeFalse
    }
    It 'fails when the app role grant is not visible on readback' {
        Reset-ExchangeFake -AppRolePostApplies $false
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false -ReadbackTimeoutSeconds 10
        ($r.Steps | Where-Object Step -eq 'Grant Exchange.ManageAsApp').Status | Should -Be 'Failed'
        $r.Success | Should -BeFalse
    }
    It 'writes nothing when everything is already in place' {
        Reset-ExchangeFake -HasAppRole $true -Roles @($exchangeAdmin)
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false
        $r.Success | Should -BeTrue
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'writes nothing with -WhatIf' {
        Reset-ExchangeFake -AppSpPresent $false
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -WhatIf
        $r.Outcome | Should -Be 'WhatIf'
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'creates the service principal when missing' {
        Reset-ExchangeFake -AppSpPresent $false
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false
        $r.Success | Should -BeTrue
        ($r.Steps | Where-Object Step -eq 'Automation app service principal').Status | Should -Be 'Changed'
    }
    It 'returns the admin consent URL when the service principal cannot be created' {
        Reset-ExchangeFake -AppSpPresent $false -SpPost 'forbidden'
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false
        $r.Success | Should -BeFalse
        $r.AdminConsentUrl | Should -Be "https://login.microsoftonline.com/$customer/adminconsent?client_id=$automation&scope=https://outlook.office365.com/.default"
        @(Get-FakeGraphCall -Method POST -Pattern 'roleAssignments|appRoleAssignedTo').Count | Should -Be 0
    }
    It 'stops when Exchange Online is not in the tenant' {
        Reset-ExchangeFake -ExoPresent $false
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'refuses Global Administrator' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Role 'Global Administrator' -Confirm:$false
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'refuses the MspGdap partner app' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $global:MspTest.PartnerAppId -Confirm:$false
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Check app').Status | Should -Be 'Failed'
    }
    It 'refuses an app whose customer service principal belongs to another tenant, before any write' {
        Reset-ExchangeFake -Owner '99999999-9999-9999-9999-999999999999'
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Check app owner').Status | Should -Be 'Failed'
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'refuses an app that is not registered in the partner tenant when no service principal exists yet' {
        Reset-ExchangeFake -AppSpPresent $false -RegisteredInPartner $false
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Where-Object Step -eq 'Check app owner').Status | Should -Be 'Failed'
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'continues with a warning for an external app only when -AllowExternalApp is given' {
        Reset-ExchangeFake -Owner '99999999-9999-9999-9999-999999999999'
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -AllowExternalApp -Confirm:$false
        ($r.Steps | Where-Object Step -eq 'Check app owner').Status | Should -Be 'Warning'
        $r.Success | Should -BeTrue
    }
    It 'raises a non-terminating error when the outcome is Failed, so -ErrorAction Stop works' {
        Reset-ExchangeFake -RolePost 'forbidden'
        { Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Confirm:$false -ErrorAction Stop } | Should -Throw -ErrorId 'MspGdap.Enable-MspExchangeAppAccess.Failed*'
    }
    It 'refuses the partner tenant as -TenantId' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $global:MspTest.PartnerTenantId -AppId $automation -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Select-Object -First 1).Detail | Should -Match 'partner tenant'
        @(Get-FakeGraphCall).Count | Should -Be 0
    }
    It 'refuses roles beyond Exchange without -AllowPrivilegedRole' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Role 'Security Administrator' -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Select-Object -First 1).Detail | Should -Match 'AllowPrivilegedRole'
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'refuses roles that do nothing for Exchange app-only access, even with -AllowPrivilegedRole' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Role 'Privileged Role Administrator' -AllowPrivilegedRole -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Select-Object -First 1).Detail | Should -Match 'does not work for Exchange Online app-only'
    }
    It 'refuses role template IDs that are not in the catalogue' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Role 'abababab-abab-abab-abab-abababababab' -Confirm:$false -ErrorAction SilentlyContinue
        $r.Success | Should -BeFalse
        ($r.Steps | Select-Object -First 1).Detail | Should -Match 'catalogue'
    }
    It 'accepts Exchange Recipient Administrator without a switch' {
        Reset-ExchangeFake
        $r = Enable-MspExchangeAppAccess -TenantId $customer -AppId $automation -Role 'Exchange Recipient Administrator' -Confirm:$false
        $r.Success | Should -BeTrue
    }
}

Describe 'Test-MspExchangeAppAccess' {
    It 'passes when everything is in place' {
        Reset-ExchangeFake -HasAppRole $true -Roles @($exchangeAdmin)
        (Test-MspExchangeAppAccess -TenantId $customer -AppId $automation).Success | Should -BeTrue
    }
    It 'fails when the role is missing, and writes nothing' {
        Reset-ExchangeFake -HasAppRole $true
        $r = Test-MspExchangeAppAccess -TenantId $customer -AppId $automation
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'tests an app-only connection and disconnects it' {
        Reset-ExchangeFake -HasAppRole $true -Roles @($exchangeAdmin)
        Mock Connect-MspExchangeOnline { [pscustomobject]@{ ConnectionId = 'c-1'; TenantId = $customer } }
        Mock Disconnect-ExchangeOnline {}
        Mock Get-MspGdapTestOrganizationConfig { [pscustomobject]@{ Name = 'fabrikam' } }
        Mock Get-OrganizationConfig { throw 'The unprefixed cmdlet must not be used.' }
        $r = Test-MspExchangeAppAccess -TenantId $customer -AppId $automation -TestConnection -CertificateThumbprint ('A' * 40)
        $r.Success | Should -BeTrue
        Should -Invoke Connect-MspExchangeOnline -Times 1 -ParameterFilter { $AppOnly -and $Organization -eq 'fabrikam.onmicrosoft.com' -and $Prefix -eq 'MspGdapTest' }
        Should -Invoke Get-MspGdapTestOrganizationConfig -Times 1
        Should -Invoke Disconnect-ExchangeOnline -Times 1 -ParameterFilter { $ConnectionId -eq 'c-1' }
    }
}
