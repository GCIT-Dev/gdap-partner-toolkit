#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Support/SetupConsent.TestSupport.ps1')
    Mock Start-Sleep {}
    Mock Assert-MspModuleAvailable { [version]'2.30.0' }
    $partner = $global:MspTest.PartnerTenantId
    $newAppId = '55555555-5555-5555-5555-555555555555'

    $manifestPath = Join-Path $TestDrive 'manifest.json'
    @{
        manifestVersion = 1; name = 'test'
        requiredResourceAccess = @(
            @{ resourceAppId = '00000003-0000-0000-c000-000000000000'; resourceDisplayName = 'Microsoft Graph'; resourceAccess = @(
                    @{ id = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'; type = 'Scope'; value = 'User.Read' },
                    @{ id = '06da0dbc-49e2-44d2-8312-53f166ab848a'; type = 'Scope'; value = 'Directory.Read.All' }) },
            @{ resourceAppId = '00000002-0000-0ff1-ce00-000000000000'; resourceDisplayName = 'Office 365 Exchange Online'; resourceAccess = @(
                    @{ id = 'ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c'; type = 'Scope'; value = 'Exchange.Manage' }) }
        )
    } | ConvertTo-Json -Depth 8 | Set-Content -Path $manifestPath

    function Reset-AppFake {
        param([object[]]$ExistingApps = @(), [string]$ContextTenant = $partner, [bool]$Connected = $true, [string]$GraphScopeValue = 'User.Read')
        $global:MspTest.State = @{ Created = $null; SpCreated = $false; Grants = @{}; Existing = $ExistingApps; Keys = @(); GraphScopeValue = $GraphScopeValue }
        if ($Connected) {
            $ctx = [pscustomobject]@{ TenantId = $ContextTenant; Account = 'admin@contoso.onmicrosoft.com'; Scopes = @('Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All') }
            Mock Get-MgContext { $ctx }.GetNewClosure()
        }
        else { Mock Get-MgContext { $null } }
        Set-FakeGraphRoute @(
            New-FakeRoute -Collection -Pattern "servicePrincipals\?\`$filter=appId eq '00000003-0000-0000-c000-000000000000'" -Response {
                [pscustomobject]@{ id = 'graph-sp'; appId = '00000003-0000-0000-c000-000000000000'; displayName = 'Microsoft Graph'; appRoles = @()
                    oauth2PermissionScopes = @([pscustomobject]@{ id = 'e1fe6dd8-ba31-4d61-89e7-88639da4683d'; value = $global:MspTest.State.GraphScopeValue }, [pscustomobject]@{ id = '06da0dbc-49e2-44d2-8312-53f166ab848a'; value = 'Directory.Read.All' }) }
            }
            New-FakeRoute -Collection -Pattern "servicePrincipals\?\`$filter=appId eq '00000002-0000-0ff1-ce00-000000000000'" -Response {
                [pscustomobject]@{ id = 'exo-sp'; appId = '00000002-0000-0ff1-ce00-000000000000'; displayName = 'Office 365 Exchange Online'; appRoles = @()
                    oauth2PermissionScopes = @([pscustomobject]@{ id = 'ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c'; value = 'Exchange.Manage' }) }
            }
            New-FakeRoute -Collection -Pattern "servicePrincipals\?\`$filter=appId eq '$newAppId'" -Response {
                if ($global:MspTest.State.SpCreated) { [pscustomobject]@{ id = 'new-sp'; appId = $newAppId; appRoleAssignmentRequired = $false } }
            }
            New-FakeRoute -Method POST -Pattern 'servicePrincipals$' -Response { $global:MspTest.State.SpCreated = $true; [pscustomobject]@{ id = 'new-sp' } }
            New-FakeRoute -Collection -Pattern 'applications\?\$filter=displayName' -Response { $global:MspTest.State.Existing }
            New-FakeRoute -Method POST -Pattern 'applications$' -Response {
                param($Body)
                $global:MspTest.State.Created = $Body
                [pscustomobject]@{ id = 'app-obj'; appId = $newAppId }
            }
            New-FakeRoute -Pattern 'applications/app-obj\?' -Response {
                $b = $global:MspTest.State.Created
                $keys = @(foreach ($k in @($b.keyCredentials)) {
                        if ($k) {
                            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (, [Convert]::FromBase64String($k.key))
                            [pscustomobject]@{ customKeyIdentifier = (ConvertTo-TestKeyIdentifier $cert.Thumbprint); type = 'AsymmetricX509Cert' }
                        }
                    })
                [pscustomobject]@{ id = 'app-obj'; appId = $newAppId; signInAudience = $b.signInAudience; web = $b.web; requiredResourceAccess = $b.requiredResourceAccess; keyCredentials = $keys }
            }
            New-FakeRoute -Collection -Pattern 'oauth2PermissionGrants\?' -Response {
                param($Body, $Uri)
                $resource = [regex]::Match($Uri, "resourceId eq '([^']+)'").Groups[1].Value
                if ($global:MspTest.State.Grants.ContainsKey($resource)) { [pscustomobject]@{ id = "g-$resource"; scope = $global:MspTest.State.Grants[$resource] } }
            }
            New-FakeRoute -Method POST -Pattern 'oauth2PermissionGrants$' -Response {
                param($Body)
                $global:MspTest.State.Grants[$Body.resourceId] = $Body.scope
                [pscustomobject]@{ id = 'g' }
            }
        )
    }
}

AfterAll { Remove-Variable -Name MspTest -Scope Global -ErrorAction SilentlyContinue }

Describe 'New-MspPartnerApp' {
    It 'creates the app, its service principal and partner tenant consent, uploading only the public key' {
        Reset-AppFake
        $cert = New-TestCertificate
        $r = New-MspPartnerApp -DisplayName 'Contoso MSP Worker' -PartnerTenantId $partner -ManifestPath $manifestPath -Certificate $cert -Confirm:$false
        $r.Success | Should -BeTrue
        $r.AppId | Should -Be $newAppId
        $r.ServicePrincipalId | Should -Be 'new-sp'
        $body = $global:MspTest.State.Created
        $body.signInAudience | Should -Be 'AzureADMultipleOrgs'
        $body.web.redirectUris | Should -Be @('http://localhost')
        $body.web.implicitGrantSettings.enableAccessTokenIssuance | Should -BeFalse
        $body.servicePrincipalLockConfiguration.isEnabled | Should -BeTrue
        $body.keyCredentials[0].key | Should -Be ([Convert]::ToBase64String($cert.RawData))
        ($body.requiredResourceAccess | ConvertTo-Json -Depth 6) | Should -Not -Match '"value"'
        $global:MspTest.State.Grants['graph-sp'] | Should -Be 'User.Read Directory.Read.All'
        $global:MspTest.State.Grants['exo-sp'] | Should -Be 'Exchange.Manage'
        $r.NextSteps | Should -Not -BeNullOrEmpty
    }
    It 'stops before writing when the Graph session is in another tenant' {
        Reset-AppFake -ContextTenant '99999999-9999-9999-9999-999999999999'
        $r = New-MspPartnerApp -DisplayName 'Contoso MSP Worker' -PartnerTenantId $partner -ManifestPath $manifestPath -Confirm:$false
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall).Count | Should -Be 0
    }
    It 'stops when not connected' {
        Reset-AppFake -Connected $false
        (New-MspPartnerApp -DisplayName 'X' -PartnerTenantId $partner -ManifestPath $manifestPath -Confirm:$false).Steps[0].Detail | Should -Match 'Connect-MgGraph'
    }
    It 'stops when a manifest permission does not match the published permission' {
        Reset-AppFake -GraphScopeValue 'User.ReadBasic.All'
        $r = New-MspPartnerApp -DisplayName 'X' -PartnerTenantId $partner -ManifestPath $manifestPath -Confirm:$false
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'with -UseExisting adds manifest permissions without removing existing ones, and keeps other web settings' {
        $existing = [pscustomobject]@{ id = 'app-obj'; appId = $newAppId; displayName = 'Contoso MSP Worker'; signInAudience = 'AzureADMultipleOrgs'
            web = [pscustomobject]@{ redirectUris = @('https://localhost:5001'); logoutUrl = 'https://contoso.example/logout'; implicitGrantSettings = [pscustomobject]@{ enableIdTokenIssuance = $false; enableAccessTokenIssuance = $false } }
            requiredResourceAccess = @([pscustomobject]@{ resourceAppId = '00000003-0000-0000-c000-000000000000'; resourceAccess = @([pscustomobject]@{ id = 'df021288-bdef-4463-88db-98f22de89214'; type = 'Role' }) })
            keyCredentials = @() }
        Reset-AppFake -ExistingApps @($existing)
        $global:MspTest.State.Created = $existing
        $global:MspTest.State.SpCreated = $true
        $script:patches = New-Object System.Collections.Generic.List[object]
        $global:MspTest.Routes = @($global:MspTest.Routes) + (New-FakeRoute -Method PATCH -Pattern 'applications/app-obj$' -Response {
                param($Body)
                $script:patches.Add($Body)
                $current = $global:MspTest.State.Created
                $global:MspTest.State.Created = [pscustomobject]@{ id = 'app-obj'; appId = $current.appId; signInAudience = $current.signInAudience
                    web = if ($Body.web) { [pscustomobject]$Body.web } else { $current.web }
                    requiredResourceAccess = if ($Body.requiredResourceAccess) { $Body.requiredResourceAccess } else { $current.requiredResourceAccess } }
                $null
            })
        $r = New-MspPartnerApp -DisplayName 'Contoso MSP Worker' -PartnerTenantId $partner -ManifestPath $manifestPath -UseExisting -SkipAdminConsent -Confirm:$false
        $r.Success | Should -BeTrue
        $patch = $script:patches[0]
        $pairs = @(foreach ($rr in $patch.requiredResourceAccess) { foreach ($a in $rr.resourceAccess) { "$($rr.resourceAppId)|$($a.id)|$($a.type)" } })
        $pairs | Should -Contain '00000003-0000-0000-c000-000000000000|df021288-bdef-4463-88db-98f22de89214|Role'
        $pairs | Should -Contain '00000003-0000-0000-c000-000000000000|e1fe6dd8-ba31-4d61-89e7-88639da4683d|Scope'
        $pairs | Should -Contain '00000002-0000-0ff1-ce00-000000000000|ab4f2b77-0b06-4fc1-a9de-02113fc2ab7c|Scope'
        $patch.web.logoutUrl | Should -Be 'https://contoso.example/logout'
        @($patch.web.redirectUris) | Should -Be @('https://localhost:5001', 'http://localhost')
        ($r.Steps | Where-Object Step -eq 'Permissions not in manifest').Status | Should -Be 'Warning'
    }
    It 'with -UseExisting -RemoveUnlisted makes the permissions match the manifest exactly' {
        $existing = [pscustomobject]@{ id = 'app-obj'; appId = $newAppId; displayName = 'Contoso MSP Worker'; signInAudience = 'AzureADMultipleOrgs'
            web = [pscustomobject]@{ redirectUris = @('http://localhost') }
            requiredResourceAccess = @([pscustomobject]@{ resourceAppId = '00000003-0000-0000-c000-000000000000'; resourceAccess = @([pscustomobject]@{ id = 'df021288-bdef-4463-88db-98f22de89214'; type = 'Role' }) })
            keyCredentials = @() }
        Reset-AppFake -ExistingApps @($existing)
        $global:MspTest.State.Created = $existing
        $global:MspTest.State.SpCreated = $true
        $script:patches = New-Object System.Collections.Generic.List[object]
        $global:MspTest.Routes = @($global:MspTest.Routes) + (New-FakeRoute -Method PATCH -Pattern 'applications/app-obj$' -Response {
                param($Body)
                $script:patches.Add($Body)
                $current = $global:MspTest.State.Created
                $global:MspTest.State.Created = [pscustomobject]@{ id = 'app-obj'; appId = $current.appId; signInAudience = $current.signInAudience; web = $current.web; requiredResourceAccess = $Body.requiredResourceAccess }
                $null
            })
        $r = New-MspPartnerApp -DisplayName 'Contoso MSP Worker' -PartnerTenantId $partner -ManifestPath $manifestPath -UseExisting -RemoveUnlisted -SkipAdminConsent -Confirm:$false
        $r.Success | Should -BeTrue
        $pairs = @(foreach ($rr in $script:patches[0].requiredResourceAccess) { foreach ($a in $rr.resourceAccess) { "$($a.id)|$($a.type)" } })
        $pairs | Should -Not -Contain 'df021288-bdef-4463-88db-98f22de89214|Role'
        $pairs.Count | Should -Be 3
        $script:patches[0].PSObject.Properties['web'] | Should -BeNullOrEmpty
    }
    It 'refuses to create a duplicate name' {
        Reset-AppFake -ExistingApps @([pscustomobject]@{ id = 'other'; appId = '66666666-6666-6666-6666-666666666666'; displayName = 'X' })
        $r = New-MspPartnerApp -DisplayName 'X' -PartnerTenantId $partner -ManifestPath $manifestPath -Confirm:$false
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
    It 'writes nothing with -WhatIf' {
        Reset-AppFake
        $r = New-MspPartnerApp -DisplayName 'X' -PartnerTenantId $partner -ManifestPath $manifestPath -WhatIf
        $r.Outcome | Should -Be 'WhatIf'
        @(Get-FakeGraphCall -Method POST).Count | Should -Be 0
    }
}

Describe 'Add-MspPartnerAppCertificate' {
    BeforeEach {
        $script:existing = New-TestCertificate -Subject 'CN=Existing'
        $global:MspTest.State = @{ Keys = @([pscustomobject]@{ type = 'AsymmetricX509Cert'; customKeyIdentifier = (ConvertTo-TestKeyIdentifier $script:existing.Thumbprint); endDateTime = [datetime]::UtcNow.AddDays(100) }) }
        $ctx = [pscustomobject]@{ TenantId = $partner; Account = 'admin'; Scopes = @('Application.ReadWrite.All') }
        Mock Get-MgContext { $ctx }.GetNewClosure()
        Set-FakeGraphRoute @(
            New-FakeRoute -Collection -Pattern "applications\?\`$filter=appId eq '$newAppId'" -Response { [pscustomobject]@{ id = 'aaaaaaaa-0000-0000-0000-000000000001'; appId = $newAppId; displayName = 'Worker'; keyCredentials = $global:MspTest.State.Keys } }
            New-FakeRoute -Pattern 'applications/aaaaaaaa-0000-0000-0000-000000000001\?' -Response { [pscustomobject]@{ id = 'aaaaaaaa-0000-0000-0000-000000000001'; keyCredentials = $global:MspTest.State.Keys } }
            New-FakeRoute -Method PATCH -Pattern 'applications/aaaaaaaa-0000-0000-0000-000000000001$' -Response {
                param($Body)
                $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (, [Convert]::FromBase64String($Body.keyCredentials[0].key))
                $global:MspTest.State.Keys = @([pscustomobject]@{ type = 'AsymmetricX509Cert'; customKeyIdentifier = (ConvertTo-TestKeyIdentifier $c.Thumbprint); endDateTime = $c.NotAfter })
            }
            New-FakeRoute -Method POST -Pattern 'applications/aaaaaaaa-0000-0000-0000-000000000001/addKey$' -Response {
                param($Body)
                $c = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (, [Convert]::FromBase64String($Body.keyCredential.key))
                $global:MspTest.State.Keys = @($global:MspTest.State.Keys) + [pscustomobject]@{ type = 'AsymmetricX509Cert'; customKeyIdentifier = (ConvertTo-TestKeyIdentifier $c.Thumbprint); endDateTime = $c.NotAfter }
                [pscustomobject]@{}
            }
        )
    }
    It 'needs a proof certificate when a valid certificate exists, and writes nothing without it' {
        $r = Add-MspPartnerAppCertificate -AppId $newAppId -PartnerTenantId $partner -Certificate (New-TestCertificate) -Confirm:$false
        $r.Success | Should -BeFalse
        @(Get-FakeGraphCall | Where-Object Method -ne 'GET').Count | Should -Be 0
    }
    It 'adds with addKey and a proof signed by the existing certificate, keeping the old key' {
        $new = New-TestCertificate -Subject 'CN=New'
        $r = Add-MspPartnerAppCertificate -AppId $newAppId -PartnerTenantId $partner -Certificate $new -ProofCertificate $script:existing -Confirm:$false
        $r.Success | Should -BeTrue
        $r.Method | Should -Be 'addKey'
        @($global:MspTest.State.Keys).Count | Should -Be 2
        (Get-FakeGraphCall -Method POST -Pattern 'addKey')[0].Body.proof.Split('.').Count | Should -Be 3
        @(Get-FakeGraphCall -Method PATCH).Count | Should -Be 0
    }
    It 'uses PATCH when no valid certificate exists' {
        $global:MspTest.State.Keys = @()
        $r = Add-MspPartnerAppCertificate -AppId $newAppId -PartnerTenantId $partner -Certificate (New-TestCertificate) -Confirm:$false
        $r.Success | Should -BeTrue
        $r.Method | Should -Be 'PATCH keyCredentials'
    }
    It 'does nothing when the certificate is already present' {
        $r = Add-MspPartnerAppCertificate -AppId $newAppId -PartnerTenantId $partner -Certificate $script:existing -Confirm:$false
        $r.Success | Should -BeTrue
        @(Get-FakeGraphCall | Where-Object Method -ne 'GET').Count | Should -Be 0
    }
    It 'recognises a present certificate when Graph reports customKeyIdentifier as the hex thumbprint' {
        $global:MspTest.State.Keys = @([pscustomobject]@{ type = 'AsymmetricX509Cert'; customKeyIdentifier = $script:existing.Thumbprint; endDateTime = [datetime]::UtcNow.AddDays(100) })
        $r = Add-MspPartnerAppCertificate -AppId $newAppId -PartnerTenantId $partner -Certificate $script:existing -Confirm:$false
        $r.Success | Should -BeTrue
        @(Get-FakeGraphCall | Where-Object Method -ne 'GET').Count | Should -Be 0
    }
}
