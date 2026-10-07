#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

# Session, configuration and transport behaviour that needs the real module scope.
# Offline only: every HTTP call, vault call and Microsoft module cmdlet is mocked or stubbed.

BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers' 'CoreTestHelpers.ps1')
    Import-Module $script:ModuleManifestPath -Force

    # Stubs for the optional Microsoft modules, defined inside the module so Mock can bind to them
    # whether or not the real modules are installed.
    InModuleScope MspGdap {
        function script:Connect-ExchangeOnline { [CmdletBinding()] param([string]$AccessToken, [string]$DelegatedOrganization, [string]$Organization, [string]$AppId, [string]$CertificateThumbprint, [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate, [switch]$ShowBanner, [string]$Prefix, [string[]]$CommandName) }
        function script:Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param([string]$ConnectionId) }
        function script:Get-ConnectionInformation { [CmdletBinding()] param() }
        function script:Connect-IPPSSession { [CmdletBinding()] param([string]$AccessToken, [string]$Organization, [string]$DelegatedOrganization, [string]$AzureADAuthorizationEndpointUri, [switch]$ShowBanner, [string]$Prefix) }
        function script:Connect-MgGraph { [CmdletBinding()] param([securestring]$AccessToken, [switch]$NoWelcome) }
        function script:Disconnect-MgGraph { [CmdletBinding()] param() }
        function script:Get-MgContext { [CmdletBinding()] param() }
        function script:Connect-MicrosoftTeams { [CmdletBinding()] param([string[]]$AccessTokens) }
        function script:Disconnect-MicrosoftTeams { [CmdletBinding()] param() }
        function script:Remove-Secret { [CmdletBinding()] param($Name, $Vault) }
    }
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'Disconnect-Msp closes the sessions Connect-Msp* opened' {
    BeforeEach {
        Initialize-TestModuleState
        $customer = $script:TestCustomerTenantId
        Mock Assert-MspModuleAvailable -ModuleName MspGdap -MockWith { [version]'3.9.2' }
        Mock Test-MspExoCompatibility -ModuleName MspGdap -MockWith { $null }
        Mock Get-MspInitialDomain -ModuleName MspGdap -MockWith { [pscustomobject]@{ InitialDomain = 'fabrikam.onmicrosoft.com' } }
        Mock Get-MspAccessToken -ModuleName MspGdap -MockWith { [pscustomobject]@{ TenantId = $TenantId; Resource = $Resource; ExpiresOn = [datetime]::UtcNow.AddHours(1); AccessToken = (New-TestSecure "token-$Resource") } }
        $script:Sessions = @()
        Mock Get-ConnectionInformation -ModuleName MspGdap -MockWith { $script:Sessions }
        Mock Connect-ExchangeOnline -ModuleName MspGdap -MockWith { $script:Sessions = @([pscustomobject]@{ ConnectionId = 'exo'; State = 'Connected'; TenantID = $script:TestCustomerTenantId }) }
        Mock Connect-IPPSSession -ModuleName MspGdap -MockWith { $script:Sessions = @([pscustomobject]@{ ConnectionId = 'scc'; State = 'Connected'; TenantID = $script:TestCustomerTenantId; IsEopSession = $true }) }
        Mock Disconnect-ExchangeOnline -ModuleName MspGdap -MockWith {}
        Mock Connect-MgGraph -ModuleName MspGdap -MockWith {}
        Mock Get-MgContext -ModuleName MspGdap -MockWith { [pscustomobject]@{ TenantId = $script:TestCustomerTenantId; Account = 'tech'; Scopes = @() } }
        Mock Disconnect-MgGraph -ModuleName MspGdap -MockWith {}
        Mock Connect-MicrosoftTeams -ModuleName MspGdap -MockWith { [pscustomobject]@{ TenantId = $script:TestCustomerTenantId } }
        Mock Disconnect-MicrosoftTeams -ModuleName MspGdap -MockWith {}
    }

    It 'disconnects the Exchange Online session it opened, by connection ID' {
        $null = Connect-MspExchangeOnline -TenantId $customer -KeepExistingConnections
        Disconnect-Msp
        Should -Invoke Disconnect-ExchangeOnline -ModuleName MspGdap -Times 1 -Exactly
        Should -Invoke Disconnect-ExchangeOnline -ModuleName MspGdap -Times 1 -ParameterFilter { $ConnectionId -eq 'exo' }
    }
    It 'disconnects Security and Compliance by connection ID' {
        $null = Connect-MspSecurityCompliance -TenantId $customer -KeepExistingConnections -WarningAction SilentlyContinue
        Disconnect-Msp
        Should -Invoke Disconnect-ExchangeOnline -ModuleName MspGdap -Times 1 -ParameterFilter { $ConnectionId -eq 'scc' }
    }
    It 'disconnects the Microsoft Graph SDK' {
        $null = Connect-MspGraph -TenantId $customer -WarningAction SilentlyContinue
        Disconnect-Msp
        Should -Invoke Disconnect-MgGraph -ModuleName MspGdap -Times 1
    }
    It 'disconnects Microsoft Teams' {
        $null = Connect-MspTeams -TenantId $customer
        Disconnect-Msp
        Should -Invoke Disconnect-MicrosoftTeams -ModuleName MspGdap -Times 1
    }
    It 'leaves sessions alone that MspGdap did not open' {
        Disconnect-Msp
        Should -Invoke Disconnect-ExchangeOnline -ModuleName MspGdap -Times 0
        Should -Invoke Disconnect-MgGraph -ModuleName MspGdap -Times 0
        Should -Invoke Disconnect-MicrosoftTeams -ModuleName MspGdap -Times 0
    }
    It 'refuses the partner tenant for app-only Exchange connections' {
        { Connect-MspExchangeOnline -TenantId $script:TestPartnerTenantId -AppOnly -AppId $script:TestAppId -CertificateThumbprint ('B' * 40) -Organization 'contoso.onmicrosoft.com' } | Should -Throw -ErrorId 'MspGdap.Tenant.PartnerTenantNotAllowed'
        Should -Invoke Connect-ExchangeOnline -ModuleName MspGdap -Times 0
    }
}

Describe 'Set-MspConfiguration' {
    BeforeEach {
        $path = Join-Path $TestDrive "cfg-$([guid]::NewGuid()).json"
        Initialize-TestModuleState -ConfigPath $path
        InModuleScope MspGdap { $script:MspTokenCache['k'] = [pscustomobject]@{ UserPrincipalName = 'x' } }
    }

    It 'changes nothing at all under -WhatIf' {
        Set-MspConfiguration -AppId '55555555-5555-5555-5555-555555555555' -CertificatePassword (New-TestSecure 'pw') -WhatIf
        Test-Path $path | Should -BeFalse
        InModuleScope MspGdap { $script:MspTokenCache.Count } | Should -Be 1
        InModuleScope MspGdap { [bool]$script:MspState.CertificatePassword } | Should -BeFalse
        (Get-MspConfiguration).AppId | Should -Be $script:TestAppId
    }
    It 'does not switch the session to another -Path under -WhatIf' {
        $other = Join-Path $TestDrive 'other.json'
        Set-MspConfiguration -Path $other -VaultName 'Other' -WhatIf
        InModuleScope MspGdap { $script:MspState.ConfigPath } | Should -Be $path
    }
    It 'does not store a client secret when the change is declined' {
        Mock Assert-MspSecretManagement -ModuleName MspGdap -MockWith {}
        Mock Set-Secret -ModuleName MspGdap -MockWith {}
        Set-MspConfiguration -ClientSecret (New-TestSecure 'not-a-real-secret') -WhatIf -WarningAction SilentlyContinue
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0
        Test-Path $path | Should -BeFalse
    }
    It 'keeps a -Certificate object for the session only and uses it for signing' {
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=MspGdap Session Test', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $cert = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(30))
        Set-MspConfiguration -Certificate $cert -Confirm:$false
        $json = Get-Content -LiteralPath $path -Raw
        $json | Should -Not -Match $cert.Thumbprint
        $json | Should -Not -Match 'MII'
        (Get-MspConfiguration).SessionCertificate | Should -Be $cert.Thumbprint
        $used = InModuleScope MspGdap { Get-MspSigningCertificate -Configuration (Get-MspConfigurationInternal) }
        $used.Thumbprint | Should -Be $cert.Thumbprint
        $body = InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId } { param($Tenant) Get-MspClientCredentialBody -Configuration (Get-MspConfigurationInternal) -TenantId $Tenant }
        $body.client_assertion_type | Should -Be 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $body.client_assertion.Split('.').Count | Should -Be 3
        Disconnect-Msp
        InModuleScope MspGdap { $script:MspState.SessionCertificate } | Should -BeNullOrEmpty
    }
    It 'refuses a -Certificate without a private key' {
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=MspGdap Public Only', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $full = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(30))
        $public = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($full.RawData)
        { Set-MspConfiguration -Certificate $public -Confirm:$false } | Should -Throw -ErrorId 'MspGdap.Certificate.NoPrivateKey*'
    }
}

Describe 'Signing certificate loading and client credentials' {
    BeforeEach { Initialize-TestModuleState }

    It 'loads a PFX from -CertificatePath with the session password and signs an assertion' {
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=MspGdap Pfx Test', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $cert = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(30))
        $pfxPath = Join-Path $TestDrive 'test.pfx'
        [System.IO.File]::WriteAllBytes($pfxPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, 'test-password'))
        InModuleScope MspGdap -Parameters @{ PfxPath = $pfxPath } {
            param($PfxPath)
            $script:MspState.Config.CertificateThumbprint = $null
            $script:MspState.Config.CertificatePath = $PfxPath
            $script:MspState.CertificatePassword = [System.Net.NetworkCredential]::new('', 'test-password').SecurePassword
        }
        $loaded = InModuleScope MspGdap { Get-MspSigningCertificate -Configuration (Get-MspConfigurationInternal) }
        $loaded.Thumbprint | Should -Be $cert.Thumbprint
        $loaded.HasPrivateKey | Should -BeTrue
        $again = InModuleScope MspGdap { Get-MspSigningCertificate -Configuration (Get-MspConfigurationInternal) }
        [object]::ReferenceEquals($loaded, $again) | Should -BeTrue
        $assertion = InModuleScope MspGdap -Parameters @{ Cert = $loaded; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
            param($Cert, $App, $Tenant)
            New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant -Algorithm PS256
        }
        $parts = $assertion.Split('.')
        $parts.Count | Should -Be 3
        $claims = InModuleScope MspGdap -Parameters @{ Token = $assertion } { param($Token) (ConvertFrom-MspJwt -Token $Token).Claims }
        $claims.aud | Should -Be "https://login.microsoftonline.com/$($script:TestCustomerTenantId)/oauth2/v2.0/token"
        $claims.iss | Should -Be $script:TestAppId
    }
    It 'reports a wrong PFX password without leaking it' {
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=MspGdap Pfx Bad', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $cert = $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays(-1), [DateTimeOffset]::UtcNow.AddDays(30))
        $pfxPath = Join-Path $TestDrive 'bad.pfx'
        [System.IO.File]::WriteAllBytes($pfxPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, 'right-password'))
        InModuleScope MspGdap -Parameters @{ PfxPath = $pfxPath } {
            param($PfxPath)
            $script:MspState.Config.CertificateThumbprint = $null
            $script:MspState.Config.CertificatePath = $PfxPath
            $script:MspState.CertificatePassword = [System.Net.NetworkCredential]::new('', 'wrong-password').SecurePassword
        }
        $err = $null
        try { InModuleScope MspGdap { Get-MspSigningCertificate -Configuration (Get-MspConfigurationInternal) } } catch { $err = $_ }
        $err.FullyQualifiedErrorId | Should -BeLike 'MspGdap.Certificate.LoadFailed*'
        $err.Exception.Message | Should -Not -Match 'wrong-password|right-password'
    }
    It 'fails clearly when the configured certificate is missing' {
        InModuleScope MspGdap { $script:MspState.Config.CertificateThumbprint = $null; $script:MspState.Config.CertificatePath = 'Z:\no\such\file.pfx' }
        { InModuleScope MspGdap { Get-MspSigningCertificate -Configuration (Get-MspConfigurationInternal) } } | Should -Throw -ErrorId 'MspGdap.Certificate.NotFound*'
    }
    It 'reads a client secret from the vault for the client-secret credential' {
        Mock Get-Secret -ModuleName MspGdap -MockWith { New-TestSecure 'not-a-real-secret' }
        InModuleScope MspGdap { $script:MspState.Config.CredentialType = 'ClientSecret' }
        $body = InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId } { param($Tenant) Get-MspClientCredentialBody -Configuration (Get-MspConfigurationInternal) -TenantId $Tenant }
        $body.client_secret | Should -Be 'not-a-real-secret'
        Should -Invoke Get-Secret -ModuleName MspGdap -Times 1 -ParameterFilter { $Vault -eq 'TestVault' -and $Name -like 'MspGdap-*' }
    }
    It 'fails clearly when the client secret is not in the vault' {
        Mock Get-Secret -ModuleName MspGdap -MockWith { throw 'not found' }
        InModuleScope MspGdap { $script:MspState.Config.CredentialType = 'ClientSecret' }
        { InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId } { param($Tenant) Get-MspClientCredentialBody -Configuration (Get-MspConfigurationInternal) -TenantId $Tenant } } | Should -Throw -ErrorId 'MspGdap.ClientSecret.NotFound*'
    }
}

Describe 'Tenant parameters accept a verified domain' {
    BeforeEach {
        Initialize-TestModuleState
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/fabrikam.onmicrosoft.com/v2.0/.well-known/openid-configuration' } -MockWith {
            [pscustomobject]@{ issuer = "https://login.microsoftonline.com/$($script:TestCustomerTenantId)/v2.0" }
        }
    }
    It 'Clear-MspTokenCache -TenantId <domain>' {
        InModuleScope MspGdap -Parameters @{ Customer = $script:TestCustomerTenantId; Partner = $script:TestPartnerTenantId } {
            param($Customer, $Partner)
            $script:MspTokenCache["$Customer|https://graph.microsoft.com|app"] = [pscustomobject]@{}
            $script:MspTokenCache["$Partner|https://graph.microsoft.com|app"] = [pscustomobject]@{}
        }
        Clear-MspTokenCache -TenantId 'fabrikam.onmicrosoft.com'
        InModuleScope MspGdap { @($script:MspTokenCache.Keys) } | Should -Be @("$($script:TestPartnerTenantId)|https://graph.microsoft.com|app")
    }
    It 'Test-MspAccessToken -TenantId <domain>' {
        $token = [pscustomobject]@{ TenantId = $script:TestCustomerTenantId; Resource = 'https://graph.microsoft.com'; ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1); AccessToken = (New-TestSecure 'opaque') }
        Test-MspAccessToken -InputObject $token -TenantId 'fabrikam.onmicrosoft.com' -Resource Graph | Should -BeTrue
    }
    It 'Get-MspCustomer -TenantId <domain>' {
        Mock Invoke-MspGraphRequest -ModuleName MspGdap -MockWith {
            [pscustomobject]@{ customerId = $script:TestCustomerTenantId; displayName = 'Fabrikam'; defaultDomainName = 'fabrikam.com' }
            [pscustomobject]@{ customerId = $script:TestOtherTenantId; displayName = 'Other'; defaultDomainName = 'other.example' }
        }
        $r = @(Get-MspCustomer -TenantId 'fabrikam.onmicrosoft.com')
        $r.Count | Should -Be 1
        $r[0].DisplayName | Should -Be 'Fabrikam'
    }
}

Describe 'Resolve-MspTenantId' {
    BeforeEach { Initialize-TestModuleState }
    It 'rejects the special name <Name> with a clear message' -ForEach @(@{ Name = 'common' }, @{ Name = 'organizations' }, @{ Name = 'consumers' }) {
        $err = $null
        try { Resolve-MspTenantId -Tenant $Name } catch { $err = $_ }
        $err.Exception.Message | Should -Match 'not a specific tenant'
    }
    It 'tells a network failure apart from an unknown domain' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { throw [System.Net.Http.HttpRequestException]::new('No such host is known.') }
        { Resolve-MspTenantId -Tenant 'contoso-offline.example' } | Should -Throw -ErrorId 'MspGdap.Tenant.LookupFailed*'
    }
    It 'reports an unknown domain as not found' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { throw (New-TestHttpError -Status 400 -Json '{"error":"invalid_tenant"}') }
        { Resolve-MspTenantId -Tenant 'no-such-tenant.example' } | Should -Throw -ErrorId 'MspGdap.Tenant.NotFound*'
    }
}

Describe 'Invoke-MspRestWithRetry' {
    BeforeEach {
        Mock Start-Sleep -ModuleName MspGdap -MockWith {}
        $script:Calls = 0
    }
    It 'does not retry a POST that returned 503, because it may have been processed' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { $script:Calls++; throw (New-TestHttpError -Status 503) }
        { InModuleScope MspGdap { Invoke-MspRestWithRetry -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Method POST -Body @{ appId = 'x' } -ServiceName Graph } } | Should -Throw -ErrorId 'MspGdap.Graph.Http503*'
        $script:Calls | Should -Be 1
    }
    It 'retries a GET that returned 503' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { $script:Calls++; if ($script:Calls -lt 3) { throw (New-TestHttpError -Status 503) } else { [pscustomobject]@{ ok = $true } } }
        (InModuleScope MspGdap { Invoke-MspRestWithRetry -Uri 'https://graph.microsoft.com/v1.0/organization' -Method GET -ServiceName Graph }).ok | Should -BeTrue
        $script:Calls | Should -Be 3
    }
    It 'retries a POST that was throttled with 429' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { $script:Calls++; if ($script:Calls -lt 2) { throw (New-TestHttpError -Status 429 -RetryAfterSeconds 1) } else { [pscustomobject]@{ id = 'n' } } }
        (InModuleScope MspGdap { Invoke-MspRestWithRetry -Uri 'https://graph.microsoft.com/v1.0/groups' -Method POST -Body @{ x = 1 } -ServiceName Graph }).id | Should -Be 'n'
        $script:Calls | Should -Be 2
    }
    It 'keeps a one-element array body as an array' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { $script:SentBody = $Body; [pscustomobject]@{} }
        $null = InModuleScope MspGdap { Invoke-MspRestWithRetry -Uri 'https://graph.microsoft.com/v1.0/x' -Method POST -Body @(@{ id = 1 }) -ServiceName Graph }
        $script:SentBody | Should -Be '[{"id":1}]'
    }
}

Describe 'Private transport wrappers' {
    BeforeEach {
        Initialize-TestModuleState
        Mock Start-Sleep -ModuleName MspGdap -MockWith {}
        Mock Get-MspAuthHeader -ModuleName MspGdap -MockWith { @{ Authorization = 'Bearer real-token' } }
    }
    It 'Invoke-MspGraphDirect refuses a non-Graph host before asking for a token' {
        { InModuleScope MspGdap { Invoke-MspGraphDirect -PartnerTenant -Method GET -Path 'https://evil.example/v1.0/me' } } | Should -Throw -ErrorId 'MspGdap.Graph.ForeignHost*'
        Should -Invoke Get-MspAuthHeader -ModuleName MspGdap -Times 0
    }
    It 'Invoke-MspGraphDirect never lets -Headers replace Authorization' {
        Mock Invoke-WebRequest -ModuleName MspGdap -MockWith { $script:SentHeaders = $Headers; [pscustomobject]@{ StatusCode = 200; Content = '{}'; Headers = @{} } }
        $null = InModuleScope MspGdap { Invoke-MspGraphDirect -PartnerTenant -Method PATCH -Path 'tenantRelationships/x' -Headers @{ Authorization = 'Bearer attacker'; 'If-Match' = 'W/"1"' } -Body @{ a = 1 } }
        $script:SentHeaders.Authorization | Should -Be 'Bearer real-token'
        $script:SentHeaders['If-Match'] | Should -Be 'W/"1"'
    }
    It 'Invoke-MspGraphCall refuses a non-Graph host' {
        { InModuleScope MspGdap { Invoke-MspGraphCall -TenantId '22222222-2222-4222-8222-222222222222' -Path 'https://evil.example/v1.0/users' } } | Should -Throw -ErrorId 'MspGdap.Graph.ForeignHost*'
    }
    It 'Invoke-MspPartnerCenterRequest reuses one ms-requestid across retries' {
        $script:RequestIds = New-Object System.Collections.Generic.List[string]
        Mock Invoke-WebRequest -ModuleName MspGdap -MockWith {
            $script:RequestIds.Add($Headers['ms-requestid'])
            if ($script:RequestIds.Count -lt 3) { throw (New-TestHttpError -Status 503) }
            [pscustomobject]@{ StatusCode = 201; Content = '{}'; Headers = @{} }
        }
        $r = InModuleScope MspGdap { Invoke-MspPartnerCenterRequest -Method POST -Path 'customers/x/applicationconsents' -Body @{ a = 1 } }
        $r.StatusCode | Should -Be 201
        $script:RequestIds.Count | Should -Be 3
        @($script:RequestIds | Select-Object -Unique).Count | Should -Be 1
        $script:RequestIds[0] | Should -Match '^[0-9a-f-]{36}$'
    }
}

Describe 'Invoke-MspGraphBatch' {
    BeforeEach {
        Initialize-TestModuleState
        Mock Start-Sleep -ModuleName MspGdap -MockWith {}
        $script:Refreshes = 0
        Mock Get-MspAuthHeader -ModuleName MspGdap -MockWith { if ($ForceRefresh) { $script:Refreshes++ }; @{ Authorization = 'Bearer t' } }
        Mock Invoke-MspRestWithRetry -ModuleName MspGdap -MockWith {
            $script:SentBatch = $Body.requests
            [pscustomobject]@{ responses = @($Body.requests | ForEach-Object { [pscustomobject]@{ id = $_.id; status = 200; body = $null } }) }
        }
    }
    It 'refuses duplicate request IDs' {
        { Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request @(@{ id = 'a'; url = '/me' }, @{ id = 'a'; url = '/users' }) } | Should -Throw -ErrorId 'MspGdap.GraphBatch.DuplicateId*'
    }
    It 'generates IDs that never collide with IDs the caller set' {
        $null = Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request @(@{ url = '/me' }, @{ id = '1'; url = '/users' })
        @($script:SentBatch.id | Select-Object -Unique).Count | Should -Be 2
    }
    It 'passes dependsOn through' {
        $null = Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request @(@{ id = 'a'; method = 'POST'; url = '/groups'; body = @{ x = 1 } }, @{ id = 'b'; method = 'GET'; url = '/me'; dependsOn = @('a') }) -Confirm:$false
        ($script:SentBatch | Where-Object id -eq 'b').dependsOn | Should -Be @('a')
    }
    It 'refuses a dependsOn that points outside its batch of 20' {
        $requests = @(1..21 | ForEach-Object { @{ id = "r$_"; url = '/me' } })
        $requests[20]['dependsOn'] = @('r1')
        { Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request $requests } | Should -Throw -ErrorId 'MspGdap.GraphBatch.InvalidDependency*'
    }
    It 'refreshes the token once when the batch call returns 401' {
        $script:BatchCalls = 0
        Mock Invoke-MspRestWithRetry -ModuleName MspGdap -MockWith {
            $script:BatchCalls++
            if ($script:BatchCalls -eq 1) { $e = [System.InvalidOperationException]::new('401'); $e.Data['StatusCode'] = 401; throw $e }
            [pscustomobject]@{ responses = @([pscustomobject]@{ id = '1'; status = 200; body = $null }) }
        }
        $r = @(Invoke-MspGraphBatch -TenantId $script:TestCustomerTenantId -Request @(@{ url = '/me' }))
        $r[0].Success | Should -BeTrue
        $script:Refreshes | Should -Be 1
    }
}

Describe 'Unregister-MspPartnerToken' {
    BeforeEach {
        Initialize-TestModuleState
        Mock Assert-MspSecretManagement -ModuleName MspGdap -MockWith {}
        Mock Get-SecretInfo -ModuleName MspGdap -MockWith { [pscustomobject]@{ Name = $Name } }
        Mock Remove-Secret -ModuleName MspGdap -MockWith {}
        InModuleScope MspGdap -Parameters @{ Upn = $script:TestUpn } {
            param($Upn)
            $script:MspTokenCache['a|g|x'] = [pscustomobject]@{ UserPrincipalName = $Upn }
            $script:MspTokenCache['b|g|x'] = [pscustomobject]@{ UserPrincipalName = 'other@contoso.onmicrosoft.com' }
            $script:MspState.CurrentUpn = $Upn
        }
    }
    It 'removes the vault secret and only that technician''s cached tokens' {
        $r = Unregister-MspPartnerToken -UserPrincipalName $script:TestUpn -Confirm:$false -WarningAction SilentlyContinue
        $r.RefreshTokenRemoved | Should -BeTrue
        $r.CachedTokensCleared | Should -Be 1
        Should -Invoke Remove-Secret -ModuleName MspGdap -Times 1 -ParameterFilter { $Vault -eq 'TestVault' -and $Name -like "MspGdap-$($script:TestAppId)-*" }
        InModuleScope MspGdap { @($script:MspTokenCache.Keys) } | Should -Be @('b|g|x')
        InModuleScope MspGdap { $script:MspState.CurrentUpn } | Should -BeNullOrEmpty
    }
    It 'changes nothing with -WhatIf' {
        Unregister-MspPartnerToken -UserPrincipalName $script:TestUpn -WhatIf
        Should -Invoke Remove-Secret -ModuleName MspGdap -Times 0
        InModuleScope MspGdap { $script:MspTokenCache.Count } | Should -Be 2
    }
    It 'warns when there is nothing stored' {
        Mock Get-SecretInfo -ModuleName MspGdap -MockWith { $null }
        $r = Unregister-MspPartnerToken -UserPrincipalName $script:TestUpn -Confirm:$false -WarningVariable w -WarningAction SilentlyContinue
        $r.RefreshTokenRemoved | Should -BeFalse
        ($w | Out-String) | Should -Match 'No stored refresh token'
        Should -Invoke Remove-Secret -ModuleName MspGdap -Times 0
    }
}

Describe 'Optional module checks' {
    It 'Assert-MspModuleAvailable fails clearly for a missing module' {
        { InModuleScope MspGdap { Assert-MspModuleAvailable -Name 'MspGdap.NoSuchModule' -MinimumVersion '1.0.0' } } | Should -Throw '*MspGdap.NoSuchModule*'
    }
    It 'Assert-MspModuleAvailable returns the version of an available module' {
        $version = InModuleScope MspGdap { Assert-MspModuleAvailable -Name 'Pester' -MinimumVersion '5.0.0' }
        [version]$version | Should -BeGreaterOrEqual ([version]'5.0.0')
    }
}
