#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers' 'CoreTestHelpers.ps1')
    Import-Module $script:ModuleManifestPath -Force

    $script:NewTestCertificate = {
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=MspGdap Test', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $request.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-5), [DateTimeOffset]::UtcNow.AddDays(30))
    }
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'New-MspClientAssertion' {
    BeforeAll {
        $script:Cert = & $script:NewTestCertificate
    }

    It 'builds a PS256 assertion with x5t#S256 and the Learn claim set' {
        $assertion = InModuleScope MspGdap -Parameters @{ Cert = $script:Cert; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
            param($Cert, $App, $Tenant)
            New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant
        }
        $parts = $assertion.Split('.')
        $parts.Count | Should -Be 3

        $decoded = InModuleScope MspGdap -Parameters @{ Token = $assertion } { param($Token) ConvertFrom-MspJwt -Token $Token }
        $decoded.Header.alg | Should -Be 'PS256'
        $decoded.Header.typ | Should -Be 'JWT'
        $expectedThumb = ConvertTo-TestBase64Url ([System.Security.Cryptography.SHA256]::HashData($script:Cert.RawData))
        $decoded.Header.'x5t#S256' | Should -Be $expectedThumb
        $decoded.Header.PSObject.Properties['x5t'] | Should -BeNullOrEmpty

        $claims = $decoded.Claims
        $claims.aud | Should -Be "https://login.microsoftonline.com/$($script:TestCustomerTenantId)/oauth2/v2.0/token"
        $claims.iss | Should -Be $script:TestAppId
        $claims.sub | Should -Be $script:TestAppId
        ($claims.exp - $claims.nbf) | Should -Be 300
        { [guid]::Parse($claims.jti) } | Should -Not -Throw
    }

    It 'produces a PSS signature that verifies with the certificate public key' {
        $assertion = InModuleScope MspGdap -Parameters @{ Cert = $script:Cert; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
            param($Cert, $App, $Tenant)
            New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant
        }
        $parts = $assertion.Split('.')
        $data = [Text.Encoding]::ASCII.GetBytes("$($parts[0]).$($parts[1])")
        $signature = InModuleScope MspGdap -Parameters @{ Value = $parts[2] } { param($Value) ConvertFrom-MspBase64Url -Value $Value }
        $public = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($script:Cert)
        $public.VerifyData($data, $signature, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pss) | Should -BeTrue
    }

    It 'uses a unique jti for every assertion' {
        $jtis = 1..3 | ForEach-Object {
            $a = InModuleScope MspGdap -Parameters @{ Cert = $script:Cert; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
                param($Cert, $App, $Tenant)
                New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant
            }
            (InModuleScope MspGdap -Parameters @{ Token = $a } { param($Token) ConvertFrom-MspJwt -Token $Token }).Claims.jti
        }
        @($jtis | Select-Object -Unique).Count | Should -Be 3
    }

    It 'signs RS256 with PKCS#1 and adds the legacy x5t header when asked' {
        $assertion = InModuleScope MspGdap -Parameters @{ Cert = $script:Cert; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
            param($Cert, $App, $Tenant)
            New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant -Algorithm RS256
        }
        $parts = $assertion.Split('.')
        $decoded = InModuleScope MspGdap -Parameters @{ Token = $assertion } { param($Token) ConvertFrom-MspJwt -Token $Token }
        $decoded.Header.alg | Should -Be 'RS256'
        $decoded.Header.x5t | Should -Be (ConvertTo-TestBase64Url $script:Cert.GetCertHash())
        $data = [Text.Encoding]::ASCII.GetBytes("$($parts[0]).$($parts[1])")
        $signature = InModuleScope MspGdap -Parameters @{ Value = $parts[2] } { param($Value) ConvertFrom-MspBase64Url -Value $Value }
        $public = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($script:Cert)
        $public.VerifyData($data, $signature, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1) | Should -BeTrue
    }

    It 'falls back to RS256 when the key provider cannot sign with PSS (legacy CSP keys)' {
        InModuleScope MspGdap -Parameters @{ Cert = $script:Cert; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
            param($Cert, $App, $Tenant)
            Mock Invoke-MspRsaSignature -MockWith {
                , $Rsa.SignData($Data, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            }
            Mock Invoke-MspRsaSignature -ParameterFilter { $Algorithm -eq 'PS256' } -MockWith {
                throw [System.Security.Cryptography.CryptographicException]::new('Specified padding mode is not valid for this algorithm.')
            }
            $assertion = New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant
            (ConvertFrom-MspJwt -Token $assertion).Header.alg | Should -Be 'RS256'
        }
    }

    It 'refuses a certificate without a private key' {
        $publicOnly = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($script:Cert.RawData)
        {
            InModuleScope MspGdap -Parameters @{ Cert = $publicOnly; App = $script:TestAppId; Tenant = $script:TestCustomerTenantId } {
                param($Cert, $App, $Tenant)
                New-MspClientAssertion -Certificate $Cert -ClientId $App -TenantId $Tenant
            }
        } | Should -Throw -ErrorId 'MspGdap.Certificate.NoPrivateKey'
    }
}

Describe 'ConvertFrom-MspJwt' {
    It 'decodes header and claims without padding' {
        $token = New-TestJwt -Claims @{ tid = $script:TestCustomerTenantId; aud = 'https://graph.microsoft.com'; n = 'x' }
        $decoded = InModuleScope MspGdap -Parameters @{ Token = $token } { param($Token) ConvertFrom-MspJwt -Token $Token }
        $decoded.Claims.tid | Should -Be $script:TestCustomerTenantId
        $decoded.Header.alg | Should -Be 'RS256'
    }

    It 'accepts a SecureString' {
        $token = New-TestSecure (New-TestJwt -Claims @{ tid = 'abc' })
        $decoded = InModuleScope MspGdap -Parameters @{ Token = $token } { param($Token) ConvertFrom-MspJwt -Token $Token }
        $decoded.Claims.tid | Should -Be 'abc'
    }

    It 'returns null for opaque or malformed tokens instead of throwing' -ForEach @(
        @{ Token = 'opaque-token-value' }
        @{ Token = 'not.base64!.x' }
        # Built at run time so secret scanners do not mistake a fake token for a real one.
        @{ Token = ('e30', 'bm90IGpzb24', 'x') -join '.' }
        @{ Token = '' }
    ) {
        $result = InModuleScope MspGdap -Parameters @{ Token = $Token } { param($Token) ConvertFrom-MspJwt -Token $Token }
        $result | Should -BeNullOrEmpty
    }
}

Describe 'PKCE, state and secret naming' {
    It 'creates an S256 challenge that matches the verifier' {
        $pkce = InModuleScope MspGdap { New-MspPkceChallenge }
        $pkce.Method | Should -Be 'S256'
        $pkce.Verifier.Length | Should -Be 43
        $pkce.Verifier | Should -Match '^[A-Za-z0-9_-]+$'
        $expected = ConvertTo-TestBase64Url ([System.Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes($pkce.Verifier)))
        $pkce.Challenge | Should -Be $expected
    }

    It 'names refresh token secrets without the UPN and within Key Vault rules' {
        $name = InModuleScope MspGdap -Parameters @{ Upn = $script:TestUpn; App = $script:TestAppId } {
            param($Upn, $App)
            Get-MspRefreshTokenSecretName -UserPrincipalName $Upn -AppId $App
        }
        $name | Should -Match "^MspGdap-$($script:TestAppId)-[0-9a-f]{16}$"
        $name.Length | Should -BeLessOrEqual 127
        $name | Should -Not -Match 'contoso|tech-admin|@|\.'
    }

    It 'gives the same secret name regardless of UPN case and whitespace' {
        $names = InModuleScope MspGdap -Parameters @{ App = $script:TestAppId } {
            param($App)
            Get-MspRefreshTokenSecretName -UserPrincipalName 'Tech-Admin@Contoso.onmicrosoft.com ' -AppId $App
            Get-MspRefreshTokenSecretName -UserPrincipalName 'tech-admin@contoso.onmicrosoft.com' -AppId $App
        }
        $names[0] | Should -Be $names[1]
    }
}

Describe 'Get-MspTokenCacheKey' {
    It 'builds tenantId|resource|appId in lower case' {
        $key = InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId.ToUpper(); App = $script:TestAppId.ToUpper() } {
            param($Tenant, $App)
            Get-MspTokenCacheKey -TenantId $Tenant -Resource 'https://Graph.Microsoft.com/' -AppId $App
        }
        $key | Should -Be "$($script:TestCustomerTenantId)|https://graph.microsoft.com|$($script:TestAppId)"
    }

    It 'normalises aliases, trailing slashes, /.default and app IDs to one key' -ForEach @(
        @{ Resource = 'Graph' }
        @{ Resource = 'https://graph.microsoft.com' }
        @{ Resource = 'https://graph.microsoft.com/' }
        @{ Resource = 'https://graph.microsoft.com/.default' }
        @{ Resource = '00000003-0000-0000-c000-000000000000' }
    ) {
        $key = InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId; App = $script:TestAppId; Resource = $Resource } {
            param($Tenant, $App, $Resource)
            Get-MspTokenCacheKey -TenantId $Tenant -Resource $Resource -AppId $App
        }
        $key | Should -Be "$($script:TestCustomerTenantId)|https://graph.microsoft.com|$($script:TestAppId)"
    }

    It 'keeps tenants and resources apart' {
        $keys = InModuleScope MspGdap -Parameters @{ A = $script:TestCustomerTenantId; B = $script:TestOtherTenantId; App = $script:TestAppId } {
            param($A, $B, $App)
            Get-MspTokenCacheKey -TenantId $A -Resource Graph -AppId $App
            Get-MspTokenCacheKey -TenantId $B -Resource Graph -AppId $App
            Get-MspTokenCacheKey -TenantId $A -Resource Exchange -AppId $App
        }
        @($keys | Select-Object -Unique).Count | Should -Be 3
    }
}

Describe 'Resolve-MspTokenRequestScope' {
    It 'defaults to Graph .default with offline_access' {
        $r = InModuleScope MspGdap { Resolve-MspTokenRequestScope }
        $r.Resource | Should -Be 'https://graph.microsoft.com'
        $r.ScopeString | Should -Be 'https://graph.microsoft.com/.default offline_access'
    }

    It 'qualifies bare Graph scopes' {
        $r = InModuleScope MspGdap { Resolve-MspTokenRequestScope -Scope 'User.Read.All', 'Group.Read.All' }
        $r.ScopeString | Should -Be 'https://graph.microsoft.com/User.Read.All https://graph.microsoft.com/Group.Read.All offline_access'
        $r.ShortScopes | Should -Be @('User.Read.All', 'Group.Read.All')
    }

    It 'refuses scopes from more than one resource' {
        { InModuleScope MspGdap { Resolve-MspTokenRequestScope -Scope 'User.Read.All', 'https://outlook.office365.com/Exchange.Manage' } } |
            Should -Throw -ErrorId 'MspGdap.Scope.MultipleResources'
    }
}

Describe 'Invoke-MspTokenRequest error mapping' {
    BeforeAll {
        Mock Start-Sleep -ModuleName MspGdap {}
    }

    It 'maps AADSTS<Code> to <Reason> with guidance' -ForEach @(
        @{ Code = '50076'; OAuth = 'interaction_required'; Reason = 'MfaRequired'; Text = 'Multifactor' }
        @{ Code = '50079'; OAuth = 'interaction_required'; Reason = 'MfaRequired'; Text = 'register for multifactor' }
        @{ Code = '65001'; OAuth = 'invalid_grant'; Reason = 'ConsentRequired'; Text = 'Grant-MspPartnerAppConsent' }
        @{ Code = '700082'; OAuth = 'invalid_grant'; Reason = 'RefreshTokenExpired'; Text = 'Register-MspPartnerToken' }
        @{ Code = '53003'; OAuth = 'invalid_grant'; Reason = 'ConditionalAccess'; Text = 'Conditional Access' }
        @{ Code = '50020'; OAuth = 'invalid_grant'; Reason = 'NoGdapAccess'; Text = 'GDAP' }
    ) {
        $json = @{
            error             = $OAuth
            error_description = "AADSTS$($Code): Something happened.`r`nTrace ID: 0000 Correlation ID: abcd Timestamp: 2026-10-07"
            error_codes       = @([int]$Code)
            trace_id          = '0000'
            correlation_id    = 'abcd'
        } | ConvertTo-Json -Compress
        $script:CurrentErrorJson = $json
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { throw (New-TestHttpError -Status 400 -Json $script:CurrentErrorJson) }

        $caught = $null
        try {
            InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId } {
                param($Tenant)
                Invoke-MspTokenRequest -TenantId $Tenant -Body ([ordered]@{ grant_type = 'refresh_token'; refresh_token = 'rt-secret-value' })
            }
        }
        catch { $caught = $_ }

        $caught | Should -Not -BeNullOrEmpty
        $caught.FullyQualifiedErrorId | Should -BeLike "MspGdap.TokenRequest.AADSTS$Code*"
        $caught.Exception.Data['Reason'] | Should -Be $Reason
        $caught.Exception.Data['AadstsCode'] | Should -Be $Code
        $caught.Exception.Message | Should -BeLike "*$Text*"
        $caught.Exception.Message | Should -Not -BeLike '*rt-secret-value*'
        $caught.Exception.Message | Should -Not -BeLike '*Trace ID*'
    }

    It 'retries transient failures and then succeeds' {
        $script:tokenCalls = 0
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith {
            $script:tokenCalls++
            if ($script:tokenCalls -eq 1) { throw (New-TestHttpError -Status 503 -Json '{"error":"temporarily_unavailable"}') }
            [pscustomobject]@{ access_token = 'x'; expires_in = 3600 }
        }
        $result = InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId } {
            param($Tenant)
            Invoke-MspTokenRequest -TenantId $Tenant -Body ([ordered]@{ grant_type = 'refresh_token' })
        }
        $result.access_token | Should -Be 'x'
        $script:tokenCalls | Should -Be 2
        Should -Invoke Start-Sleep -ModuleName MspGdap -Times 1
    }

    It 'never writes the request body to verbose output' {
        Mock Invoke-RestMethod -ModuleName MspGdap -MockWith { [pscustomobject]@{ access_token = 'x' } }
        $verbose = InModuleScope MspGdap -Parameters @{ Tenant = $script:TestCustomerTenantId } {
            param($Tenant)
            Invoke-MspTokenRequest -TenantId $Tenant -Body ([ordered]@{ grant_type = 'refresh_token'; refresh_token = 'rt-secret-value'; client_assertion = 'assertion-secret' }) -Verbose 4>&1 |
                Where-Object { $_ -is [System.Management.Automation.VerboseRecord] }
        }
        ($verbose | Out-String) | Should -Not -Match 'rt-secret-value|assertion-secret'
    }
}
