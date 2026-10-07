#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers' 'CoreTestHelpers.ps1')
    Import-Module $script:ModuleManifestPath -Force
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'Start-MspLoopbackListener' {
    It 'returns the code for a callback with the right state and ignores others' {
        $port = InModuleScope MspGdap { Get-MspFreeLoopbackPort }
        $job = Start-ThreadJob -ArgumentList $port -ScriptBlock {
            param($Port)
            Start-Sleep -Milliseconds 700
            $results = @()
            $results += (Invoke-WebRequest -Uri "http://localhost:$Port/favicon.ico" -SkipHttpErrorCheck -TimeoutSec 5).StatusCode
            $results += (Invoke-WebRequest -Uri "http://localhost:$Port/?code=forged&state=wrong" -SkipHttpErrorCheck -TimeoutSec 5).StatusCode
            $results += (Invoke-WebRequest -Uri "http://localhost:$Port/?code=good-code&state=expected-state" -SkipHttpErrorCheck -TimeoutSec 5).StatusCode
            $results
        }
        try {
            $result = InModuleScope MspGdap -Parameters @{ Port = $port } {
                param($Port)
                Start-MspLoopbackListener -Port $Port -ExpectedState 'expected-state' -AuthorizationUrl 'https://login.microsoftonline.com/x' -TimeoutSeconds 20 -NoBrowser 6>$null 3>$null
            }
            $statuses = $job | Wait-Job -Timeout 20 | Receive-Job
        }
        finally {
            $job | Remove-Job -Force
        }
        $result.Code | Should -Be 'good-code'
        $result.RedirectUri | Should -Be "http://localhost:$port"
        $statuses | Should -Be @(404, 400, 200)
    }

    It 'closes the listener when it returns' {
        $port = InModuleScope MspGdap { Get-MspFreeLoopbackPort }
        $job = Start-ThreadJob -ArgumentList $port -ScriptBlock {
            param($Port)
            Start-Sleep -Milliseconds 500
            $null = Invoke-WebRequest -Uri "http://localhost:$Port/?code=c&state=s" -SkipHttpErrorCheck -TimeoutSec 5
        }
        try {
            $null = InModuleScope MspGdap -Parameters @{ Port = $port } {
                param($Port)
                Start-MspLoopbackListener -Port $Port -ExpectedState 's' -AuthorizationUrl 'https://login.microsoftonline.com/x' -TimeoutSeconds 20 -NoBrowser 6>$null
            }
            $null = $job | Wait-Job -Timeout 20
        }
        finally {
            $job | Remove-Job -Force
        }
        # The port can be bound again, so the listener released it.
        $probe = [System.Net.HttpListener]::new()
        $probe.Prefixes.Add("http://localhost:$port/")
        { $probe.Start() } | Should -Not -Throw
        $probe.Close()
    }

    It 'times out cleanly' {
        $port = InModuleScope MspGdap { Get-MspFreeLoopbackPort }
        {
            InModuleScope MspGdap -Parameters @{ Port = $port } {
                param($Port)
                Start-MspLoopbackListener -Port $Port -ExpectedState 's' -AuthorizationUrl 'https://login.microsoftonline.com/x' -TimeoutSeconds 1 -NoBrowser 6>$null
            }
        } | Should -Throw -ErrorId 'MspGdap.Listener.Timeout'
    }

    It 'turns an OAuth error callback into an error' {
        $port = InModuleScope MspGdap { Get-MspFreeLoopbackPort }
        $job = Start-ThreadJob -ArgumentList $port -ScriptBlock {
            param($Port)
            Start-Sleep -Milliseconds 500
            $null = Invoke-WebRequest -Uri "http://localhost:$Port/?error=access_denied&error_description=User%20cancelled&state=s" -SkipHttpErrorCheck -TimeoutSec 5
        }
        try {
            {
                InModuleScope MspGdap -Parameters @{ Port = $port } {
                    param($Port)
                    Start-MspLoopbackListener -Port $Port -ExpectedState 's' -AuthorizationUrl 'https://login.microsoftonline.com/x' -TimeoutSeconds 20 -NoBrowser 6>$null
                }
            } | Should -Throw -ErrorId 'MspGdap.Listener.access_denied'
        }
        finally {
            $job | Remove-Job -Force
        }
    }
}

Describe 'Register-MspPartnerToken' {
    BeforeAll {
        Mock Assert-MspSecretManagement -ModuleName MspGdap -MockWith {}
        Mock Get-MspFreeLoopbackPort -ModuleName MspGdap -MockWith { 50123 }
        Mock New-MspRandomString -ModuleName MspGdap -MockWith { 'fixed-random' }
        Mock Get-MspClientCredentialBody -ModuleName MspGdap -MockWith { [ordered]@{ client_assertion_type = 't'; client_assertion = 'test-assertion' } }
        Mock Start-MspLoopbackListener -ModuleName MspGdap -MockWith {
            [pscustomobject]@{ Code = 'auth-code-value'; RedirectUri = "http://localhost:$Port" }
        }
        Mock Get-SecretInfo -ModuleName MspGdap -MockWith { $null }
        Mock Set-MspConfiguration -ModuleName MspGdap -MockWith {}

        $script:NewSignInResponse = {
            param([string[]]$Amr = @('pwd', 'mfa'), [string]$Tenant = $script:TestPartnerTenantId, [string]$Upn = $script:TestUpn, [string]$Nonce = 'fixed-random')
            $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
            [pscustomobject]@{
                token_type    = 'Bearer'
                access_token  = New-TestJwt -Claims @{ aud = 'https://graph.microsoft.com'; tid = $Tenant; exp = $now + 3600; amr = $Amr; upn = $Upn }
                id_token      = New-TestJwt -Claims @{ aud = $script:TestAppId; tid = $Tenant; preferred_username = $Upn; nonce = $Nonce }
                refresh_token = 'rt-registered-value'
                expires_in    = 3600
                scope         = 'https://graph.microsoft.com/.default'
            }
        }
    }

    BeforeEach {
        Initialize-TestModuleState
        Mock Set-Secret -ModuleName MspGdap -MockWith {}
        $script:SignIn = & $script:NewSignInResponse
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/oauth2/v2.0/token' } -MockWith { $script:SignIn }
    }

    It 'opens an authorisation code + PKCE request at the partner tenant' {
        $null = Register-MspPartnerToken -UserPrincipalName $script:TestUpn
        Should -Invoke Start-MspLoopbackListener -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter {
            $AuthorizationUrl -like "https://login.microsoftonline.com/$($script:TestPartnerTenantId)/oauth2/v2.0/authorize?*" -and
            $AuthorizationUrl -match 'code_challenge_method=S256' -and
            $AuthorizationUrl -match 'response_type=code' -and
            $AuthorizationUrl -match 'redirect_uri=http%3A%2F%2Flocalhost%3A50123' -and
            $AuthorizationUrl -match 'offline_access' -and
            $AuthorizationUrl -notmatch 'devicecode' -and
            $ExpectedState -eq 'fixed-random'
        }
    }

    It 'redeems the code with the PKCE verifier and the app credential' {
        $null = Register-MspPartnerToken -UserPrincipalName $script:TestUpn
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter {
            "$Uri" -eq "https://login.microsoftonline.com/$($script:TestPartnerTenantId)/oauth2/v2.0/token" -and
            $Body['grant_type'] -eq 'authorization_code' -and
            $Body['code'] -eq 'auth-code-value' -and
            $Body['redirect_uri'] -eq 'http://localhost:50123' -and
            $Body['code_verifier'].Length -eq 43 -and
            $Body['client_assertion'] -eq 'test-assertion'
        }
    }

    It 'stores the refresh token in the vault under a hashed name' {
        $null = Register-MspPartnerToken -UserPrincipalName $script:TestUpn
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter {
            (ConvertFrom-TestSecure $SecureStringSecret) -eq 'rt-registered-value' -and
            $Vault -eq 'TestVault' -and
            $Name -match "^MspGdap-$($script:TestAppId)-[0-9a-f]{16}$"
        }
    }

    It 'returns an object without any token values' {
        $result = Register-MspPartnerToken -UserPrincipalName $script:TestUpn
        $result.UserPrincipalName | Should -Be $script:TestUpn
        $result.MfaConfirmed | Should -BeTrue
        $result.RefreshTokenStored | Should -BeTrue
        $json = $result | ConvertTo-Json -Depth 5
        $json | Should -Not -Match 'eyJ|rt-registered-value|auth-code-value|test-assertion'
    }

    It 'caches the partner-tenant Graph token from the sign-in' {
        $null = Register-MspPartnerToken -UserPrincipalName $script:TestUpn
        Mock Get-Secret -ModuleName MspGdap -MockWith { New-TestSecure 'unused' }
        $token = Get-MspAccessToken -PartnerTenant
        $token.FromCache | Should -BeTrue
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly
    }

    It 'refuses to store a token from a sign-in without MFA' {
        $script:SignIn = & $script:NewSignInResponse -Amr @('pwd')
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.MfaMissing*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'stores a non-MFA token only with -SkipMfaCheck and warns' {
        $script:SignIn = & $script:NewSignInResponse -Amr @('pwd')
        $result = Register-MspPartnerToken -UserPrincipalName $script:TestUpn -SkipMfaCheck -WarningAction SilentlyContinue -WarningVariable warnings
        $result.MfaConfirmed | Should -BeFalse
        ($warnings | Out-String) | Should -BeLike '*MFA*'
    }

    It 'refuses a sign-in from another tenant' {
        $script:SignIn = & $script:NewSignInResponse -Tenant $script:TestOtherTenantId
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.WrongTenant*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'refuses when the signed-in account is not the expected technician' {
        $script:SignIn = & $script:NewSignInResponse -Upn 'someone-else@contoso.onmicrosoft.com'
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.UserMismatch*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'refuses an id_token with the wrong nonce' {
        $script:SignIn = & $script:NewSignInResponse -Nonce 'replayed'
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.NonceMismatch*'
    }

    It 'refuses an id_token without a nonce' {
        $script:SignIn = & $script:NewSignInResponse -Nonce ''
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.NonceMismatch*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'refuses a sign-in that returned no id_token' {
        $script:SignIn = & $script:NewSignInResponse
        $script:SignIn.PSObject.Properties.Remove('id_token')
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.NoIdToken*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'refuses to store a token when MFA cannot be confirmed because amr is missing' {
        $script:SignIn = & $script:NewSignInResponse -Amr @()
        { Register-MspPartnerToken -UserPrincipalName $script:TestUpn } | Should -Throw -ErrorId 'MspGdap.Register.MfaUnverified*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'stores a token without amr only with -SkipMfaCheck, and warns' {
        $script:SignIn = & $script:NewSignInResponse -Amr @()
        $result = Register-MspPartnerToken -UserPrincipalName $script:TestUpn -SkipMfaCheck -WarningAction SilentlyContinue -WarningVariable warnings
        $result.RefreshTokenStored | Should -BeTrue
        $result.MfaConfirmed | Should -BeNullOrEmpty
        ($warnings | Out-String) | Should -BeLike '*SkipMfaCheck*'
    }

    It 'does nothing with -WhatIf' {
        Register-MspPartnerToken -UserPrincipalName $script:TestUpn -WhatIf
        Should -Invoke Start-MspLoopbackListener -ModuleName MspGdap -Times 0 -Exactly
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'never writes token values to verbose output' {
        $streams = Register-MspPartnerToken -UserPrincipalName $script:TestUpn -Verbose 4>&1 3>&1 |
            Where-Object { $_ -is [System.Management.Automation.VerboseRecord] -or $_ -is [System.Management.Automation.WarningRecord] }
        ($streams | ForEach-Object { $_.ToString() }) -join "`n" | Should -Not -Match 'eyJ|rt-registered-value|auth-code-value|test-assertion'
    }
}
