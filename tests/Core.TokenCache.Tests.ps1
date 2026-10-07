#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers' 'CoreTestHelpers.ps1')
    Import-Module $script:ModuleManifestPath -Force
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'Test-MspAccessToken' {
    BeforeAll {
        $script:Now = [DateTimeOffset]::UtcNow
        $script:NewEntry = {
            param([int]$Minutes = 60, [string]$Tenant = $script:TestCustomerTenantId, [string]$Aud = 'https://graph.microsoft.com', [string]$TokenTenant = $Tenant, [string]$Scp = 'User.Read.All')
            $exp = $script:Now.AddMinutes($Minutes)
            [pscustomobject]@{
                TenantId    = $Tenant
                Resource    = 'https://graph.microsoft.com'
                AccessToken = New-TestSecure (New-TestJwt -Claims @{ tid = $TokenTenant; aud = $Aud; exp = $exp.ToUnixTimeSeconds(); scp = $Scp })
                ExpiresOn   = $exp
                Scopes      = @($Scp -split ' ')
            }
        }
    }

    It 'accepts a token with more than the 5 minute skew left' {
        $entry = & $script:NewEntry -Minutes 6
        Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph | Should -BeTrue
    }

    It 'rejects a token inside the 5 minute skew window' {
        $entry = & $script:NewEntry -Minutes 4
        $result = Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph -Detailed
        $result.IsValid | Should -BeFalse
        $result.Reasons | Should -Contain 'Expired'
    }

    It 'honours a custom skew' {
        $entry = & $script:NewEntry -Minutes 4
        Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph -SkewMinutes 2 | Should -BeTrue
    }

    It 'rejects a tenant mismatch from metadata' {
        $entry = & $script:NewEntry -Tenant $script:TestOtherTenantId -TokenTenant $script:TestOtherTenantId
        $result = Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph -Detailed
        $result.IsValid | Should -BeFalse
        $result.Reasons | Should -Contain 'TenantMismatch'
    }

    It 'rejects a tenant mismatch from the tid claim' {
        $entry = & $script:NewEntry -TokenTenant $script:TestOtherTenantId
        $result = Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph -Detailed
        $result.Reasons | Should -Contain 'TenantMismatch'
    }

    It 'rejects an audience mismatch' {
        $entry = & $script:NewEntry -Aud 'https://outlook.office365.com'
        $result = Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph -Detailed
        $result.IsValid | Should -BeFalse
        $result.Reasons | Should -Contain 'AudienceMismatch'
    }

    It 'accepts the GUID form of the audience' {
        $entry = & $script:NewEntry -Aud '00000003-0000-0000-c000-000000000000'
        Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Graph | Should -BeTrue
    }

    It 'rejects a resource mismatch from metadata' {
        $entry = & $script:NewEntry
        $result = Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -Resource Exchange -Detailed
        $result.Reasons | Should -Contain 'ResourceMismatch'
    }

    It 'accepts an opaque token when the response metadata is good' {
        $result = Test-MspAccessToken -AccessToken 'opaque-token' -TenantId $script:TestCustomerTenantId -Resource Graph -ExpiresOn $script:Now.AddMinutes(30) -Detailed
        $result.IsValid | Should -BeTrue
        $result.ClaimsChecked | Should -BeFalse
    }

    It 'rejects an opaque token with no known expiry' {
        $result = Test-MspAccessToken -AccessToken 'opaque-token' -TenantId $script:TestCustomerTenantId -Detailed
        $result.IsValid | Should -BeFalse
        $result.Reasons | Should -Contain 'ExpiryUnknown'
    }

    It 'checks required scopes' {
        $entry = & $script:NewEntry -Scp 'User.Read.All'
        Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -RequiredScope 'User.Read.All' | Should -BeTrue
        $result = Test-MspAccessToken -InputObject $entry -TenantId $script:TestCustomerTenantId -RequiredScope 'Directory.ReadWrite.All' -Detailed
        $result.Reasons | Should -Contain 'MissingScope:Directory.ReadWrite.All'
    }

    It 'checks required roles on decodable app-only tokens' {
        $token = New-TestJwt -Claims @{ tid = $script:TestCustomerTenantId; exp = $script:Now.AddHours(1).ToUnixTimeSeconds(); roles = @('Exchange.ManageAsApp') }
        Test-MspAccessToken -AccessToken $token -TenantId $script:TestCustomerTenantId -RequiredRole 'Exchange.ManageAsApp' | Should -BeTrue
        Test-MspAccessToken -AccessToken $token -TenantId $script:TestCustomerTenantId -RequiredRole 'Mail.Read' | Should -BeFalse
    }

    It 'requires TenantId' {
        $parameter = (Get-Command Test-MspAccessToken).Parameters['TenantId']
        @($parameter.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory }).Count | Should -BeGreaterThan 0
    }
}

Describe 'Get-MspAccessToken' {
    BeforeAll {
        Mock Get-MspClientCredentialBody -ModuleName MspGdap -MockWith {
            [ordered]@{ client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'; client_assertion = 'test-assertion' }
        }
        Mock Get-SecretInfo -ModuleName MspGdap -MockWith { $null }
    }

    BeforeEach {
        Initialize-TestModuleState
        $script:TokenCalls = 0
        $script:IssuedTenant = $null
        Mock Get-Secret -ModuleName MspGdap -MockWith { New-TestSecure 'rt-old-value' }
        Mock Set-Secret -ModuleName MspGdap -MockWith {}
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/oauth2/v2.0/token' } -MockWith {
            $script:TokenCalls++
            $tenant = ([uri]"$Uri").Segments[1].TrimEnd('/')
            $claimTenant = if ($script:IssuedTenant) { $script:IssuedTenant } else { $tenant }
            New-TestTokenResponse -TenantId $claimTenant -RefreshToken "rt-new-value-$($script:TokenCalls)"
        }
    }

    It 'requires -TenantId or -PartnerTenant in every parameter set' {
        $command = Get-Command Get-MspAccessToken
        foreach ($set in $command.ParameterSets) {
            $mandatory = @($set.Parameters | Where-Object IsMandatory | ForEach-Object Name)
            ($mandatory -contains 'TenantId' -or $mandatory -contains 'PartnerTenant') | Should -BeTrue -Because "parameter set $($set.Name) must name a tenant"
        }
    }

    It 'rejects an empty TenantId' {
        { Get-MspAccessToken -TenantId '' } | Should -Throw
    }

    It 'refuses the partner tenant ID unless -PartnerTenant is used' {
        { Get-MspAccessToken -TenantId $script:TestPartnerTenantId } | Should -Throw -ErrorId 'MspGdap.Tenant.PartnerTenantNotAllowed*'
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 0 -Exactly
    }

    It 'redeems the refresh token at the CUSTOMER tenant token endpoint' {
        $token = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $token.TenantId | Should -Be $script:TestCustomerTenantId
        $token.AccessToken | Should -BeOfType [securestring]
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter {
            "$Uri" -eq "https://login.microsoftonline.com/$($script:TestCustomerTenantId)/oauth2/v2.0/token" -and
            $Body['grant_type'] -eq 'refresh_token' -and
            $Body['refresh_token'] -eq 'rt-old-value' -and
            $Body['scope'] -eq 'https://graph.microsoft.com/.default offline_access' -and
            $Body['client_id'] -eq $script:TestAppId
        }
    }

    It 'uses the partner tenant endpoint only with -PartnerTenant' {
        $token = Get-MspAccessToken -PartnerTenant
        $token.TenantId | Should -Be $script:TestPartnerTenantId
        Should -Invoke Invoke-RestMethod -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { "$Uri" -like "*/$($script:TestPartnerTenantId)/oauth2/v2.0/token" }
    }

    It 'writes the rotated refresh token back to the vault' {
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter {
            (ConvertFrom-TestSecure $SecureStringSecret) -eq 'rt-new-value-1' -and
            $Vault -eq 'TestVault' -and
            $Name -match "^MspGdap-$($script:TestAppId)-[0-9a-f]{16}$" -and
            $Metadata.upn -eq $script:TestUpn
        }
    }

    It 'still returns the token but warns when the write-back fails' {
        Mock Set-Secret -ModuleName MspGdap -MockWith { throw 'vault is read only' }
        $token = Get-MspAccessToken -TenantId $script:TestCustomerTenantId -WarningVariable warnings -WarningAction SilentlyContinue
        $token | Should -Not -BeNullOrEmpty
        ($warnings | Out-String) | Should -BeLike '*could not be written back*'
    }

    It 'retries without metadata when the vault does not support metadata' {
        Mock Set-Secret -ModuleName MspGdap -ParameterFilter { $null -ne $Metadata } -MockWith { throw 'The vault does not support secret metadata.' }
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter { $null -eq $Metadata -and (ConvertFrom-TestSecure $SecureStringSecret) -eq 'rt-new-value-1' }
    }

    It 'reuses a valid cached token without another token request' {
        $first = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $second = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $first.FromCache | Should -BeFalse
        $second.FromCache | Should -BeTrue
        $script:TokenCalls | Should -Be 1
    }

    It 'keeps separate cache entries per tenant and per resource' {
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $null = Get-MspAccessToken -TenantId $script:TestOtherTenantId
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/oauth2/v2.0/token' } -MockWith {
            $script:TokenCalls++
            $tenant = ([uri]"$Uri").Segments[1].TrimEnd('/')
            New-TestTokenResponse -TenantId $tenant -Audience 'https://outlook.office365.com' -Scope 'https://outlook.office365.com/Exchange.Manage' -RefreshToken 'rt-x'
        }
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId -Resource Exchange
        $script:TokenCalls | Should -Be 3
        $keys = InModuleScope MspGdap { @($script:MspTokenCache.Keys) }
        $keys.Count | Should -Be 3
        $keys | Should -Contain "$($script:TestCustomerTenantId)|https://outlook.office365.com|$($script:TestAppId)"
    }

    It 'requests a new token when the cached one is inside the skew window' {
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        InModuleScope MspGdap {
            foreach ($key in @($script:MspTokenCache.Keys)) { $script:MspTokenCache[$key].ExpiresOn = [DateTimeOffset]::UtcNow.AddMinutes(4) }
        }
        $again = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $again.FromCache | Should -BeFalse
        $script:TokenCalls | Should -Be 2
    }

    It 'requests a new token when required scopes are missing from the cached one' {
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId -Scope 'Policy.ReadWrite.ConditionalAccess'
        $script:TokenCalls | Should -Be 2
    }

    It 'discards a token whose tid is not the requested tenant and stores nothing' {
        $script:IssuedTenant = $script:TestOtherTenantId
        { Get-MspAccessToken -TenantId $script:TestCustomerTenantId } | Should -Throw -ErrorId 'MspGdap.Token.TenantMismatch*'
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 0 -Exactly
        InModuleScope MspGdap { $script:MspTokenCache.Count } | Should -Be 0
    }

    It 'never writes token values to verbose, warning or information output' {
        $streams = Get-MspAccessToken -TenantId $script:TestCustomerTenantId -Verbose 4>&1 3>&1 6>&1 |
            Where-Object { $_ -is [System.Management.Automation.VerboseRecord] -or $_ -is [System.Management.Automation.WarningRecord] -or $_ -is [System.Management.Automation.InformationRecord] }
        $text = ($streams | ForEach-Object { $_.ToString() }) -join "`n"
        $text | Should -Not -BeNullOrEmpty
        $text | Should -Not -Match 'eyJ'
        $text | Should -Not -Match 'rt-old-value|rt-new-value|test-assertion'
    }

    It 'does not expose the token in default formatting' {
        $token = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        ($token | Format-List * | Out-String) | Should -Not -Match 'eyJ'
        ($token | ConvertTo-Json -Depth 3) | Should -Not -Match 'eyJ'
    }

    It 'returns the raw token only with -AsPlainText' {
        $plain = Get-MspAccessToken -TenantId $script:TestCustomerTenantId -AsPlainText
        $plain | Should -BeOfType [string]
        $plain | Should -Match '^eyJ'
    }

    It 'resolves a verified domain to the tenant ID first' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/.well-known/openid-configuration' } -MockWith {
            [pscustomobject]@{ issuer = "https://login.microsoftonline.com/$($script:TestCustomerTenantId)/v2.0"; token_endpoint = 'x' }
        }
        $token = Get-MspAccessToken -TenantId 'contoso.onmicrosoft.com'
        $token.TenantId | Should -Be $script:TestCustomerTenantId
    }

    It 'does not keep tokens or refresh tokens in global variables' {
        $null = Get-MspAccessToken -TenantId $script:TestCustomerTenantId
        $globals = Get-Variable -Scope Global | Where-Object { $_.Value -is [string] -and ($_.Value -match '^eyJ' -or $_.Value -match 'rt-new-value') }
        $globals | Should -BeNullOrEmpty
    }

    It 'surfaces an expired refresh token with clear guidance' {
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/oauth2/v2.0/token' } -MockWith {
            throw (New-TestHttpError -Status 400 -Json '{"error":"invalid_grant","error_description":"AADSTS700082: The refresh token has expired due to inactivity.","error_codes":[700082]}')
        }
        $caught = $null
        try { Get-MspAccessToken -TenantId $script:TestCustomerTenantId } catch { $caught = $_ }
        $caught.FullyQualifiedErrorId | Should -BeLike 'MspGdap.TokenRequest.AADSTS700082*'
        $caught.Exception.Message | Should -BeLike '*Register-MspPartnerToken*'
    }
}

Describe 'Get-MspAuthHeader' {
    BeforeEach {
        Initialize-TestModuleState
        Mock Get-MspClientCredentialBody -ModuleName MspGdap -MockWith { [ordered]@{ client_assertion_type = 't'; client_assertion = 'a' } }
        Mock Get-SecretInfo -ModuleName MspGdap -MockWith { $null }
        Mock Get-Secret -ModuleName MspGdap -MockWith { New-TestSecure 'rt-old-value' }
        Mock Set-Secret -ModuleName MspGdap -MockWith {}
        Mock Invoke-RestMethod -ModuleName MspGdap -ParameterFilter { "$Uri" -like '*/oauth2/v2.0/token' } -MockWith {
            New-TestTokenResponse -TenantId (([uri]"$Uri").Segments[1].TrimEnd('/')) -RefreshToken 'rt-new'
        }
    }

    It 'returns a bearer header and merges extra headers' {
        $headers = Get-MspAuthHeader -TenantId $script:TestCustomerTenantId -AdditionalHeaders @{ ConsistencyLevel = 'eventual'; Authorization = 'ignored' }
        $headers.Authorization | Should -Match '^Bearer eyJ'
        $headers.ConsistencyLevel | Should -Be 'eventual'
    }

    It 'has the same tenant safety as Get-MspAccessToken' {
        { Get-MspAuthHeader -TenantId $script:TestPartnerTenantId } | Should -Throw -ErrorId 'MspGdap.Tenant.PartnerTenantNotAllowed*'
    }
}

Describe 'Clear-MspTokenCache and Disconnect-Msp' {
    BeforeEach {
        Initialize-TestModuleState
        InModuleScope MspGdap -Parameters @{ A = $script:TestCustomerTenantId; B = $script:TestOtherTenantId; App = $script:TestAppId } {
            param($A, $B, $App)
            $script:MspTokenCache["$A|https://graph.microsoft.com|$App"] = [pscustomobject]@{ TenantId = $A }
            $script:MspTokenCache["$A|https://outlook.office365.com|$App"] = [pscustomobject]@{ TenantId = $A }
            $script:MspTokenCache["$B|https://graph.microsoft.com|$App"] = [pscustomobject]@{ TenantId = $B }
            $script:MspState.CurrentUpn = 'someone@contoso.onmicrosoft.com'
        }
    }

    It 'clears one tenant and resource' {
        Clear-MspTokenCache -TenantId $script:TestCustomerTenantId -Resource Exchange
        InModuleScope MspGdap { $script:MspTokenCache.Count } | Should -Be 2
    }

    It 'clears one tenant' {
        Clear-MspTokenCache -TenantId $script:TestCustomerTenantId
        InModuleScope MspGdap { @($script:MspTokenCache.Keys) } | Should -Be @("$($script:TestOtherTenantId)|https://graph.microsoft.com|$($script:TestAppId)")
    }

    It 'clears everything' {
        Clear-MspTokenCache
        InModuleScope MspGdap { $script:MspTokenCache.Count } | Should -Be 0
    }

    It 'Disconnect-Msp clears the cache and session state' {
        Disconnect-Msp
        InModuleScope MspGdap {
            $script:MspTokenCache.Count | Should -Be 0
            $script:MspState.CurrentUpn | Should -BeNullOrEmpty
            $script:MspState.Config | Should -BeNullOrEmpty
            $script:MspState.CertificatePassword | Should -BeNullOrEmpty
        }
    }
}
