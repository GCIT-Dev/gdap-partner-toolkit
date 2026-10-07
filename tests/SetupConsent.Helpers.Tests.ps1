#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Support/SetupConsent.TestSupport.ps1')
    Mock Start-Sleep {}
}

AfterAll { Remove-Variable -Name MspTest -Scope Global -ErrorAction SilentlyContinue }

Describe 'New-MspOperationResult' {
    It 'never reports success when a step failed' {
        $r = New-MspOperationResult -Operation 'x' -Steps @(
            (New-MspStepResult -Step 'a' -Status Changed), (New-MspStepResult -Step 'b' -Status Failed))
        $r.Success | Should -BeFalse
        $r.Outcome | Should -Be 'Failed'
    }
    It 'treats Unknown as a failure' {
        (New-MspOperationResult -Operation 'x' -Steps @(New-MspStepResult -Step 'a' -Status Unknown)).Success | Should -BeFalse
    }
    It 'reports WhatIf runs as not successful' {
        $r = New-MspOperationResult -Operation 'x' -Steps @(New-MspStepResult -Step 'a' -Status WhatIf)
        $r.Success | Should -BeFalse
        $r.Outcome | Should -Be 'WhatIf'
    }
    It 'succeeds when every step passed, changed, was skipped or warned' {
        $r = New-MspOperationResult -Operation 'x' -Steps @(
            (New-MspStepResult -Step 'a' -Status Passed), (New-MspStepResult -Step 'b' -Status Changed),
            (New-MspStepResult -Step 'c' -Status Skipped), (New-MspStepResult -Step 'd' -Status Warning))
        $r.Success | Should -BeTrue
    }
    It 'is not successful with no steps' {
        (New-MspOperationResult -Operation 'x' -Steps @()).Success | Should -BeFalse
    }
}

Describe 'Get-MspHttpStatusCode' {
    It 'reads Exception.Data StatusCode' {
        $record = New-Object System.Management.Automation.ErrorRecord((New-FakeHttpError -StatusCode 409), 'x', 'NotSpecified', $null)
        Get-MspHttpStatusCode -ErrorRecord $record | Should -Be 409
    }
    It 'reads the status from the message' {
        $record = New-Object System.Management.Automation.ErrorRecord((New-Object System.Exception 'Response status code does not indicate success: 404 (Not Found).'), 'x', 'NotSpecified', $null)
        Get-MspHttpStatusCode -ErrorRecord $record | Should -Be 404
    }
    It 'returns null when there is no status' {
        $record = New-Object System.Management.Automation.ErrorRecord((New-Object System.Exception 'DNS failure'), 'x', 'NotSpecified', $null)
        Get-MspHttpStatusCode -ErrorRecord $record | Should -BeNullOrEmpty
    }
}

Describe 'Read-MspPermissionManifest' {
    It 'loads <Name> and strips documentation fields from the Graph view' -TestCases @(
        @{ Name = 'partner-app.minimal.json' }, @{ Name = 'partner-app.full.json' }, @{ Name = 'automation-app.example.json' }
    ) {
        param($Name)
        $m = Read-MspPermissionManifest -Path (Join-Path $global:MspTest.RepoRoot (Join-Path 'manifests' $Name))
        @($m.GraphRequiredResourceAccess).Count | Should -BeGreaterThan 0
        $json = $m.GraphRequiredResourceAccess | ConvertTo-Json -Depth 5
        $json | Should -Not -Match 'resourceDisplayName'
        $json | Should -Not -Match '"value"'
    }
    It 'rejects duplicate permissions' {
        $path = Join-Path $TestDrive 'dup.json'
        '{"requiredResourceAccess":[{"resourceAppId":"00000003-0000-0000-c000-000000000000","resourceAccess":[{"id":"e1fe6dd8-ba31-4d61-89e7-88639da4683d","type":"Scope"},{"id":"e1fe6dd8-ba31-4d61-89e7-88639da4683d","type":"Scope"}]}]}' | Set-Content -Path $path
        { Read-MspPermissionManifest -Path $path } | Should -Throw '*twice*'
    }
    It 'rejects unknown permission types' {
        $path = Join-Path $TestDrive 'type.json'
        '{"requiredResourceAccess":[{"resourceAppId":"00000003-0000-0000-c000-000000000000","resourceAccess":[{"id":"e1fe6dd8-ba31-4d61-89e7-88639da4683d","type":"Delegated"}]}]}' | Set-Content -Path $path
        { Read-MspPermissionManifest -Path $path } | Should -Throw '*Scope or Role*'
    }
}

Describe 'GDAP role catalogue' {
    BeforeAll { $catalog = @(Get-MspGdapRoleCatalog) }
    It 'never puts Global Administrator in the default set' {
        @($catalog | Where-Object { $_.Default -and $_.RoleTemplateId -eq '62e90394-69f5-4237-9190-012177145e10' }).Count | Should -Be 0
    }
    It 'has unique, well-formed template IDs' {
        @($catalog.RoleTemplateId | Sort-Object -Unique).Count | Should -Be $catalog.Count
        $catalog.RoleTemplateId | ForEach-Object { $_ | Should -Match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$' }
    }
    It 'keeps Privileged Role Administrator out of the default set' {
        ($catalog | Where-Object DisplayName -eq 'Privileged Role Administrator').Default | Should -BeFalse
    }
    It 'resolves names, IDs and objects' {
        $r = @(Resolve-MspGdapRole -Role @('Exchange Administrator', 'f2ef992c-3afb-46b9-b7cf-a126ee74c451', [pscustomobject]@{ displayName = 'User Administrator'; roleTemplateId = 'fe930be7-5e62-47db-91af-98c3a49a38b1' }) -Catalog $catalog)
        $r.DisplayName | Should -Be @('Exchange Administrator', 'Global Reader', 'User Administrator')
    }
    It 'throws on an unknown role name' {
        { Resolve-MspGdapRole -Role 'Exchange Admin' -Catalog $catalog } | Should -Throw '*Unknown role name*'
    }
    It 'refuses an unknown template ID unless -AllowUnknown, and marks it Unknown' {
        { Resolve-MspGdapRole -Role 'abababab-abab-abab-abab-abababababab' -Catalog $catalog } | Should -Throw '*not in the MspGdap role catalogue*'
        $r = Resolve-MspGdapRole -Role 'abababab-abab-abab-abab-abababababab' -Catalog $catalog -AllowUnknown
        $r.Unknown | Should -BeTrue
        $r.DisplayName | Should -Be 'Unknown role abababab-abab-abab-abab-abababababab'
    }
    It 'agrees with the well-known role IDs used by the guardrails' {
        foreach ($name in 'GlobalAdministrator', 'PrivilegedRoleAdministrator', 'ExchangeAdministrator', 'ExchangeRecipientAdministrator') {
            $id = Get-MspWellKnownRoleId -Name $name
            @($catalog | Where-Object RoleTemplateId -eq $id).Count | Should -Be 1
        }
        foreach ($id in Get-MspWellKnownRoleId -Name ExchangeAppOnlySupported) { @($catalog | Where-Object RoleTemplateId -eq $id).Count | Should -Be 1 }
        ($catalog | Where-Object RoleTemplateId -eq (Get-MspWellKnownRoleId -Name GlobalAdministrator)).DisplayName | Should -Be 'Global Administrator'
    }
    It 'throws when a name and ID disagree' {
        { Resolve-MspGdapRole -Role @([pscustomobject]@{ displayName = 'Global Reader'; roleTemplateId = '29232cdf-9323-42fd-ade2-1d097af3e4de' }) -Catalog $catalog } | Should -Throw '*does not match*'
    }
}

Describe 'Read-MspGdapAccessMap' {
    It 'refuses the shipped example until group IDs are replaced' {
        { Read-MspGdapAccessMap -Path (Join-Path $global:MspTest.ModuleRoot 'Data/gdap-access.example.json') } | Should -Throw '*Replace it with the object ID*'
    }
    It 'reads a valid map with object and string roles' {
        $path = Join-Path $TestDrive 'map.json'
        @{ accessMapVersion = 1; assignments = @(
                @{ groupDisplayName = 'Helpdesk'; groupId = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'; roles = @('Helpdesk Administrator', @{ displayName = 'Directory Readers'; roleTemplateId = '88d8e3e3-8f55-4a1e-953a-9b9898b8876b' }) }
            ) } | ConvertTo-Json -Depth 6 | Set-Content -Path $path
        $map = @(Read-MspGdapAccessMap -Path $path)
        $map.Count | Should -Be 1
        $map[0].Roles.RoleTemplateId | Should -Contain '729827e3-9c14-49f7-bb1b-9608f156bbb8'
    }
}

Describe 'New-MspAddKeyProof' {
    It 'builds an RS256 proof token Graph addKey accepts' {
        $cert = New-TestCertificate
        $objectId = 'aaaaaaaa-1111-2222-3333-bbbbbbbbbbbb'
        $now = [datetime]::UtcNow
        $jwt = New-MspAddKeyProof -Certificate $cert -ApplicationObjectId $objectId -Now $now
        $parts = $jwt.Split('.')
        $parts.Count | Should -Be 3
        $decode = { param($s) $p = $s.Replace('-', '+').Replace('_', '/'); while ($p.Length % 4) { $p += '=' }; [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json }
        $header = & $decode $parts[0]
        $payload = & $decode $parts[1]
        $header.alg | Should -Be 'RS256'
        $header.x5t | Should -Not -BeNullOrEmpty
        $payload.aud | Should -Be '00000002-0000-0000-c000-000000000000'
        $payload.iss | Should -Be $objectId
        ($payload.exp - $payload.nbf) | Should -Be 600
        $sig = $parts[2].Replace('-', '+').Replace('_', '/'); while ($sig.Length % 4) { $sig += '=' }
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cert)
        $rsa.VerifyData([Text.Encoding]::UTF8.GetBytes("$($parts[0]).$($parts[1])"), [Convert]::FromBase64String($sig), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1) | Should -BeTrue
    }
}

Describe 'Certificates' {
    It 'matches a keyCredential by thumbprint' {
        $cert = New-TestCertificate
        Test-MspKeyCredentialThumbprint -KeyEntry @([pscustomobject]@{ customKeyIdentifier = (ConvertTo-TestKeyIdentifier $cert.Thumbprint) }) -Thumbprint $cert.Thumbprint | Should -BeTrue
        Test-MspKeyCredentialThumbprint -KeyEntry @([pscustomobject]@{ customKeyIdentifier = 'AAAA' }) -Thumbprint $cert.Thumbprint | Should -BeFalse
    }
    It 'matches the hex thumbprint form that Graph v1.0 returns (regression: live test 7 Oct 2026)' {
        # A 40 character hex thumbprint is also valid Base64, so it must be compared before any Base64 decode.
        $thumb = '2CEF3E3B03EB4CF1913B1CD52054A43BD8602A48'
        Test-MspKeyCredentialThumbprint -KeyEntry @([pscustomobject]@{ customKeyIdentifier = $thumb }) -Thumbprint $thumb | Should -BeTrue
        Test-MspKeyCredentialThumbprint -KeyEntry @([pscustomobject]@{ customKeyIdentifier = $thumb.ToLowerInvariant() }) -Thumbprint $thumb | Should -BeTrue
        Test-MspKeyCredentialThumbprint -KeyEntry @([pscustomobject]@{ customKeyIdentifier = '996E18A2D026CA376CCFA05B8CF239E7ADE5CB08' }) -Thumbprint $thumb | Should -BeFalse
    }
    It 'uploads only the public key' {
        $cert = New-TestCertificate
        $key = ConvertTo-MspKeyCredential -Certificate $cert
        $key.key | Should -Be ([Convert]::ToBase64String($cert.RawData))
        $key.usage | Should -Be 'Verify'
    }
    It 'refuses an expired certificate' {
        $cert = New-TestCertificate -NotBefore ([datetime]::UtcNow.AddDays(-30)) -NotAfter ([datetime]::UtcNow.AddDays(-1))
        { Resolve-MspCertificate -Certificate $cert } | Should -Throw '*expired*'
    }
    It 'refuses keys under 2048 bits' {
        $cert = New-TestCertificate -KeySize 1024
        { Resolve-MspCertificate -Certificate $cert } | Should -Throw '*2048*'
    }
    It 'refuses PFX files' {
        $path = Join-Path $TestDrive 'x.pfx'
        Set-Content -Path $path -Value 'x'
        { Resolve-MspCertificate -Path $path } | Should -Throw '*PFX*'
    }
}

Describe 'Test-MspExoCompatibility' {
    It 'flags EXO 3.10 on PowerShell 7.4' { Test-MspExoCompatibility -ModuleVersion '3.10.1' -PSVersion '7.4.6' -Edition Core | Should -Match '7.6' }
    It 'accepts EXO 3.10 on PowerShell 7.6' { Test-MspExoCompatibility -ModuleVersion '3.10.1' -PSVersion '7.6.0' -Edition Core | Should -BeNullOrEmpty }
    It 'accepts EXO 3.9.2 on PowerShell 7.4' { Test-MspExoCompatibility -ModuleVersion '3.9.2' -PSVersion '7.4.0' -Edition Core | Should -BeNullOrEmpty }
    It 'accepts EXO 3.10 on Windows PowerShell 5.1' { Test-MspExoCompatibility -ModuleVersion '3.10.1' -PSVersion '5.1' -Edition Desktop | Should -BeNullOrEmpty }
    It 'flags modules older than 3.1.0' { Test-MspExoCompatibility -ModuleVersion '3.0.0' -PSVersion '7.6.0' -Edition Core | Should -Match 'too old' }
}

Describe 'Token conversion' {
    It 'returns the string from a token object holding a SecureString' {
        $token = Get-MspAccessToken -TenantId 'x' -Resource 'r'
        ConvertTo-MspPlainToken -InputObject $token | Should -Be 'token-for-r'
    }
    It 'returns a read-only SecureString' {
        $secure = ConvertTo-MspSecureToken -InputObject 'abc'
        $secure | Should -BeOfType [securestring]
        $secure.IsReadOnly() | Should -BeTrue
    }
    It 'throws on a null token' { { ConvertTo-MspPlainToken -InputObject $null } | Should -Throw }
}

Describe 'Wait-MspCondition' {
    It 'returns the value once the condition holds' {
        $counter = @{ n = 0 }
        $r = Wait-MspCondition -TimeoutSeconds 30 -IntervalSeconds 5 -Condition { $counter.n++; if ($counter.n -ge 3) { 'ok' } }
        $r.Satisfied | Should -BeTrue
        $r.Attempts | Should -Be 3
    }
    It 'gives up after the attempt budget' {
        $r = Wait-MspCondition -TimeoutSeconds 10 -IntervalSeconds 5 -Condition { $null }
        $r.Satisfied | Should -BeFalse
        $r.Attempts | Should -Be 3
    }
}

Describe 'Invoke-MspGraphCall' {
    It 'sends absolute URIs to the partner tenant and unwraps raw collections' {
        Set-FakeGraphRoute @(New-FakeRoute -Pattern 'organization' -Response { [pscustomobject]@{ '@odata.context' = 'x'; value = @([pscustomobject]@{ id = 1 }, [pscustomobject]@{ id = 2 }) } })
        $r = @(Invoke-MspGraphCall -PartnerTenant -Path 'organization')
        $r.Count | Should -Be 2
        $call = Get-FakeGraphCall -Method GET
        $call[0].Uri | Should -Be 'https://graph.microsoft.com/v1.0/organization'
        $call[0].PartnerTenant | Should -BeTrue
    }
    It 'turns off nested confirmation on writes' {
        Mock Invoke-MspGraphRequest { [pscustomobject]@{ ok = $true } } -ParameterFilter { $Method -eq 'POST' }
        $null = Invoke-MspGraphCall -TenantId $global:MspTest.CustomerTenantId -Method POST -Path 'x' -Body @{}
        Should -Invoke Invoke-MspGraphRequest -Times 1 -ParameterFilter { $Confirm -eq $false -and $TenantId -eq $global:MspTest.CustomerTenantId }
    }
}

Describe 'Invoke-MspPartnerCenterRequest' {
    It 'sends ValidateMfa and reads isMfaCompliant' {
        Mock Invoke-WebRequest { [pscustomobject]@{ StatusCode = 201; Content = '{"applicationId":"x"}'; Headers = @{ 'isMfaCompliant' = @('true') } } }
        $r = Invoke-MspPartnerCenterRequest -Method POST -Path 'customers/abc/applicationconsents' -Body @{ a = 1 }
        $r.StatusCode | Should -Be 201
        $r.IsMfaCompliant | Should -BeTrue
        Should -Invoke Invoke-WebRequest -Times 1 -ParameterFilter { $Headers['ValidateMfa'] -eq 'true' -and $Uri -eq 'https://api.partnercenter.microsoft.com/v1/customers/abc/applicationconsents' }
    }
    It 'maps 401 MFA required to MfaRequired' {
        Mock Invoke-WebRequest { throw (New-FakeHttpError -StatusCode 401 -Message 'Unauthorized - MFA required') }
        $r = Invoke-MspPartnerCenterRequest -Method GET -Path 'customers'
        $r.MfaRequired | Should -BeTrue
        $r.Success | Should -BeFalse
        $r.Error | Should -Match 'Register-MspPartnerToken'
    }
    It 'retries 503 then succeeds' {
        $script:pcAttempts = 0
        Mock Invoke-WebRequest {
            $script:pcAttempts++
            if ($script:pcAttempts -eq 1) { throw (New-FakeHttpError -StatusCode 503) }
            [pscustomobject]@{ StatusCode = 200; Content = '{}'; Headers = @{} }
        }
        (Invoke-MspPartnerCenterRequest -Method GET -Path 'customers').StatusCode | Should -Be 200
        Should -Invoke Invoke-WebRequest -Times 2
    }
}

Describe 'Test-MspConsentMissingError' {
    BeforeAll {
        function New-TestErrorRecord {
            param([System.Exception]$Exception, [string]$Details)
            $record = New-Object System.Management.Automation.ErrorRecord ($Exception, 'Test', 'NotSpecified', $null)
            if ($Details) { $record.ErrorDetails = New-Object System.Management.Automation.ErrorDetails ($Details) }
            $record
        }
    }
    It 'recognises AADSTS65001 in the message' {
        Test-MspConsentMissingError -ErrorRecord (New-TestErrorRecord -Exception ([System.Exception]::new('AADSTS65001: The user or administrator has not consented to use the application.'))) | Should -BeTrue
    }
    It 'recognises AADSTS65001 in the error details' {
        Test-MspConsentMissingError -ErrorRecord (New-TestErrorRecord -Exception ([System.Exception]::new('Token request failed (400).')) -Details '{"error":"invalid_grant","error_description":"AADSTS65001: The user or administrator has not consented."}') | Should -BeTrue
    }
    It 'recognises AADSTS65001 in an inner exception' {
        $inner = [System.Exception]::new('AADSTS65001: not consented')
        Test-MspConsentMissingError -ErrorRecord (New-TestErrorRecord -Exception ([System.Exception]::new('Could not get a token.', $inner))) | Should -BeTrue
    }
    It 'does not match other errors, including longer codes that start with 65001 (<Text>)' -ForEach @(
        @{ Text = 'AADSTS650052: The app needs access to a service that your organization has not subscribed to.' }
        @{ Text = 'AADSTS650051: The service principal already exists.' }
        @{ Text = 'AADSTS50020: User account does not exist in tenant.' }
        @{ Text = 'Response status code does not indicate success: 403 (Forbidden).' }
    ) {
        Test-MspConsentMissingError -ErrorRecord (New-TestErrorRecord -Exception ([System.Exception]::new($Text))) | Should -BeFalse
    }
}

Describe 'Get-MspEmptyConsentState' {
    It 'marks every desired scope as missing and every resource as present, with no grant' {
        $grants = @(
            [pscustomobject]@{ ResourceAppId = '00000003-0000-0000-c000-000000000000'; ResourceDisplayName = 'Microsoft Graph'; Scopes = @('User.Read', 'Directory.Read.All') }
            [pscustomobject]@{ ResourceAppId = '00000002-0000-0ff1-ce00-000000000000'; ResourceDisplayName = 'Office 365 Exchange Online'; Scopes = @('Exchange.Manage') }
        )
        $state = Get-MspEmptyConsentState -TenantId $global:MspTest.CustomerTenantId -Grant $grants
        $state.NotConsented | Should -BeTrue
        $state.AppServicePrincipal | Should -BeNullOrEmpty
        $state.Resources.Count | Should -Be 2
        foreach ($resource in $state.Resources) {
            $resource.ResourcePresent | Should -BeTrue
            $resource.GrantId | Should -BeNullOrEmpty
            @($resource.CurrentScopes).Count | Should -Be 0
            $resource.MissingScopes | Should -Be $resource.DesiredScopes
        }
        @($state.OtherGrants).Count | Should -Be 0
    }
}