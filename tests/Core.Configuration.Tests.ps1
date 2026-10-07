#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.5.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Helpers' 'CoreTestHelpers.ps1')
    Import-Module $script:ModuleManifestPath -Force
}

AfterAll {
    Remove-Module MspGdap -Force -ErrorAction SilentlyContinue
}

Describe 'Set-MspConfiguration and Get-MspConfiguration' {
    BeforeEach {
        $script:ConfigFile = Join-Path $TestDrive "cfg-$([guid]::NewGuid())" 'config.json'
        InModuleScope MspGdap { $script:MspState.Config = $null; $script:MspState.CertificatePassword = $null }
        Mock Assert-MspSecretManagement -ModuleName MspGdap -MockWith {}
        Mock Set-Secret -ModuleName MspGdap -MockWith {}
    }

    It 'writes non-secret settings to JSON and reads them back' {
        Set-MspConfiguration -Path $script:ConfigFile -PartnerTenantId $script:TestPartnerTenantId -AppId $script:TestAppId -CertificateThumbprint ('ab' * 20) -VaultName 'MspGdapVault' -TechnicianUpn $script:TestUpn -Confirm:$false
        Test-Path $script:ConfigFile | Should -BeTrue
        $stored = Get-Content $script:ConfigFile -Raw | ConvertFrom-Json
        $stored.PartnerTenantId | Should -Be $script:TestPartnerTenantId
        $stored.CertificateThumbprint | Should -Be ('AB' * 20)
        $stored.CredentialType | Should -Be 'Certificate'

        InModuleScope MspGdap { $script:MspState.Config = $null }
        $config = Get-MspConfiguration -Path $script:ConfigFile
        $config.AppId | Should -Be $script:TestAppId
        $config.VaultName | Should -Be 'MspGdapVault'
        $config.SigningAlgorithm | Should -Be 'PS256'
        $config.FileExists | Should -BeTrue
    }

    It 'only changes the parameters passed' {
        Set-MspConfiguration -Path $script:ConfigFile -PartnerTenantId $script:TestPartnerTenantId -AppId $script:TestAppId -VaultName 'V1' -Confirm:$false
        Set-MspConfiguration -VaultName 'V2' -Confirm:$false
        $stored = Get-Content $script:ConfigFile -Raw | ConvertFrom-Json
        $stored.VaultName | Should -Be 'V2'
        $stored.AppId | Should -Be $script:TestAppId
    }

    It 'keeps the certificate password for the session only' {
        $password = New-TestSecure 'pfx-password-value'
        Set-MspConfiguration -Path $script:ConfigFile -CertificatePath (Join-Path $TestDrive 'app.pfx') -CertificatePassword $password -Confirm:$false
        Get-Content $script:ConfigFile -Raw | Should -Not -Match 'pfx-password-value|Password'
        (Get-MspConfiguration).SessionCertificatePassword | Should -BeTrue
    }

    It 'stores a client secret in the vault, never in the file, and warns' {
        $secret = New-TestSecure 'client-secret-value'
        Set-MspConfiguration -Path $script:ConfigFile -AppId $script:TestAppId -VaultName 'MspGdapVault' -ClientSecret $secret -Confirm:$false -WarningVariable warnings -WarningAction SilentlyContinue
        Should -Invoke Set-Secret -ModuleName MspGdap -Times 1 -Exactly -ParameterFilter {
            $Name -eq "MspGdap-$($script:TestAppId)-clientsecret" -and (ConvertFrom-TestSecure $SecureStringSecret) -eq 'client-secret-value'
        }
        $content = Get-Content $script:ConfigFile -Raw
        $content | Should -Not -Match 'client-secret-value'
        ($content | ConvertFrom-Json).CredentialType | Should -Be 'ClientSecret'
        ($warnings | Out-String) | Should -BeLike '*discouraged*'
    }

    It 'does not write anything with -WhatIf' {
        Set-MspConfiguration -Path $script:ConfigFile -PartnerTenantId $script:TestPartnerTenantId -WhatIf
        Test-Path $script:ConfigFile | Should -BeFalse
    }

    It 'rejects a thumbprint and a path together' {
        { Set-MspConfiguration -Path $script:ConfigFile -CertificateThumbprint ('a' * 40) -CertificatePath 'x.pfx' -Confirm:$false } | Should -Throw -ErrorId 'MspGdap.Configuration.Conflict'
    }

    It 'rejects a tenant ID that is not a GUID' {
        { Set-MspConfiguration -Path $script:ConfigFile -PartnerTenantId 'contoso.onmicrosoft.com' -Confirm:$false } | Should -Throw
    }

    It 'reports missing settings before any token request' {
        Set-MspConfiguration -Path $script:ConfigFile -PartnerTenantId $script:TestPartnerTenantId -Confirm:$false
        { Get-MspAccessToken -TenantId $script:TestCustomerTenantId } | Should -Throw -ErrorId 'MspGdap.Configuration.Incomplete*'
    }
}
