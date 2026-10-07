#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'Support/SetupConsent.TestSupport.ps1')
    Mock Start-Sleep {}
    $customer = $global:MspTest.CustomerTenantId
    Set-FakeGraphRoute @(
        New-FakeRoute -Pattern 'organization' -Response { [pscustomobject]@{ id = $customer; displayName = 'Fabrikam'; verifiedDomains = @([pscustomobject]@{ name = 'fabrikam.com'; isInitial = $false }, [pscustomobject]@{ name = 'fabrikam.onmicrosoft.com'; isInitial = $true }) } }
    )
}

AfterAll { Remove-Variable -Name MspTest -Scope Global -ErrorAction SilentlyContinue }

Describe 'Connect-MspExchangeOnline' {
    BeforeEach {
        Mock Assert-MspModuleAvailable { [version]'3.9.2' }
        $global:MspTest.State = @{ Sessions = @([pscustomobject]@{ ConnectionId = 'old'; State = 'Connected'; TenantID = '99999999-9999-9999-9999-999999999999' }); NewTenant = $customer }
        Mock Get-ConnectionInformation { $global:MspTest.State.Sessions }
        Mock Disconnect-ExchangeOnline { if (-not $ConnectionId) { $global:MspTest.State.Sessions = @() } }
        Mock Connect-ExchangeOnline {
            $global:MspTest.State.Sessions = @($global:MspTest.State.Sessions) + [pscustomobject]@{ ConnectionId = 'new'; State = 'Connected'; TenantID = $global:MspTest.State.NewTenant; DelegatedOrganization = $DelegatedOrganization; UserPrincipalName = 'tech@contoso.onmicrosoft.com'; TokenExpiryTimeUTC = [datetime]::UtcNow.AddHours(1) }
        }
    }
    It 'connects with the delegated token and the initial domain, after closing other sessions' {
        $r = Connect-MspExchangeOnline -TenantId 'fabrikam.onmicrosoft.com'
        $r.TenantId | Should -Be $customer
        $r.Organization | Should -Be 'fabrikam.onmicrosoft.com'
        $r.ConnectionId | Should -Be 'new'
        $r.Mode | Should -Be 'Delegated'
        Should -Invoke Disconnect-ExchangeOnline -Times 1
        Should -Invoke Connect-ExchangeOnline -Times 1 -ParameterFilter {
            $AccessToken -eq 'token-for-https://outlook.office365.com' -and $DelegatedOrganization -eq 'fabrikam.onmicrosoft.com' -and $ShowBanner -eq $false
        }
    }
    It 'closes the session and throws when Exchange lands in another tenant' {
        $global:MspTest.State.NewTenant = '88888888-8888-8888-8888-888888888888'
        { Connect-MspExchangeOnline -TenantId $customer } | Should -Throw '*not 33333333*'
        Should -Invoke Disconnect-ExchangeOnline -ParameterFilter { $ConnectionId -eq 'new' }
    }
    It 'refuses an Exchange module that needs a newer PowerShell' {
        Mock Assert-MspModuleAvailable { [version]'3.10.1' }
        Mock Test-MspExoCompatibility { 'needs PowerShell 7.6' }
        { Connect-MspExchangeOnline -TenantId $customer } | Should -Throw '*7.6*'
        Should -Invoke Connect-ExchangeOnline -Times 0
    }
    It 'falls back to the tenant ID for -DelegatedOrganization when the initial domain cannot be read' {
        Mock Get-MspInitialDomain { throw 'Graph returned 403 Forbidden' }
        # The caught 403 from the domain lookup is recorded too, so look only for validation errors.
        $r = Connect-MspExchangeOnline -TenantId $customer -WarningVariable warnings -WarningAction SilentlyContinue -ErrorVariable errors
        @($errors | Where-Object { $_.FullyQualifiedErrorId -match 'Validation' -or $_.Exception -is [System.Management.Automation.ValidationMetadataException] }).Count | Should -Be 0
        $r.TenantId | Should -Be $customer
        $r.Organization | Should -Be $customer
        $r.ConnectionId | Should -Be 'new'
        ($warnings -join ' ') | Should -Match 'Using the tenant ID'
        Should -Invoke Connect-ExchangeOnline -Times 1 -Exactly -ParameterFilter {
            $DelegatedOrganization -eq $customer -and $AccessToken -eq 'token-for-https://outlook.office365.com'
        }
    }
    It 'still refuses an -Organization value that is not an onmicrosoft domain' {
        { Connect-MspExchangeOnline -TenantId $customer -Organization $customer } | Should -Throw -ErrorId 'ParameterArgumentValidationError*'
        Should -Invoke Connect-ExchangeOnline -Times 0
    }
    It 'requires -Organization for app-only when the initial domain cannot be read' {
        Mock Get-MspInitialDomain { throw 'Graph returned 403 Forbidden' }
        { Connect-MspExchangeOnline -TenantId $customer -AppOnly -AppId $global:MspTest.AutomationAppId -CertificateThumbprint ('B' * 40) } | Should -Throw '*Pass -Organization*'
        Should -Invoke Connect-ExchangeOnline -Times 0
    }
    It 'connects app-only with a certificate and the initial domain' {
        $r = Connect-MspExchangeOnline -TenantId $customer -AppOnly -AppId $global:MspTest.AutomationAppId -CertificateThumbprint ('B' * 40)
        $r.Mode | Should -Be 'AppOnly'
        Should -Invoke Connect-ExchangeOnline -Times 1 -ParameterFilter { $AppId -eq $global:MspTest.AutomationAppId -and $Organization -eq 'fabrikam.onmicrosoft.com' -and -not $AccessToken }
    }
}

Describe 'Connect-MspSecurityCompliance' {
    It 'is no longer experimental and uses -Organization with a Security and Compliance token' {
        Mock Assert-MspModuleAvailable { [version]'3.9.2' }
        $global:MspTest.State = @{ Sessions = @() }
        Mock Get-ConnectionInformation { $global:MspTest.State.Sessions }
        Mock Connect-IPPSSession { $global:MspTest.State.Sessions = @([pscustomobject]@{ ConnectionId = 'scc'; State = 'Connected'; TenantID = $customer; ConnectionUri = 'https://ps.compliance.protection.outlook.com'; IsEopSession = $true }) }
        $r = Connect-MspSecurityCompliance -TenantId $customer -WarningVariable warnings
        $r.Experimental | Should -BeFalse
        $r.Mode | Should -Be 'SecurityComplianceDelegated'
        @($warnings).Count | Should -Be 0
        Should -Invoke Connect-IPPSSession -Times 1 -ParameterFilter { $Organization -eq 'fabrikam.onmicrosoft.com' -and $AccessToken -eq 'token-for-https://ps.compliance.protection.outlook.com' }
    }
    It 'only accepts Microsoft login endpoints for -AzureADAuthorizationEndpointUri' {
        { Connect-MspSecurityCompliance -TenantId $customer -UseDelegatedOrganization -AzureADAuthorizationEndpointUri 'https://login.contoso.example/x' -WarningAction SilentlyContinue } | Should -Throw
    }
}

Describe 'Partner tenant refusal on customer-only connections' {
    It 'refuses <Name> against the partner tenant' -ForEach @(
        @{ Name = 'Connect-MspExchangeOnline -AppOnly'; Script = { Connect-MspExchangeOnline -TenantId $global:MspTest.PartnerTenantId -AppOnly -AppId $global:MspTest.AutomationAppId -CertificateThumbprint ('B' * 40) -Organization 'contoso.onmicrosoft.com' } }
        @{ Name = 'Connect-MspExchangeOnline'; Script = { Connect-MspExchangeOnline -TenantId $global:MspTest.PartnerTenantId } }
        @{ Name = 'Connect-MspGraph'; Script = { Connect-MspGraph -TenantId $global:MspTest.PartnerTenantId } }
        @{ Name = 'Connect-MspTeams'; Script = { Connect-MspTeams -TenantId $global:MspTest.PartnerTenantId } }
        @{ Name = 'Connect-MspSecurityCompliance'; Script = { Connect-MspSecurityCompliance -TenantId $global:MspTest.PartnerTenantId -WarningAction SilentlyContinue } }
    ) {
        Mock Assert-MspModuleAvailable { [version]'3.9.2' }
        Mock Test-MspExoCompatibility { $null }
        Mock Connect-ExchangeOnline {}
        Mock Connect-MgGraph {}
        Mock Connect-MicrosoftTeams {}
        Mock Connect-IPPSSession {}
        $Script | Should -Throw '*partner tenant*'
        Should -Invoke Connect-ExchangeOnline -Times 0
        Should -Invoke Connect-MgGraph -Times 0
        Should -Invoke Connect-MicrosoftTeams -Times 0
        Should -Invoke Connect-IPPSSession -Times 0
    }
}

Describe 'Connect-MspGraph' {
    BeforeEach { Mock Assert-MspModuleAvailable { [version]'2.30.0' } }
    It 'passes a SecureString token and checks the tenant' {
        Mock Connect-MgGraph {}
        Mock Get-MgContext { [pscustomobject]@{ TenantId = $customer; Account = 'tech'; Scopes = @(); AuthType = 'UserProvidedAccessToken' } }
        $r = Connect-MspGraph -TenantId $customer -WarningAction SilentlyContinue
        $r.TenantId | Should -Be $customer
        Should -Invoke Connect-MgGraph -Times 1 -ParameterFilter { $AccessToken -is [securestring] -and $NoWelcome }
    }
    It 'disconnects and throws on a tenant mismatch' {
        Mock Connect-MgGraph {}
        Mock Disconnect-MgGraph {}
        Mock Get-MgContext { [pscustomobject]@{ TenantId = '88888888-8888-8888-8888-888888888888' } }
        { Connect-MspGraph -TenantId $customer -WarningAction SilentlyContinue } | Should -Throw '*not 33333333*'
        Should -Invoke Disconnect-MgGraph -Times 1
    }
}

Describe 'Connect-MspTeams' {
    It 'passes Graph and Teams admin tokens and checks the tenant' {
        Mock Assert-MspModuleAvailable { [version]'7.0.0' }
        Mock Connect-MicrosoftTeams { [pscustomobject]@{ TenantId = $customer; Account = 'tech' } }
        $r = Connect-MspTeams -TenantId $customer
        $r.Verified | Should -BeTrue
        Should -Invoke Connect-MicrosoftTeams -Times 1 -ParameterFilter {
            $AccessTokens.Count -eq 2 -and $AccessTokens[0] -eq 'token-for-https://graph.microsoft.com' -and $AccessTokens[1] -eq 'token-for-48ac35b8-9aa8-4d74-927d-1f4a14a0b239'
        }
    }
}
