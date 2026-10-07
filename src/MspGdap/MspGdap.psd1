@{
    RootModule           = 'MspGdap.psm1'
    ModuleVersion        = '0.2.1'
    GUID                 = '6f0f3a4e-2b8d-4c1e-9a57-3d6b1e0c8f21'
    Author               = 'MspGdap contributors'
    CompanyName          = 'Community'
    Copyright            = '(c) 2026 GCIT Pty Ltd. MIT licence.'
    Description          = 'GDAP and Secure Application Model toolkit for Microsoft partners and MSPs. Register per-technician delegated refresh tokens securely, pre-consent a multi-tenant partner app into GDAP customers and get cached, validated access tokens per customer and resource. Based on the methods GCIT (gcit.com.au) uses to manage customer tenants.'
    PowerShellVersion    = '7.4'
    CompatiblePSEditions = @('Core')

    # Explicit list. Add new public functions here as well as in Public/.
    FunctionsToExport    = @(
        'Add-MspPartnerAppCertificate'
        'Clear-MspTokenCache'
        'Connect-MspExchangeOnline'
        'Connect-MspGraph'
        'Connect-MspSecurityCompliance'
        'Connect-MspTeams'
        'Disconnect-Msp'
        'Enable-MspExchangeAppAccess'
        'Get-MspAccessToken'
        'Get-MspAuthHeader'
        'Get-MspConfiguration'
        'Get-MspCustomer'
        'Get-MspGdapRelationship'
        'Grant-MspPartnerAppConsent'
        'Invoke-MspGraphBatch'
        'Invoke-MspGraphRequest'
        'New-MspGdapRelationship'
        'New-MspPartnerApp'
        'Register-MspPartnerToken'
        'Remove-MspPartnerAppConsent'
        'Resolve-MspTenantId'
        'Set-MspConfiguration'
        'Set-MspGdapAccessAssignment'
        'Test-MspAccessToken'
        'Test-MspExchangeAppAccess'
        'Test-MspGdapAccess'
        'Test-MspPartnerAppConsent'
        'Unregister-MspPartnerToken'
    )
    CmdletsToExport      = @()
    VariablesToExport    = @()
    AliasesToExport      = @()

    PrivateData          = @{
        PSData = @{
            Tags         = @('GDAP', 'PartnerCenter', 'CSP', 'MSP', 'MicrosoftGraph', 'EntraID', 'SecureApplicationModel', 'ExchangeOnline')
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            ReleaseNotes = 'Initial preview. See CHANGELOG.md.'
        }
    }
}
