# Test support for the setup, consent, GDAP and Exchange commands.
# Loads only those source files, plus stubs for the core functions and the optional Microsoft modules,
# so every test runs offline with mocks. Nothing here contacts a tenant.

$moduleRoot = Join-Path (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSCommandPath))) (Join-Path 'src' 'MspGdap')

$setupConsentFiles = @(
    'Private/New-MspStepResult.ps1', 'Private/New-MspOperationResult.ps1', 'Private/Get-MspHttpStatusCode.ps1',
    'Private/Invoke-MspGraphCall.ps1', 'Private/Invoke-MspGraphDirect.ps1', 'Private/Invoke-MspPartnerCenterRequest.ps1',
    'Private/ConvertTo-MspPlainToken.ps1', 'Private/ConvertTo-MspSecureToken.ps1', 'Private/Get-MspInitialDomain.ps1',
    'Private/Get-MspServicePrincipalByAppId.ps1', 'Private/Resolve-MspDataPath.ps1', 'Private/Read-MspPermissionManifest.ps1',
    'Private/Get-MspDesiredDelegatedGrant.ps1', 'Private/Get-MspGdapRoleCatalog.ps1', 'Private/Resolve-MspGdapRole.ps1',
    'Private/Read-MspGdapAccessMap.ps1', 'Private/Resolve-MspPartnerAppId.ps1', 'Private/Assert-MspModuleAvailable.ps1',
    'Private/Test-MspExoCompatibility.ps1', 'Private/Resolve-MspCertificate.ps1', 'Private/ConvertTo-MspKeyCredential.ps1',
    'Private/Test-MspKeyCredentialThumbprint.ps1', 'Private/New-MspAddKeyProof.ps1', 'Private/Invoke-MspPartnerGraph.ps1',
    'Private/Test-MspSetupContext.ps1', 'Private/ConvertTo-MspCustomerReference.ps1', 'Private/Wait-MspCondition.ps1',
    'Private/Get-MspCustomerConsentState.ps1', 'Private/Test-MspConsentMissingError.ps1', 'Private/Get-MspEmptyConsentState.ps1', 'Private/ConvertTo-MspGdapRelationship.ps1', 'Private/Get-MspExchangeAppAccessState.ps1',
    'Private/Resolve-MspExchangeAppRole.ps1', 'Private/New-MspErrorRecord.ps1', 'Private/Resolve-MspGraphUri.ps1',
    'Private/Resolve-MspCustomerTenant.ps1', 'Private/Get-MspWellKnownRoleId.ps1', 'Public/Remove-MspPartnerAppConsent.ps1',
    'Public/New-MspPartnerApp.ps1', 'Public/Add-MspPartnerAppCertificate.ps1', 'Public/Grant-MspPartnerAppConsent.ps1',
    'Public/Test-MspPartnerAppConsent.ps1', 'Public/Get-MspGdapRelationship.ps1', 'Public/New-MspGdapRelationship.ps1',
    'Public/Set-MspGdapAccessAssignment.ps1', 'Public/Test-MspGdapAccess.ps1', 'Public/Enable-MspExchangeAppAccess.ps1',
    'Public/Test-MspExchangeAppAccess.ps1', 'Public/Connect-MspExchangeOnline.ps1', 'Public/Connect-MspSecurityCompliance.ps1',
    'Public/Connect-MspGraph.ps1', 'Public/Connect-MspTeams.ps1'
)
foreach ($relative in $setupConsentFiles) { . (Join-Path $moduleRoot $relative) }

$global:MspTest = @{
    PartnerTenantId    = '11111111-1111-1111-1111-111111111111'
    PartnerAppId       = '22222222-2222-2222-2222-222222222222'
    CustomerTenantId   = '33333333-3333-3333-3333-333333333333'
    AutomationAppId    = '44444444-4444-4444-4444-444444444444'
    ModuleRoot         = $moduleRoot
    RepoRoot           = Split-Path -Parent (Split-Path -Parent $moduleRoot)
    Routes             = @()
    Calls              = New-Object System.Collections.Generic.List[object]
    State              = @{}
}

# ---------- Stubs for core functions (replaced by mocks in tests) ----------
function Invoke-MspGraphRequest {
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Tenant')]
    param(
        [Parameter(ParameterSetName = 'Tenant')][string]$TenantId,
        [Parameter(ParameterSetName = 'Partner')][switch]$PartnerTenant,
        [string]$Uri, [string]$Method = 'GET', [object]$Body, [string]$ApiVersion = 'v1.0', [hashtable]$Headers
    )
    Invoke-FakeGraph -Method $Method -Uri $Uri -Body $Body -TenantId $TenantId -PartnerTenant:$PartnerTenant
}
function Get-MspAuthHeader {
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    param([Parameter(ParameterSetName = 'Tenant')][string]$TenantId, [Parameter(ParameterSetName = 'Partner')][switch]$PartnerTenant, [string]$Resource)
    @{ Authorization = 'Bearer test-token' }
}
function Get-MspAccessToken {
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    param([Parameter(ParameterSetName = 'Tenant')][string]$TenantId, [Parameter(ParameterSetName = 'Partner')][switch]$PartnerTenant, [string]$Resource)
    $secure = New-Object System.Security.SecureString
    foreach ($c in "token-for-$Resource".ToCharArray()) { $secure.AppendChar($c) }
    [pscustomobject]@{ PSTypeName = 'MspGdap.AccessToken'; TenantId = $TenantId; Resource = $Resource; ExpiresOn = [datetime]::UtcNow.AddMinutes(60); AccessToken = $secure }
}
function Resolve-MspTenantId {
    [CmdletBinding()]
    param([Parameter(Mandatory)][Alias('TenantId')][string]$Tenant)
    if ($Tenant -match '^[0-9a-fA-F-]{36}$') { return $Tenant.ToLowerInvariant() }
    if ($Tenant -eq 'fabrikam.onmicrosoft.com') { return $global:MspTest.CustomerTenantId }
    throw "Unknown test tenant $Tenant"
}
function Get-MspCustomer { [CmdletBinding()] param() throw 'Get-MspCustomer not mocked' }
function Get-MspConfiguration {
    [CmdletBinding()] param()
    [pscustomobject]@{ PartnerTenantId = $global:MspTest.PartnerTenantId; AppId = $global:MspTest.PartnerAppId }
}

# ---------- Stubs for optional Microsoft modules ----------
function Get-MgContext { [CmdletBinding()] param() $null }
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method = 'GET', [string]$Uri, [object]$Body, [string]$ContentType, [string]$OutputType)
    $parsed = if ($Body -is [string] -and $Body) { $Body | ConvertFrom-Json } else { $Body }
    $result = Invoke-FakeGraph -Method $Method -Uri $Uri -Body $parsed -Transport 'MgGraph'
    $result
}
function Connect-MgGraph { [CmdletBinding()] param([securestring]$AccessToken, [switch]$NoWelcome) }
function Disconnect-MgGraph { [CmdletBinding()] param() }
function Connect-ExchangeOnline {
    [CmdletBinding()]
    param([string]$AccessToken, [string]$DelegatedOrganization, [string]$Organization, [string]$AppId, [string]$CertificateThumbprint,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate, [switch]$ShowBanner, [string]$Prefix, [string[]]$CommandName)
}
function Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param([string]$ConnectionId) }
function Get-ConnectionInformation { [CmdletBinding()] param() @() }
function Connect-IPPSSession {
    [CmdletBinding()]
    param([string]$AccessToken, [string]$Organization, [string]$DelegatedOrganization, [string]$AzureADAuthorizationEndpointUri, [switch]$ShowBanner, [string]$Prefix)
}
function Get-OrganizationConfig { [CmdletBinding()] param() [pscustomobject]@{ Name = 'fabrikam' } }
function Get-MspGdapTestOrganizationConfig { [CmdletBinding()] param() [pscustomobject]@{ Name = 'fabrikam' } }
function Connect-MicrosoftTeams { [CmdletBinding()] param([string[]]$AccessTokens) }
function Disconnect-MicrosoftTeams { [CmdletBinding()] param() }

# ---------- Fake Graph router ----------
function Set-FakeGraphRoute {
    param([object[]]$Route)
    $global:MspTest.Routes = @($Route)
    $global:MspTest.Calls.Clear()
}
function New-FakeRoute {
    param([string]$Method = 'GET', [Parameter(Mandatory)][string]$Pattern, [Parameter(Mandatory)][scriptblock]$Response, [switch]$Collection)
    [pscustomobject]@{ Method = $Method; Pattern = $Pattern; Response = $Response; Collection = [bool]$Collection }
}
function Invoke-FakeGraph {
    param([string]$Method, [string]$Uri, [object]$Body, [string]$TenantId, [switch]$PartnerTenant, [string]$Transport = 'MspGdap')
    $global:MspTest.Calls.Add([pscustomobject]@{ Method = $Method; Uri = [uri]::UnescapeDataString($Uri); Body = $Body; TenantId = $TenantId; PartnerTenant = [bool]$PartnerTenant; Transport = $Transport })
    $decoded = [uri]::UnescapeDataString($Uri)
    foreach ($route in $global:MspTest.Routes) {
        if ($route.Method -eq $Method -and $decoded -match $route.Pattern) {
            $result = & $route.Response $Body $decoded $TenantId
            if ($Transport -eq 'MgGraph' -and $route.Collection) {
                return [pscustomobject]@{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#test'; value = @($result) }
            }
            return $result
        }
    }
    throw "Unexpected Graph call: $Method $decoded"
}
function Get-FakeGraphCall {
    param([string]$Method, [string]$Pattern = '.')
    @($global:MspTest.Calls | Where-Object { (-not $Method -or $_.Method -eq $Method) -and $_.Uri -match $Pattern })
}
function New-FakeHttpError {
    param([int]$StatusCode, [string]$Message = 'Request failed')
    $exception = New-Object System.Exception ("Response status code does not indicate success: {0}. {1}" -f $StatusCode, $Message)
    $exception.Data['StatusCode'] = $StatusCode
    $exception
}

# ---------- Certificates (in memory, never stored) ----------
function New-TestCertificate {
    param([int]$KeySize = 2048, [datetime]$NotBefore = [datetime]::UtcNow.AddDays(-1), [datetime]$NotAfter = [datetime]::UtcNow.AddDays(365), [string]$Subject = 'CN=MspGdap Test')
    $rsa = [System.Security.Cryptography.RSA]::Create($KeySize)
    $request = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest($Subject, $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $request.CreateSelfSigned([DateTimeOffset]$NotBefore, [DateTimeOffset]$NotAfter)
}
function ConvertTo-TestKeyIdentifier {
    param([string]$Thumbprint)
    $bytes = for ($i = 0; $i -lt $Thumbprint.Length; $i += 2) { [Convert]::ToByte($Thumbprint.Substring($i, 2), 16) }
    [Convert]::ToBase64String([byte[]]$bytes)
}
