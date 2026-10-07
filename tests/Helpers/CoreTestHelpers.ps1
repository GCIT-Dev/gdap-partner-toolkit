# Shared helpers for the MspGdap core tests. Offline only: every HTTP call and
# every SecretManagement call is mocked. No real tenant is contacted.

$script:ModuleManifestPath = Join-Path -Path $PSScriptRoot -ChildPath '..' -AdditionalChildPath '..', 'src', 'MspGdap', 'MspGdap.psd1'

# Placeholder identifiers (not real tenants or apps).
$script:TestPartnerTenantId = '11111111-1111-4111-8111-111111111111'
$script:TestCustomerTenantId = '22222222-2222-4222-8222-222222222222'
$script:TestOtherTenantId = '33333333-3333-4333-8333-333333333333'
$script:TestAppId = '44444444-4444-4444-8444-444444444444'
$script:TestUpn = 'tech-admin@contoso.onmicrosoft.com'

# Stub SecretManagement commands when the module is not installed, so Mock can bind to them.
if (-not (Get-Command -Name 'Get-Secret' -ErrorAction SilentlyContinue)) {
    Import-Module -Name 'Microsoft.PowerShell.SecretManagement' -ErrorAction SilentlyContinue
}
if (-not (Get-Command -Name 'Get-Secret' -ErrorAction SilentlyContinue)) {
    function global:Get-Secret { param($Name, $Vault, [switch]$AsPlainText) }
    function global:Set-Secret { param($Name, $SecureStringSecret, $Secret, $Vault, $Metadata) }
    function global:Get-SecretInfo { param($Name, $Vault) }
    function global:Get-SecretVault { param($Name) }
}

function ConvertTo-TestBase64Url {
    param([byte[]]$Bytes)
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function New-TestJwt {
    param(
        [Parameter(Mandatory)][hashtable]$Claims,
        [hashtable]$Header = @{ alg = 'RS256'; typ = 'JWT' }
    )
    $h = ConvertTo-TestBase64Url ([Text.Encoding]::UTF8.GetBytes(($Header | ConvertTo-Json -Compress)))
    $c = ConvertTo-TestBase64Url ([Text.Encoding]::UTF8.GetBytes(($Claims | ConvertTo-Json -Compress -Depth 5)))
    "$h.$c.c2lnbmF0dXJl"
}

function New-TestSecure {
    param([string]$Value)
    [System.Net.NetworkCredential]::new('', $Value).SecurePassword
}

function ConvertFrom-TestSecure {
    param([securestring]$Value)
    [System.Net.NetworkCredential]::new('', $Value).Password
}

function New-TestHttpError {
    param(
        [Parameter(Mandatory)][int]$Status,
        [string]$Json,
        [int]$RetryAfterSeconds
    )
    $message = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Status)
    if ($RetryAfterSeconds) {
        $message.Headers.RetryAfter = [System.Net.Http.Headers.RetryConditionHeaderValue]::new([timespan]::FromSeconds($RetryAfterSeconds))
    }
    $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Status.", $message)
    $record = [System.Management.Automation.ErrorRecord]::new($exception, 'WebCmdletWebResponseException,Microsoft.PowerShell.Commands.InvokeRestMethodCommand', 'InvalidOperation', $null)
    if ($Json) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Json) }
    $record
}

function New-TestTokenResponse {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [string]$Audience = 'https://graph.microsoft.com',
        [string]$RefreshToken,
        [int]$ExpiresIn = 3600,
        [string]$Scope = 'https://graph.microsoft.com/User.Read.All https://graph.microsoft.com/Directory.Read.All',
        [string[]]$Amr = @('pwd', 'mfa')
    )
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $claims = @{
        aud = $Audience
        tid = $TenantId
        exp = $now + $ExpiresIn
        nbf = $now
        scp = (($Scope -split ' ') | ForEach-Object { ($_ -split '/')[-1] }) -join ' '
        upn = $script:TestUpn
        amr = $Amr
    }
    $response = [ordered]@{
        token_type   = 'Bearer'
        access_token = New-TestJwt -Claims $claims
        expires_in   = $ExpiresIn
        scope        = $Scope
    }
    if ($RefreshToken) { $response.refresh_token = $RefreshToken }
    [pscustomobject]$response
}

function Initialize-TestModuleState {
    param([string]$ConfigPath = (Join-Path ([System.IO.Path]::GetTempPath()) "mspgdap-test-$([guid]::NewGuid()).json"))
    InModuleScope MspGdap -Parameters @{
        Partner    = $script:TestPartnerTenantId
        App        = $script:TestAppId
        Upn        = $script:TestUpn
        ConfigPath = $ConfigPath
    } {
        param($Partner, $App, $Upn, $ConfigPath)
        $script:MspTokenCache.Clear()
        $script:MspState.ConfigPath = $ConfigPath
        $script:MspState.Config = [pscustomobject]@{
            ConfigVersion            = 1
            PartnerTenantId          = $Partner
            AppId                    = $App
            CredentialType           = 'Certificate'
            CertificateThumbprint    = ('A' * 40)
            CertificateStoreLocation = 'CurrentUser'
            CertificatePath          = $null
            SigningAlgorithm         = 'PS256'
            VaultName                = 'TestVault'
            LoopbackPort             = 0
            TechnicianUpn            = $Upn
        }
        $script:MspState.CurrentUpn = $null
        $script:MspState.Certificate = $null
        $script:MspState.SessionCertificate = $null
        $script:MspState.CertificatePassword = $null
        $script:MspState.CertificateSource = $null
        $script:MspState.TenantIdCache.Clear()
        $script:MspState.Connections.Clear()
        $script:MspState.ExchangeConnections.Clear()
    }
}
