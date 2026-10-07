# MspGdap root module.
#
# All state lives in module scope ($script:). Nothing is written to $global:.
# Access tokens are held as SecureString in the cache. Refresh tokens are never
# held in module state at all: they are read from SecretManagement when needed
# and the rotated token is written straight back.

# Fixed endpoints and constants (Microsoft commercial cloud).
$script:MspConstants = @{
    Authority           = 'https://login.microsoftonline.com'
    GraphBaseUri        = 'https://graph.microsoft.com'
    ClientAssertionType = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
    OidcScopes          = @('openid', 'profile', 'email', 'offline_access')
    DefaultSkewMinutes  = 5
    GuidPattern         = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
}

# Token cache. Key: "<tenantId>|<resource>|<appId>" (lower case).
$script:MspTokenCache = [hashtable]::Synchronized(@{})

# Session state. Secrets are never persisted from here.
$script:MspState = @{
    Config              = $null
    ConfigPath          = $null
    CurrentUpn          = $null
    CertificatePassword = $null
    Certificate         = $null
    CertificateSource   = $null
    SessionCertificate  = $null
    TenantIdCache       = [hashtable]::Synchronized(@{})
    Connections         = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    ExchangeConnections = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
}

$defaultConfigPath = if ($env:MSPGDAP_CONFIG_PATH) {
    $env:MSPGDAP_CONFIG_PATH
}
else {
    Join-Path -Path $HOME -ChildPath '.mspgdap' -AdditionalChildPath 'config.json'
}
$script:MspState.ConfigPath = $defaultConfigPath

$privateFiles = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -File -ErrorAction SilentlyContinue)
$publicFiles = @(Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public') -Filter '*.ps1' -File -ErrorAction SilentlyContinue)

foreach ($file in @($privateFiles + $publicFiles)) {
    try {
        . $file.FullName
    }
    catch {
        throw "MspGdap: failed to load $($file.Name): $($_.Exception.Message)"
    }
}

Export-ModuleMember -Function $publicFiles.BaseName

# Clear in-memory tokens and certificate handles if the module is removed.
$ExecutionContext.SessionState.Module.OnRemove = {
    $script:MspTokenCache.Clear()
    if ($script:MspState.Certificate) {
        try { $script:MspState.Certificate.Dispose() } catch { Write-Verbose 'Certificate handle already released.' }
    }
    $script:MspState.Certificate = $null
    $script:MspState.SessionCertificate = $null
    $script:MspState.CertificatePassword = $null
}
