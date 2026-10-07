function Get-MspConfiguration {
    <#
    .SYNOPSIS
        Shows the MspGdap configuration. Contains no secrets.
    .DESCRIPTION
        Returns the non-secret settings from the configuration file plus
        session details: whether a session-only certificate password is set
        and which technician is selected.
    .PARAMETER Path
        Read and use this configuration file for the session instead of the default.
    .EXAMPLE
        Get-MspConfiguration
    .EXAMPLE
        Get-MspConfiguration -Path ./lab-config.json
    #>
    [CmdletBinding()]
    [OutputType('MspGdap.Configuration')]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )
    if ($Path) {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if ($resolvedPath -ne $script:MspState.ConfigPath) {
            $script:MspState.ConfigPath = $resolvedPath
            $script:MspState.Config = $null
        }
    }
    $config = Get-MspConfigurationInternal
    [pscustomobject]@{
        PSTypeName                 = 'MspGdap.Configuration'
        PartnerTenantId            = $config.PartnerTenantId
        AppId                      = $config.AppId
        CredentialType             = $config.CredentialType
        CertificateThumbprint      = $config.CertificateThumbprint
        CertificateStoreLocation   = $config.CertificateStoreLocation
        CertificatePath            = $config.CertificatePath
        SigningAlgorithm           = $config.SigningAlgorithm
        VaultName                  = $config.VaultName
        LoopbackPort               = $config.LoopbackPort
        TechnicianUpn              = $config.TechnicianUpn
        SessionTechnicianUpn       = $script:MspState.CurrentUpn
        SessionCertificatePassword = [bool]$script:MspState.CertificatePassword
        SessionCertificate         = if ($script:MspState.SessionCertificate) { $script:MspState.SessionCertificate.Thumbprint } else { $null }
        CachedTokenCount           = $script:MspTokenCache.Count
        Path                       = $script:MspState.ConfigPath
        FileExists                 = Test-Path -LiteralPath $script:MspState.ConfigPath -PathType Leaf
    }
}
