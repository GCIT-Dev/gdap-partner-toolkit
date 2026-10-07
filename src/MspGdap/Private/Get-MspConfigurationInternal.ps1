function Get-MspConfigurationInternal {
    <#
    .SYNOPSIS
        Returns the active configuration object, loading it from disk once per session.
    .PARAMETER RequireComplete
        Throw when the settings needed to request tokens are missing.
    .EXAMPLE
        $config = Get-MspConfigurationInternal -RequireComplete
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [switch]$RequireComplete
    )
    if ($null -eq $script:MspState.Config) {
        $script:MspState.Config = Read-MspConfigurationFile -Path $script:MspState.ConfigPath
    }
    $config = $script:MspState.Config

    if ($RequireComplete) {
        $missing = [System.Collections.Generic.List[string]]::new()
        if (-not $config.PartnerTenantId) { $missing.Add('PartnerTenantId') }
        if (-not $config.AppId) { $missing.Add('AppId') }
        if (-not $config.VaultName) { $missing.Add('VaultName') }
        if ($config.CredentialType -eq 'Certificate' -and -not $config.CertificateThumbprint -and -not $config.CertificatePath -and -not $script:MspState.SessionCertificate) {
            $missing.Add('CertificateThumbprint, CertificatePath or a session -Certificate')
        }
        if ($missing.Count -gt 0) {
            throw (New-MspErrorRecord -Message "MspGdap is not fully configured. Missing: $($missing -join ', '). Run Set-MspConfiguration." -ErrorId 'MspGdap.Configuration.Incomplete' -Category NotSpecified -TargetObject $script:MspState.ConfigPath)
        }
    }
    $config
}
