function Set-MspConfiguration {
    <#
    .SYNOPSIS
        Sets the MspGdap configuration (partner tenant, partner app, credential, vault).
    .DESCRIPTION
        Settings are stored as non-secret JSON in $HOME/.mspgdap/config.json
        (or the path in the MSPGDAP_CONFIG_PATH environment variable, or -Path).
        Only the parameters you pass are changed.

        Nothing secret is written to that file:
        - -CertificatePassword is kept for the current session only.
        - -Certificate (an X509Certificate2 with its private key, for example
          loaded from Azure Key Vault in an Azure Function) is kept for the
          current session only and takes precedence over -CertificateThumbprint
          and -CertificatePath. MspGdap never disposes or exports it.
        - -ClientSecret is written to the SecretManagement vault, never to the
          file. Client secrets are supported but a certificate is strongly
          preferred.

        -WhatIf changes nothing at all: no file write, no vault write and no
        change to the session (token cache, loaded certificate, password).
    .PARAMETER PartnerTenantId
        Your partner (CSP) tenant GUID.
    .PARAMETER AppId
        Application (client) ID of your multi-tenant partner app.
    .PARAMETER CertificateThumbprint
        Thumbprint of the partner app certificate in the certificate store (preferred).
    .PARAMETER CertificateStoreLocation
        CurrentUser (default) or LocalMachine.
    .PARAMETER CertificatePath
        Path to a PFX file holding the partner app certificate.
    .PARAMETER CertificatePassword
        PFX password for this session only. Never saved.
    .PARAMETER Certificate
        Partner app certificate object with a private key, for this session only. Never saved.
    .PARAMETER ClientSecret
        Partner app client secret, stored in the vault. Discouraged.
    .PARAMETER SigningAlgorithm
        PS256 (default, Microsoft's current guidance) or RS256. PS256 falls back
        to RS256 automatically for legacy CSP keys that cannot sign with PSS.
    .PARAMETER VaultName
        SecretManagement vault that holds refresh tokens (for example an
        Az.KeyVault vault with per-technician RBAC, or SecretStore).
    .PARAMETER LoopbackPort
        Port for the sign-in redirect listener. 0 (default) picks a free port.
    .PARAMETER TechnicianUpn
        Default technician (partner admin account) whose refresh token is used.
    .PARAMETER Path
        Configuration file to write and use for this session.
    .PARAMETER PassThru
        Return the resulting configuration.
    .EXAMPLE
        Set-MspConfiguration -PartnerTenantId '<PartnerTenantId>' -AppId '<PartnerAppId>' -CertificateThumbprint '<Thumbprint>' -VaultName 'MspGdapVault'
    .EXAMPLE
        Set-MspConfiguration -TechnicianUpn 'tech-admin@contoso.onmicrosoft.com' -PassThru
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType('MspGdap.Configuration')]
    param(
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$PartnerTenantId,

        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$AppId,

        [ValidatePattern('^[0-9a-fA-F]{40}$')]
        [string]$CertificateThumbprint,

        [ValidateSet('CurrentUser', 'LocalMachine')]
        [string]$CertificateStoreLocation,

        [ValidateNotNullOrEmpty()]
        [string]$CertificatePath,

        [securestring]$CertificatePassword,

        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        [securestring]$ClientSecret,

        [ValidateSet('PS256', 'RS256')]
        [string]$SigningAlgorithm,

        [ValidateNotNullOrEmpty()]
        [string]$VaultName,

        [ValidateRange(0, 65535)]
        [int]$LoopbackPort,

        [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
        [string]$TechnicianUpn,

        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [switch]$PassThru
    )

    if ($CertificateThumbprint -and $CertificatePath) {
        throw (New-MspErrorRecord -Message 'Use either -CertificateThumbprint or -CertificatePath, not both.' -ErrorId 'MspGdap.Configuration.Conflict' -Category InvalidArgument)
    }
    if ($Certificate -and ($CertificateThumbprint -or $CertificatePath -or $ClientSecret)) {
        throw (New-MspErrorRecord -Message 'Use -Certificate on its own, without -CertificateThumbprint, -CertificatePath or -ClientSecret.' -ErrorId 'MspGdap.Configuration.Conflict' -Category InvalidArgument)
    }
    if ($Certificate -and -not $Certificate.HasPrivateKey) {
        throw (New-MspErrorRecord -Message "Certificate $($Certificate.Thumbprint) has no private key, so it cannot sign client assertions." -ErrorId 'MspGdap.Certificate.NoPrivateKey' -Category InvalidArgument -TargetObject $Certificate.Thumbprint)
    }
    if ($ClientSecret -and ($CertificateThumbprint -or $CertificatePath)) {
        throw (New-MspErrorRecord -Message 'Use either a certificate or -ClientSecret, not both.' -ErrorId 'MspGdap.Configuration.Conflict' -Category InvalidArgument)
    }

    $configPath = $script:MspState.ConfigPath
    $current = $null
    if ($Path) {
        $resolvedPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        if ($resolvedPath -ne $script:MspState.ConfigPath) {
            # Read the other file without switching the session yet (that only happens when the write is approved).
            $configPath = $resolvedPath
            $current = Read-MspConfigurationFile -Path $resolvedPath
        }
    }
    if ($null -eq $current) { $current = Get-MspConfigurationInternal }

    $updated = [ordered]@{}
    foreach ($property in $current.PSObject.Properties) { $updated[$property.Name] = $property.Value }

    foreach ($name in 'PartnerTenantId', 'AppId', 'TechnicianUpn') {
        if ($PSBoundParameters.ContainsKey($name)) { $updated[$name] = ([string]$PSBoundParameters[$name]).ToLowerInvariant() }
    }
    foreach ($name in 'CertificateStoreLocation', 'SigningAlgorithm', 'VaultName', 'LoopbackPort') {
        if ($PSBoundParameters.ContainsKey($name)) { $updated[$name] = $PSBoundParameters[$name] }
    }
    if ($CertificateThumbprint) {
        $updated.CertificateThumbprint = $CertificateThumbprint.ToUpperInvariant()
        $updated.CertificatePath = $null
        $updated.CredentialType = 'Certificate'
    }
    if ($CertificatePath) {
        $updated.CertificatePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CertificatePath)
        $updated.CertificateThumbprint = $null
        $updated.CredentialType = 'Certificate'
    }
    $secretName = $null
    if ($ClientSecret) {
        if (-not $updated.AppId -or -not $updated.VaultName) {
            throw (New-MspErrorRecord -Message 'Set -AppId and -VaultName before storing a client secret.' -ErrorId 'MspGdap.Configuration.Incomplete' -Category InvalidArgument)
        }
        Write-Warning 'Client secrets are supported but discouraged. A certificate credential (Set-MspConfiguration -CertificateThumbprint) is strongly preferred.'
        Assert-MspSecretManagement -VaultName $updated.VaultName
        $secretName = Get-MspClientSecretName -AppId $updated.AppId
        $updated.CredentialType = 'ClientSecret'
        $updated.CertificateThumbprint = $null
        $updated.CertificatePath = $null
    }
    if ($Certificate) {
        $updated.CredentialType = 'Certificate'
    }

    # One decision covers the whole change, so the file, the vault and the session never disagree.
    $action = if ($secretName) { "Write MspGdap configuration and store the partner app client secret as $secretName in vault '$($updated.VaultName)'" } else { 'Write MspGdap configuration' }
    if (-not $PSCmdlet.ShouldProcess($configPath, $action)) {
        return
    }

    if ($secretName) {
        Set-Secret -Name $secretName -SecureStringSecret $ClientSecret -Vault $updated.VaultName -ErrorAction Stop
    }

    $directory = Split-Path -Path $configPath -Parent
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        $null = New-Item -ItemType Directory -Path $directory -Force
    }
    $json = [pscustomobject]$updated | ConvertTo-Json -Depth 4
    $temporary = "$configPath.tmp"
    Set-Content -LiteralPath $temporary -Value $json -Encoding utf8NoBOM -ErrorAction Stop
    Move-Item -LiteralPath $temporary -Destination $configPath -Force -ErrorAction Stop
    $script:MspState.ConfigPath = $configPath
    $script:MspState.Config = [pscustomobject]$updated
    Write-Verbose "Configuration written to $configPath."

    if ($PSBoundParameters.ContainsKey('CertificatePassword')) {
        $script:MspState.CertificatePassword = $CertificatePassword
        Write-Verbose 'Certificate password set for this session only. It is not saved.'
    }

    $credentialChanged = $CertificateThumbprint -or $CertificatePath -or $ClientSecret -or $Certificate -or $PSBoundParameters.ContainsKey('CertificateStoreLocation') -or $PSBoundParameters.ContainsKey('CertificatePassword')
    if ($credentialChanged -and $script:MspState.Certificate) {
        try { $script:MspState.Certificate.Dispose() } catch { Write-Verbose 'Certificate handle already released.' }
        $script:MspState.Certificate = $null
        $script:MspState.CertificateSource = $null
    }
    if ($CertificateThumbprint -or $CertificatePath -or $ClientSecret) {
        # A persisted credential replaces any session-only certificate.
        $script:MspState.SessionCertificate = $null
    }
    if ($Certificate) {
        $script:MspState.SessionCertificate = $Certificate
        Write-Verbose "Certificate $($Certificate.Thumbprint) set for this session only. It is not saved."
    }
    if ($PSBoundParameters.ContainsKey('AppId') -or $PSBoundParameters.ContainsKey('PartnerTenantId') -or $credentialChanged) {
        $script:MspTokenCache.Clear()
    }

    if ($PassThru) {
        Get-MspConfiguration
    }
}
