function Get-MspSigningCertificate {
    <#
    .SYNOPSIS
        Loads the partner app certificate used to sign client assertions.
    .DESCRIPTION
        Uses the certificate store (thumbprint, CurrentUser or LocalMachine My
        store) or a PFX file. A PFX password can only be supplied for the
        current session (Set-MspConfiguration -CertificatePassword) and is never
        written to disk. The loaded certificate is cached for the session.
        A session-only certificate set with Set-MspConfiguration -Certificate
        takes precedence and is used as given (never disposed by MspGdap).
        Prefer a store certificate with a non-exportable key.
    .PARAMETER Configuration
        The active configuration object.
    .EXAMPLE
        $cert = Get-MspSigningCertificate -Configuration $config
    #>
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Configuration
    )

    if ($script:MspState.SessionCertificate) {
        $session = $script:MspState.SessionCertificate
        if ($session.NotAfter.ToUniversalTime() -le [datetime]::UtcNow) {
            throw (New-MspErrorRecord -Message "Session certificate $($session.Thumbprint) expired on $($session.NotAfter.ToString('yyyy-MM-dd')). Load the renewed certificate and run Set-MspConfiguration -Certificate again." -ErrorId 'MspGdap.Certificate.Expired' -Category InvalidData -TargetObject $session.Thumbprint)
        }
        return $session
    }

    $source = if ($Configuration.CertificateThumbprint) {
        "store:$($Configuration.CertificateStoreLocation):$($Configuration.CertificateThumbprint)".ToLowerInvariant()
    }
    elseif ($Configuration.CertificatePath) {
        "file:$($Configuration.CertificatePath)"
    }
    else {
        throw (New-MspErrorRecord -Message 'No certificate is configured. Run Set-MspConfiguration -CertificateThumbprint (preferred) or -CertificatePath.' -ErrorId 'MspGdap.Certificate.NotConfigured' -Category NotSpecified)
    }

    if ($script:MspState.Certificate -and $script:MspState.CertificateSource -eq $source) {
        return $script:MspState.Certificate
    }

    $certificate = $null
    if ($Configuration.CertificateThumbprint) {
        $location = if ($Configuration.CertificateStoreLocation -eq 'LocalMachine') {
            [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
        }
        else {
            [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
        }
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', $location)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            $found = $store.Certificates.Find([System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $Configuration.CertificateThumbprint, $false)
            if ($found.Count -gt 0) { $certificate = $found[0] }
        }
        finally {
            $store.Close()
        }
        if (-not $certificate) {
            throw (New-MspErrorRecord -Message "Certificate $($Configuration.CertificateThumbprint) was not found in $location\My." -ErrorId 'MspGdap.Certificate.NotFound' -Category ObjectNotFound -TargetObject $Configuration.CertificateThumbprint)
        }
    }
    else {
        $path = $Configuration.CertificatePath
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw (New-MspErrorRecord -Message "Certificate file '$path' was not found." -ErrorId 'MspGdap.Certificate.NotFound' -Category ObjectNotFound -TargetObject $path)
        }
        $password = if ($script:MspState.CertificatePassword) { ConvertFrom-MspSecureString -SecureString $script:MspState.CertificatePassword } else { $null }
        $flagSets = @(
            [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet,
            [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet
        )
        $loaderType = 'System.Security.Cryptography.X509Certificates.X509CertificateLoader' -as [type]
        foreach ($flags in $flagSets) {
            try {
                $certificate = if ($loaderType) {
                    $loaderType::LoadPkcs12FromFile($path, $password, $flags)
                }
                else {
                    [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($path, $password, $flags)
                }
                break
            }
            catch {
                Write-Verbose "Loading the PFX with $flags failed. Trying the next key storage option."
                $certificate = $null
            }
        }
        $password = $null
        if (-not $certificate) {
            throw (New-MspErrorRecord -Message "Certificate file '$path' could not be loaded. Check the session password set with Set-MspConfiguration -CertificatePassword." -ErrorId 'MspGdap.Certificate.LoadFailed' -Category ReadError -TargetObject $path)
        }
    }

    if (-not $certificate.HasPrivateKey) {
        throw (New-MspErrorRecord -Message "Certificate $($certificate.Thumbprint) has no private key available to this user." -ErrorId 'MspGdap.Certificate.NoPrivateKey' -Category PermissionDenied -TargetObject $certificate.Thumbprint)
    }
    $now = [datetime]::UtcNow
    if ($certificate.NotAfter.ToUniversalTime() -le $now) {
        throw (New-MspErrorRecord -Message "Certificate $($certificate.Thumbprint) expired on $($certificate.NotAfter.ToString('yyyy-MM-dd')). Upload a new certificate to the partner app and update the configuration." -ErrorId 'MspGdap.Certificate.Expired' -Category InvalidData -TargetObject $certificate.Thumbprint)
    }
    if ($certificate.NotAfter.ToUniversalTime() -le $now.AddDays(30)) {
        Write-Warning "Partner app certificate $($certificate.Thumbprint) expires on $($certificate.NotAfter.ToString('yyyy-MM-dd')). Renew it soon."
    }

    $script:MspState.Certificate = $certificate
    $script:MspState.CertificateSource = $source
    $certificate
}
