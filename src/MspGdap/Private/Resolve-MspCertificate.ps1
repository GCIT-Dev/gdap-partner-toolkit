function Resolve-MspCertificate {
    <#
    .SYNOPSIS
        Returns an X509Certificate2 from an object, a thumbprint (CurrentUser\My then LocalMachine\My) or a public
        certificate file (.cer, .crt, .pem). Validates dates and key size.
    .DESCRIPTION
        PFX files are refused: only the public key is ever uploaded, and a PFX would need a password.
        -RequirePrivateKey is used when the certificate must sign (proof of possession for addKey).
    #>
    [CmdletBinding(DefaultParameterSetName = 'Object')]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Object')][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$Thumbprint,
        [Parameter(Mandatory, ParameterSetName = 'Path')][string]$Path,
        [switch]$RequirePrivateKey,
        [datetime]$Now = [datetime]::UtcNow
    )
    $cert = $null
    switch ($PSCmdlet.ParameterSetName) {
        'Object' { $cert = $Certificate }
        'Thumbprint' {
            foreach ($location in @('CurrentUser', 'LocalMachine')) {
                $store = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', $location)
                try {
                    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
                    $found = $store.Certificates.Find([System.Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $Thumbprint, $false)
                    if ($found.Count -gt 0) { $cert = $found[0]; break }
                }
                catch { Write-Verbose "Could not open the $location\My store: $($_.Exception.Message)" }
                finally { $store.Close() }
            }
            if (-not $cert) { throw "Certificate $Thumbprint was not found in CurrentUser\My or LocalMachine\My." }
        }
        'Path' {
            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Certificate file not found: $Path" }
            $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
            if ($extension -in @('.pfx', '.p12')) { throw 'PFX files are not accepted. Export the public certificate (.cer) and pass that, or import the PFX into your certificate store and use -CertificateThumbprint.' }
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList (Resolve-Path -LiteralPath $Path).ProviderPath
        }
    }

    if ($cert.NotAfter.ToUniversalTime() -le $Now) { throw "Certificate $($cert.Thumbprint) expired on $($cert.NotAfter.ToString('yyyy-MM-dd'))." }
    if ($cert.NotBefore.ToUniversalTime() -gt $Now.AddMinutes(5)) { throw "Certificate $($cert.Thumbprint) is not valid until $($cert.NotBefore.ToString('yyyy-MM-dd'))." }
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cert)
    if (-not $rsa) { throw "Certificate $($cert.Thumbprint) does not have an RSA key. Microsoft Entra certificate credentials need RSA." }
    # Read KeySize through the base class: on Linux RSAOpenSsl overrides it with a setter only, which PowerShell
    # reports as a write-only property.
    $keySize = [int]([System.Security.Cryptography.AsymmetricAlgorithm].GetProperty('KeySize').GetValue($rsa))
    if ($keySize -lt 2048) { throw "Certificate $($cert.Thumbprint) has a $keySize-bit key. Use 2048 bits or more." }
    if ($RequirePrivateKey -and -not $cert.HasPrivateKey) { throw "Certificate $($cert.Thumbprint) has no private key on this machine." }
    return $cert
}
