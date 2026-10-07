function New-MspClientAssertion {
    <#
    .SYNOPSIS
        Builds a signed client assertion JWT for certificate credentials.
    .DESCRIPTION
        Creates the client_assertion used with client_assertion_type
        urn:ietf:params:oauth:client-assertion-type:jwt-bearer when the partner
        app authenticates with a certificate.

        Header: alg (PS256 by default, RS256 as fallback), typ JWT and x5t#S256
        (base64url SHA-256 thumbprint of the DER certificate). When RS256 is used
        the legacy x5t (SHA-1) header is also included for compatibility.

        Claims: aud is the v2.0 token endpoint of the tenant the assertion is
        sent to (the customer tenant for customer tokens), iss and sub are the
        app ID, jti is unique per assertion, nbf and iat are now, exp is nbf
        plus 5 minutes by default.

        Microsoft Learn specifies PS256 with PSS padding. Legacy CryptoAPI (CSP)
        keys cannot sign with PSS, and Exchange app-only authentication needs a
        CSP key, so when PS256 signing fails with a cryptographic error the
        function falls back to RS256 and says so in verbose output.
    .PARAMETER Certificate
        X509Certificate2 with an accessible RSA private key.
    .PARAMETER ClientId
        Application (client) ID of the partner app.
    .PARAMETER TenantId
        Tenant whose token endpoint receives the assertion.
    .PARAMETER Algorithm
        PS256 (default) or RS256.
    .PARAMETER LifetimeSeconds
        Assertion lifetime. Learn recommends 5 to 10 minutes at most.
    .EXAMPLE
        New-MspClientAssertion -Certificate $cert -ClientId $appId -TenantId $customerTenantId
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Builds an in-memory string and changes no state.')]
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$ClientId,

        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$|^organizations$')]
        [string]$TenantId,

        [ValidateSet('PS256', 'RS256')]
        [string]$Algorithm = 'PS256',

        [ValidateRange(60, 600)]
        [int]$LifetimeSeconds = 300
    )

    if (-not $Certificate.HasPrivateKey) {
        throw (New-MspErrorRecord -Message "Certificate $($Certificate.Thumbprint) has no accessible private key." -ErrorId 'MspGdap.Certificate.NoPrivateKey' -Category InvalidArgument -TargetObject $Certificate.Thumbprint)
    }
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if ($null -eq $rsa) {
        throw (New-MspErrorRecord -Message "Certificate $($Certificate.Thumbprint) does not have an RSA private key. Entra ID client assertions need RSA." -ErrorId 'MspGdap.Certificate.NotRsa' -Category InvalidArgument -TargetObject $Certificate.Thumbprint)
    }

    try {
        $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $claims = [ordered]@{
            aud = '{0}/{1}/oauth2/v2.0/token' -f $script:MspConstants.Authority, $TenantId.ToLowerInvariant()
            exp = $now + $LifetimeSeconds
            iss = $ClientId.ToLowerInvariant()
            jti = [guid]::NewGuid().ToString()
            nbf = $now
            iat = $now
            sub = $ClientId.ToLowerInvariant()
        }
        $claimsSegment = ConvertTo-MspBase64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes(($claims | ConvertTo-Json -Compress)))
        $x5tS256 = ConvertTo-MspBase64Url -Bytes ([System.Security.Cryptography.SHA256]::HashData($Certificate.RawData))

        $attempts = if ($Algorithm -eq 'PS256') { @('PS256', 'RS256') } else { @('RS256') }
        foreach ($alg in $attempts) {
            $header = [ordered]@{ alg = $alg; typ = 'JWT'; 'x5t#S256' = $x5tS256 }
            if ($alg -eq 'RS256') {
                $header['x5t'] = ConvertTo-MspBase64Url -Bytes $Certificate.GetCertHash()
            }
            $headerSegment = ConvertTo-MspBase64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json -Compress)))
            $signingInput = [System.Text.Encoding]::ASCII.GetBytes("$headerSegment.$claimsSegment")
            try {
                $signature = Invoke-MspRsaSignature -Rsa $rsa -Data $signingInput -Algorithm $alg
            }
            catch {
                $isCryptoError = $false
                $detail = $null
                $current = $_.Exception
                while ($current) {
                    if ($current -is [System.Security.Cryptography.CryptographicException] -or $current -is [System.NotSupportedException]) {
                        $isCryptoError = $true
                        $detail = $current.Message
                        break
                    }
                    $current = $current.InnerException
                }
                if ($alg -eq 'PS256' -and $isCryptoError) {
                    Write-Verbose "The certificate key provider cannot sign with PSS padding ($detail). Falling back to RS256."
                    continue
                }
                throw
            }
            return '{0}.{1}.{2}' -f $headerSegment, $claimsSegment, (ConvertTo-MspBase64Url -Bytes $signature)
        }
    }
    finally {
        $rsa.Dispose()
    }
}
