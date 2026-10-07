function New-MspAddKeyProof {
    <#
    .SYNOPSIS
        Creates the proof-of-possession JWT for Microsoft Graph application addKey.
    .DESCRIPTION
        Microsoft Learn (application: addKey): the proof must be signed with the private key of one of the
        application's existing valid certificates, with aud 00000002-0000-0000-c000-000000000000, iss the ID of
        the application object, nbf, and exp = nbf + 10 minutes. Signed RS256 (PKCS#1 v1.5) with an x5t header.
        The private key never leaves the certificate provider.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory token only.')]
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$ApplicationObjectId,
        [datetime]$Now = [datetime]::UtcNow
    )
    $toB64Url = {
        param([byte[]]$Bytes)
        [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
    $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw "Certificate $($Certificate.Thumbprint) has no usable RSA private key for the proof token." }

    $nbf = [int64]([DateTimeOffset]$Now.ToUniversalTime()).ToUnixTimeSeconds()
    $header = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = (& $toB64Url $Certificate.GetCertHash()) }
    $payload = [ordered]@{ aud = '00000002-0000-0000-c000-000000000000'; iss = $ApplicationObjectId; nbf = $nbf; exp = $nbf + 600 }
    $encodedHeader = & $toB64Url ([Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json -Compress)))
    $encodedPayload = & $toB64Url ([Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress)))
    $unsigned = $encodedHeader + '.' + $encodedPayload
    $signature = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($unsigned), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    return $unsigned + '.' + (& $toB64Url $signature)
}
