function Invoke-MspRsaSignature {
    <#
    .SYNOPSIS
        Signs bytes with SHA-256 and PSS (PS256) or PKCS#1 v1.5 (RS256) padding.
    .DESCRIPTION
        Kept separate from New-MspClientAssertion so the PS256 to RS256 fallback
        for legacy CSP keys can be tested.
    .PARAMETER Rsa
        The RSA private key.
    .PARAMETER Data
        Bytes to sign.
    .PARAMETER Algorithm
        PS256 or RS256.
    .EXAMPLE
        Invoke-MspRsaSignature -Rsa $rsa -Data $bytes -Algorithm PS256
    #>
    [CmdletBinding()]
    [OutputType([byte[]], [System.Array])]
    param(
        [Parameter(Mandatory)]
        [System.Security.Cryptography.RSA]$Rsa,

        [Parameter(Mandatory)]
        [byte[]]$Data,

        [Parameter(Mandatory)]
        [ValidateSet('PS256', 'RS256')]
        [string]$Algorithm
    )
    $padding = if ($Algorithm -eq 'PS256') { [System.Security.Cryptography.RSASignaturePadding]::Pss } else { [System.Security.Cryptography.RSASignaturePadding]::Pkcs1 }
    , $Rsa.SignData($Data, [System.Security.Cryptography.HashAlgorithmName]::SHA256, $padding)
}
