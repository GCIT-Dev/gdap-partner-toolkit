function New-MspPkceChallenge {
    <#
    .SYNOPSIS
        Creates a PKCE code verifier and its S256 code challenge (RFC 7636).
    .DESCRIPTION
        The verifier is 32 random bytes encoded as base64url (43 characters).
        The challenge is base64url(SHA-256(ASCII(verifier))). Only S256 is
        produced, as Microsoft recommends.
    .EXAMPLE
        $pkce = New-MspPkceChallenge
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates in-memory random values only.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $verifier = ConvertTo-MspBase64Url -Bytes ([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
    $challenge = ConvertTo-MspBase64Url -Bytes ([System.Security.Cryptography.SHA256]::HashData([System.Text.Encoding]::ASCII.GetBytes($verifier)))
    [pscustomobject]@{
        Verifier  = $verifier
        Challenge = $challenge
        Method    = 'S256'
    }
}
