function New-MspRandomString {
    <#
    .SYNOPSIS
        Returns a cryptographically random base64url string (for OAuth state and nonce).
    .PARAMETER ByteCount
        Number of random bytes before encoding.
    .EXAMPLE
        $state = New-MspRandomString
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Creates an in-memory random value only.')]
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [ValidateRange(16, 128)]
        [int]$ByteCount = 32
    )
    ConvertTo-MspBase64Url -Bytes ([System.Security.Cryptography.RandomNumberGenerator]::GetBytes($ByteCount))
}
