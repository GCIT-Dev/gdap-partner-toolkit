function ConvertTo-MspSecureString {
    <#
    .SYNOPSIS
        Wraps a token string in a read-only SecureString.
    .DESCRIPTION
        Used so that access tokens in the cache and refresh tokens passed to
        SecretManagement are not plain strings. On Linux and macOS a
        SecureString is not encrypted in memory. It only keeps the value out of
        default formatting and logs.
    .PARAMETER Value
        The plain text value.
    .EXAMPLE
        $secure = ConvertTo-MspSecureString -Value $response.access_token
    #>
    [CmdletBinding()]
    [OutputType([securestring])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value
    )
    $secure = [System.Net.NetworkCredential]::new('', $Value).SecurePassword
    $secure.MakeReadOnly()
    $secure
}
