function ConvertFrom-MspSecureString {
    <#
    .SYNOPSIS
        Returns the plain text of a SecureString for the moment it is needed in
        an HTTP request body or header.
    .PARAMETER SecureString
        The SecureString to read.
    .EXAMPLE
        $plain = ConvertFrom-MspSecureString -SecureString $refreshToken
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [securestring]$SecureString
    )
    [System.Net.NetworkCredential]::new('', $SecureString).Password
}
