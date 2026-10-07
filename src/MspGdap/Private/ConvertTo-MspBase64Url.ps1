function ConvertTo-MspBase64Url {
    <#
    .SYNOPSIS
        Encodes bytes as base64url (RFC 4648 section 5) without padding.
    .PARAMETER Bytes
        The bytes to encode.
    .EXAMPLE
        ConvertTo-MspBase64Url -Bytes ([Text.Encoding]::UTF8.GetBytes('{"alg":"PS256"}'))
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}
