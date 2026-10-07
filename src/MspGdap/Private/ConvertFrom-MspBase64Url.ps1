function ConvertFrom-MspBase64Url {
    <#
    .SYNOPSIS
        Decodes a base64url string (with or without padding) to bytes.
    .PARAMETER Value
        The base64url text.
    .EXAMPLE
        ConvertFrom-MspBase64Url -Value 'eyJhbGciOiJQUzI1NiJ9'
    #>
    [CmdletBinding()]
    [OutputType([byte[]], [System.Array])]
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )
    $text = $Value.Replace('-', '+').Replace('_', '/')
    switch ($text.Length % 4) {
        0 { break }
        2 { $text += '=='; break }
        3 { $text += '='; break }
        default { throw [System.FormatException]::new('Invalid base64url length.') }
    }
    , [Convert]::FromBase64String($text)
}
