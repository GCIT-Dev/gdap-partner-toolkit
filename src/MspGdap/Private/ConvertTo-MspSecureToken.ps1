function ConvertTo-MspSecureToken {
    <#
    .SYNOPSIS
        Returns the access token from whatever Get-MspAccessToken returned, as a read-only SecureString.
    .DESCRIPTION
        Used for Connect-MgGraph -AccessToken, which takes a SecureString. Builds the SecureString one character
        at a time so no plain-text conversion cmdlet is needed.
    #>
    [CmdletBinding()]
    [OutputType([System.Security.SecureString])]
    param([Parameter(Mandatory)][AllowNull()][object]$InputObject)

    if ($InputObject -is [System.Security.SecureString]) { return $InputObject }
    foreach ($name in @('AccessToken', 'Token')) {
        if ($null -ne $InputObject -and $InputObject -isnot [string]) {
            $prop = $InputObject.PSObject.Properties[$name]
            if ($prop -and $prop.Value -is [System.Security.SecureString]) { return $prop.Value }
        }
    }
    $plain = ConvertTo-MspPlainToken -InputObject $InputObject
    $secure = New-Object System.Security.SecureString
    foreach ($character in $plain.ToCharArray()) { $secure.AppendChar($character) }
    $secure.MakeReadOnly()
    $plain = $null
    return $secure
}
