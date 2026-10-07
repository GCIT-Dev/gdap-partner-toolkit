function ConvertTo-MspPlainToken {
    <#
    .SYNOPSIS
        Returns the bearer string from whatever Get-MspAccessToken returned.
    .DESCRIPTION
        Accepts a string, a SecureString, or an object with an AccessToken or Token property (string or SecureString).
        Only used immediately before handing a token to a Microsoft module that requires a string
        (Connect-ExchangeOnline -AccessToken, Connect-IPPSSession -AccessToken, Connect-MicrosoftTeams -AccessTokens).
        Callers must not store or log the result.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowNull()][object]$InputObject)

    if ($null -eq $InputObject) { throw 'No access token was returned.' }
    if ($InputObject -is [string]) {
        if ([string]::IsNullOrWhiteSpace($InputObject)) { throw 'The access token is empty.' }
        return $InputObject
    }
    if ($InputObject -is [System.Security.SecureString]) {
        return [System.Net.NetworkCredential]::new('', $InputObject).Password
    }
    foreach ($name in @('AccessToken', 'Token', 'access_token')) {
        $prop = $InputObject.PSObject.Properties[$name]
        if ($prop -and $null -ne $prop.Value) { return (ConvertTo-MspPlainToken -InputObject $prop.Value) }
    }
    throw 'The access token object has no AccessToken or Token property.'
}
