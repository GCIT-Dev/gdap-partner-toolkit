function Resolve-MspGraphUri {
    <#
    .SYNOPSIS
        Builds an absolute Microsoft Graph URI and refuses any other host.
    .DESCRIPTION
        Relative paths get https://graph.microsoft.com/<ApiVersion>/ prepended
        unless they already start with v1.0/ or beta/. Absolute URIs (including
        @odata.nextLink values) must be https://graph.microsoft.com, so a
        bearer token can never be sent to another host.
    .PARAMETER Uri
        Relative path such as users?$select=id or an absolute Graph URI.
    .PARAMETER ApiVersion
        v1.0 or beta.
    .EXAMPLE
        Resolve-MspGraphUri -Uri 'users' -ApiVersion 'v1.0'
    #>
    [CmdletBinding()]
    [OutputType([uri])]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [ValidateSet('v1.0', 'beta')]
        [string]$ApiVersion = 'v1.0'
    )
    $base = if ($script:MspConstants -and $script:MspConstants.GraphBaseUri) { $script:MspConstants.GraphBaseUri } else { 'https://graph.microsoft.com' }
    if ($Uri -match '^[a-zA-Z][a-zA-Z0-9+.-]*://') {
        $absolute = [uri]$Uri
        if ($absolute.Scheme -ne 'https' -or $absolute.Host -ne ([uri]$base).Host -or -not $absolute.IsDefaultPort) {
            throw (New-MspErrorRecord -Message "Refusing to send a Microsoft Graph token to '$($absolute.GetLeftPart([System.UriPartial]::Authority))'. Only $base is allowed." -ErrorId 'MspGdap.Graph.ForeignHost' -Category SecurityError -TargetObject $absolute.Host)
        }
        return $absolute
    }
    $path = $Uri.TrimStart('/')
    if ($path -notmatch '^(v1\.0|beta)(/|$)') {
        $path = "$ApiVersion/$path"
    }
    [uri]"$base/$path"
}
