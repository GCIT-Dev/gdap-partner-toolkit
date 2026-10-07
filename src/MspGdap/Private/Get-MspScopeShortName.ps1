function Get-MspScopeShortName {
    <#
    .SYNOPSIS
        Strips the resource prefix from a scope (https://graph.microsoft.com/User.Read becomes User.Read).
    .PARAMETER Scope
        A scope as returned by the token endpoint or supplied by the caller.
    .EXAMPLE
        Get-MspScopeShortName -Scope 'https://graph.microsoft.com/User.Read'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$Scope
    )
    $index = $Scope.LastIndexOf('/')
    if ($index -ge 0 -and $index -lt $Scope.Length - 1) { $Scope.Substring($index + 1) } else { $Scope }
}
