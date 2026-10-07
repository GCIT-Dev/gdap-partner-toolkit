function Resolve-MspTokenRequestScope {
    <#
    .SYNOPSIS
        Turns -Resource or -Scope input into one resource plus the v2.0 scope string.
    .DESCRIPTION
        The v2.0 endpoint issues a token for one resource at a time. Scopes
        without a resource prefix (for example User.Read.All) are treated as
        Microsoft Graph scopes. OIDC scopes are allowed alongside. Scopes from
        more than one resource cause an error. offline_access is always added
        so the rotated refresh token comes back.
    .PARAMETER Resource
        Optional resource alias, URI or app ID. Defaults to Microsoft Graph when
        neither Resource nor Scope is given.
    .PARAMETER Scope
        Optional specific scopes.
    .EXAMPLE
        Resolve-MspTokenRequestScope -Scope 'User.Read.All', 'Group.Read.All'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Resource,
        [string[]]$Scope
    )

    $oidc = $script:MspConstants.OidcScopes
    $requested = @($Scope | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })

    if ($requested.Count -eq 0) {
        $res = Resolve-MspResource -Resource ($(if ($Resource) { $Resource } else { 'Graph' }))
        return [pscustomobject]@{
            Resource    = $res.Uri
            ResourceInfo = $res
            ScopeString = "$($res.Uri)/.default offline_access"
            ShortScopes = @()
        }
    }

    $resourceUris = [System.Collections.Generic.List[string]]::new()
    $qualified = [System.Collections.Generic.List[string]]::new()
    $short = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $requested) {
        if ($oidc -contains $item.ToLowerInvariant()) {
            continue
        }
        $index = $item.LastIndexOf('/')
        if ($index -gt 0 -and ($item -match '^(https|api)://' -or $item -match '^[0-9a-fA-F-]{36}/')) {
            $res = Resolve-MspResource -Resource $item.Substring(0, $index)
            $name = $item.Substring($index + 1)
        }
        else {
            $res = Resolve-MspResource -Resource 'Graph'
            $name = $item
        }
        if (-not $resourceUris.Contains($res.Uri)) { $resourceUris.Add($res.Uri) }
        $qualified.Add("$($res.Uri)/$name")
        if ($name -ne '.default') { $short.Add($name) }
    }

    if ($Resource) {
        $explicit = Resolve-MspResource -Resource $Resource
        if ($resourceUris.Count -gt 0 -and ($resourceUris.Count -gt 1 -or $resourceUris[0] -ne $explicit.Uri)) {
            throw (New-MspErrorRecord -Message "The scopes do not belong to resource '$($explicit.Uri)'." -ErrorId 'MspGdap.Scope.ResourceMismatch' -Category InvalidArgument -TargetObject $Scope)
        }
        if ($resourceUris.Count -eq 0) { $resourceUris.Add($explicit.Uri) }
    }
    if ($resourceUris.Count -eq 0) {
        $resourceUris.Add((Resolve-MspResource -Resource 'Graph').Uri)
        $qualified.Add("$($resourceUris[0])/.default")
    }
    if ($resourceUris.Count -gt 1) {
        throw (New-MspErrorRecord -Message "Scopes from more than one resource were requested ($($resourceUris -join ', ')). Request one resource per token." -ErrorId 'MspGdap.Scope.MultipleResources' -Category InvalidArgument -TargetObject $Scope)
    }

    $all = [System.Collections.Generic.List[string]]::new()
    foreach ($q in $qualified) { if (-not $all.Contains($q)) { $all.Add($q) } }
    foreach ($o in $requested | Where-Object { $oidc -contains $_.ToLowerInvariant() }) {
        if (-not $all.Contains($o.ToLowerInvariant())) { $all.Add($o.ToLowerInvariant()) }
    }
    if (-not $all.Contains('offline_access')) { $all.Add('offline_access') }

    [pscustomobject]@{
        Resource     = $resourceUris[0]
        ResourceInfo = Resolve-MspResource -Resource $resourceUris[0]
        ScopeString  = $all -join ' '
        ShortScopes  = @($short | Select-Object -Unique)
    }
}
