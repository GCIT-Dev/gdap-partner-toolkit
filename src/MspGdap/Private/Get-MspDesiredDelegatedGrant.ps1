function Get-MspDesiredDelegatedGrant {
    <#
    .SYNOPSIS
        Works out which delegated scopes the partner app should hold in a customer, per resource.
    .DESCRIPTION
        Source is either a permission manifest (-ManifestPath) or the partner app's own requiredResourceAccess
        (-AppId alone). Either way each Scope id is resolved to its name through the resource service
        principal's oauth2PermissionScopes in the partner tenant, so the scope names sent to Partner Center
        never come from free text:
          - manifest mode throws when an id is not a published delegated scope of the resource, or when the
            manifest's value does not match the published name
          - manifest mode with -AppId also throws when a scope is not in the partner app registration's
            requiredResourceAccess, unless -AllowUnregisteredScope is given (then it is a warning)
        Only "Scope" (delegated) entries are returned. Application permissions cannot be pre-consented through
        Partner Center under GDAP, so "Role" entries are reported in SkippedRoles.
        Partner Center and Partner Customer Delegated Administration are always excluded: they are used in the
        partner tenant only.
    #>
    [CmdletBinding(DefaultParameterSetName = 'App')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'App')][Parameter(ParameterSetName = 'Manifest')][string]$AppId,
        [Parameter(Mandatory, ParameterSetName = 'Manifest')][string]$ManifestPath,
        [Parameter(ParameterSetName = 'Manifest')][switch]$AllowUnregisteredScope,
        [string[]]$ExcludeResource
    )
    # Parameters: -AppId (partner app ID), -ManifestPath (permission manifest), -AllowUnregisteredScope
    # (manifest scopes missing from the app registration only warn), -ExcludeResource (app IDs or names).
    $alwaysExcluded = @(
        'fa3d9a0c-3fb0-42cc-9193-47c7ecd2edbd', # Microsoft Partner Center
        '2832473f-ec63-45fb-976f-5d45a7d4bb91'  # Partner Customer Delegated Administration
    )
    $warnings = New-Object System.Collections.Generic.List[string]
    $skippedRoles = New-Object System.Collections.Generic.List[string]
    $grants = New-Object System.Collections.Generic.List[object]
    $appDisplayName = $null

    if ($PSCmdlet.ParameterSetName -eq 'Manifest') {
        $manifest = Read-MspPermissionManifest -Path $ManifestPath
        $registered = $null
        if ($AppId) {
            $app = @(Invoke-MspGraphCall -PartnerTenant -Path ("applications?`$filter=appId eq '{0}'&`$select=id,appId,displayName,requiredResourceAccess" -f $AppId)) | Select-Object -First 1
            if (-not $app) { throw "Application $AppId was not found in the partner tenant, so the manifest cannot be checked against it." }
            $appDisplayName = $app.displayName
            $registered = @{}
            foreach ($r in @($app.requiredResourceAccess)) {
                foreach ($a in @($r.resourceAccess)) { $registered[('{0}|{1}|{2}' -f $r.resourceAppId, $a.id, $a.type).ToLowerInvariant()] = $true }
            }
        }
        $problems = New-Object System.Collections.Generic.List[string]
        $resources = foreach ($resource in $manifest.RequiredResourceAccess) {
            if ($resource.resourceAppId -in $alwaysExcluded) { continue }
            $scopeEntries = @($resource.resourceAccess | Where-Object { $_.type -eq 'Scope' })
            foreach ($role in @($resource.resourceAccess | Where-Object { $_.type -eq 'Role' })) { $skippedRoles.Add(('{0}:{1}' -f $resource.resourceDisplayName, $role.value)) }
            if ($scopeEntries.Count -eq 0) { continue }
            $resourceSp = Get-MspServicePrincipalByAppId -PartnerTenant -AppId $resource.resourceAppId -Select 'id,appId,displayName,oauth2PermissionScopes'
            if (-not $resourceSp) { $warnings.Add("Resource $($resource.resourceDisplayName) ($($resource.resourceAppId)) has no service principal in the partner tenant, so its scopes cannot be checked. Skipped."); continue }
            $scopeNames = New-Object System.Collections.Generic.List[string]
            foreach ($entry in $scopeEntries) {
                $published = @($resourceSp.oauth2PermissionScopes | Where-Object { $_.id -eq $entry.id }) | Select-Object -First 1
                if (-not $published) { $problems.Add("$($entry.id) ($($entry.value)) is not a delegated scope of $($resourceSp.displayName)"); continue }
                if ($entry.value -and ([string]$published.value -cne [string]$entry.value)) { $problems.Add("$($entry.id) is '$($published.value)' in $($resourceSp.displayName), but the manifest says '$($entry.value)'"); continue }
                if ($null -ne $registered -and -not $registered.ContainsKey(('{0}|{1}|Scope' -f $resource.resourceAppId, $entry.id).ToLowerInvariant())) {
                    $text = "$($resourceSp.displayName) $($published.value) is in the manifest but not in the partner app registration"
                    if ($AllowUnregisteredScope) { $warnings.Add("$text. Consenting it anyway because -Force was used.") }
                    else { $problems.Add("$text (add it with New-MspPartnerApp -UseExisting, or use -Force)"); continue }
                }
                $scopeNames.Add([string]$published.value)
            }
            [pscustomobject]@{ ResourceAppId = $resource.resourceAppId; ResourceDisplayName = [string]$resourceSp.displayName; Scopes = $scopeNames.ToArray() }
        }
        if ($problems.Count -gt 0) {
            throw ("Manifest {0} does not match Microsoft's published permissions or your app registration. Nothing was consented. {1}" -f $manifest.Name, ($problems -join '; '))
        }
    }
    else {
        $app = @(Invoke-MspGraphCall -PartnerTenant -Path ("applications?`$filter=appId eq '{0}'&`$select=id,appId,displayName,requiredResourceAccess" -f $AppId)) | Select-Object -First 1
        if (-not $app) { throw "Application $AppId was not found in the partner tenant. Check -AppId." }
        $appDisplayName = $app.displayName
        $resources = foreach ($resource in @($app.requiredResourceAccess)) {
            $resourceAppId = [string]$resource.resourceAppId
            if ($resourceAppId -in $alwaysExcluded) { continue }
            $scopeEntries = @($resource.resourceAccess | Where-Object { $_.type -eq 'Scope' })
            foreach ($role in @($resource.resourceAccess | Where-Object { $_.type -eq 'Role' })) { $skippedRoles.Add(('{0}:{1}' -f $resourceAppId, $role.id)) }
            if ($scopeEntries.Count -eq 0) { continue }
            $resourceSp = Get-MspServicePrincipalByAppId -PartnerTenant -AppId $resourceAppId -Select 'id,appId,displayName,oauth2PermissionScopes'
            if (-not $resourceSp) { $warnings.Add("Resource $resourceAppId has no service principal in the partner tenant, so its scope IDs cannot be named. Skipped."); continue }
            $scopeNames = New-Object System.Collections.Generic.List[string]
            foreach ($entry in $scopeEntries) {
                $match = @($resourceSp.oauth2PermissionScopes | Where-Object { $_.id -eq $entry.id }) | Select-Object -First 1
                if ($match) { $scopeNames.Add([string]$match.value) }
                else { $warnings.Add("Scope id $($entry.id) is not published by $($resourceSp.displayName). Skipped.") }
            }
            [pscustomobject]@{ ResourceAppId = $resourceAppId; ResourceDisplayName = [string]$resourceSp.displayName; Scopes = $scopeNames.ToArray() }
        }
    }

    foreach ($resource in @($resources)) {
        if ($resource.ResourceAppId -in $alwaysExcluded) { continue }
        $excluded = $false
        foreach ($exclusion in @($ExcludeResource)) {
            if ([string]::IsNullOrWhiteSpace($exclusion)) { continue }
            if ($resource.ResourceAppId -ieq $exclusion -or $resource.ResourceDisplayName -ieq $exclusion -or $resource.ResourceDisplayName -like $exclusion) { $excluded = $true }
        }
        if ($excluded) { continue }
        $unique = @($resource.Scopes | Sort-Object -Unique)
        if ($unique.Count -eq 0) { continue }
        $grants.Add([pscustomobject]@{ ResourceAppId = $resource.ResourceAppId; ResourceDisplayName = $resource.ResourceDisplayName; Scopes = $unique })
    }

    [pscustomobject]@{
        AppDisplayName = $appDisplayName
        Grants         = $grants.ToArray()
        SkippedRoles   = $skippedRoles.ToArray()
        Warnings       = $warnings.ToArray()
    }
}
