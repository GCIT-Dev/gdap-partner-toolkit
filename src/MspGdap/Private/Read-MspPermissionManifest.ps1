function Read-MspPermissionManifest {
    <#
    .SYNOPSIS
        Loads and validates a permission manifest (manifests/*.json).
    .DESCRIPTION
        Returns the manifest with two views of requiredResourceAccess:
        - RequiredResourceAccess: as written, including the documentation fields resourceDisplayName and value
        - GraphRequiredResourceAccess: only resourceAppId and resourceAccess[{id, type}], safe to send to Microsoft Graph
        Throws on malformed IDs, unknown types or duplicate entries.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Path)

    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Manifest not found: $Path" }
    try { $manifest = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "Manifest $Path is not valid JSON: $($_.Exception.Message)" }

    if (-not $manifest.PSObject.Properties['requiredResourceAccess'] -or @($manifest.requiredResourceAccess).Count -eq 0) {
        throw "Manifest $Path has no requiredResourceAccess entries."
    }

    $seenResources = @{}
    $graphView = New-Object System.Collections.Generic.List[object]
    $readable = New-Object System.Collections.Generic.List[object]
    foreach ($resource in @($manifest.requiredResourceAccess)) {
        $resourceAppId = [string]$resource.resourceAppId
        if ($resourceAppId -notmatch $guid) { throw "Manifest $Path has an invalid resourceAppId '$resourceAppId'." }
        if ($seenResources.ContainsKey($resourceAppId.ToLowerInvariant())) { throw "Manifest $Path lists resource $resourceAppId more than once." }
        $seenResources[$resourceAppId.ToLowerInvariant()] = $true

        $seenIds = @{}
        $access = New-Object System.Collections.Generic.List[object]
        $readableAccess = New-Object System.Collections.Generic.List[object]
        foreach ($entry in @($resource.resourceAccess)) {
            $id = [string]$entry.id
            $type = [string]$entry.type
            if ($id -notmatch $guid) { throw "Manifest $Path has an invalid permission id '$id' under $resourceAppId." }
            if ($type -cnotin @('Scope', 'Role')) { throw "Manifest $Path has type '$type' for $id. Use Scope or Role." }
            $key = ('{0}|{1}' -f $id, $type).ToLowerInvariant()
            if ($seenIds.ContainsKey($key)) { throw "Manifest $Path lists permission $id ($type) twice under $resourceAppId." }
            $seenIds[$key] = $true
            $access.Add([pscustomobject]@{ id = $id; type = $type })
            $value = if ($entry.PSObject.Properties['value']) { [string]$entry.value } else { $null }
            $readableAccess.Add([pscustomobject]@{ id = $id; type = $type; value = $value })
        }
        if ($access.Count -eq 0) { throw "Manifest $Path has no resourceAccess entries for $resourceAppId." }
        $displayName = if ($resource.PSObject.Properties['resourceDisplayName']) { [string]$resource.resourceDisplayName } else { $resourceAppId }
        $graphView.Add([pscustomobject]@{ resourceAppId = $resourceAppId; resourceAccess = $access.ToArray() })
        $readable.Add([pscustomobject]@{ resourceAppId = $resourceAppId; resourceDisplayName = $displayName; resourceAccess = $readableAccess.ToArray() })
    }

    [pscustomobject]@{
        Path                        = (Resolve-Path -LiteralPath $Path).ProviderPath
        Name                        = if ($manifest.PSObject.Properties['name']) { [string]$manifest.name } else { [System.IO.Path]::GetFileNameWithoutExtension($Path) }
        RequiredResourceAccess      = $readable.ToArray()
        GraphRequiredResourceAccess = $graphView.ToArray()
    }
}
