function Read-MspGdapAccessMap {
    <#
    .SYNOPSIS
        Loads and validates a GDAP access map (security group to roles) supplied by the partner.
    .DESCRIPTION
        Format (accessMapVersion 1):
          { "accessMapVersion": 1, "assignments": [ { "groupDisplayName": "...", "groupId": "<guid>",
            "roles": [ { "displayName": "...", "roleTemplateId": "<guid>" } | "<name or guid>" ] } ] }
        Throws on placeholders, invalid or duplicate group IDs, unknown role names and empty role lists.
    .PARAMETER Path
        Access map JSON file.
    .PARAMETER AllowUnknownRole
        Accept role template IDs that are not in the role catalogue.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$AllowUnknownRole
    )

    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Access map not found: $Path" }
    try { $map = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop }
    catch { throw "Access map $Path is not valid JSON: $($_.Exception.Message)" }

    if (-not $map.PSObject.Properties['accessMapVersion'] -or [int]$map.accessMapVersion -ne 1) {
        throw "Access map $Path must have accessMapVersion 1."
    }
    $assignments = @($map.assignments)
    if ($assignments.Count -eq 0) { throw "Access map $Path has no assignments." }

    $catalog = @(Get-MspGdapRoleCatalog)
    $seenGroups = @{}
    foreach ($assignment in $assignments) {
        $groupId = [string]$assignment.groupId
        $groupName = if ($assignment.PSObject.Properties['groupDisplayName']) { [string]$assignment.groupDisplayName } else { $groupId }
        if ($groupId -notmatch $guid) {
            throw "Access map entry '$groupName' has groupId '$groupId'. Replace it with the object ID of a security group in your partner tenant."
        }
        if ($seenGroups.ContainsKey($groupId.ToLowerInvariant())) { throw "Access map lists group $groupId more than once." }
        $seenGroups[$groupId.ToLowerInvariant()] = $true
        $roles = @(Resolve-MspGdapRole -Role @($assignment.roles) -Catalog $catalog -AllowUnknown:$AllowUnknownRole)
        if ($roles.Count -eq 0) { throw "Access map entry '$groupName' has no roles." }
        [pscustomobject]@{
            GroupId          = $groupId.ToLowerInvariant()
            GroupDisplayName = $groupName
            Roles            = $roles
        }
    }
}
