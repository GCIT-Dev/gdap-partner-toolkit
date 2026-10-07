function Resolve-MspGdapRole {
    <#
    .SYNOPSIS
        Turns role names, template IDs or {displayName, roleTemplateId} objects into catalogue entries.
    .DESCRIPTION
        A template ID that is not in the catalogue is refused unless -AllowUnknown is given (for custom or
        newer built-in roles). An accepted unknown ID is named "Unknown role <id>" and marked Unknown, and
        its privilege level cannot be checked. A name that is not in the catalogue always throws, so typos
        fail before any write. When an object carries both a name and an ID that disagree with the
        catalogue, it throws.
    .PARAMETER Role
        Role names, template IDs or objects with displayName and roleTemplateId.
    .PARAMETER Catalog
        Catalogue entries from Get-MspGdapRoleCatalog. Loaded when omitted.
    .PARAMETER AllowUnknown
        Accept template IDs that are not in the catalogue.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Role,
        [object[]]$Catalog,
        [switch]$AllowUnknown
    )
    $guid = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
    if (-not $Catalog) { $Catalog = @(Get-MspGdapRoleCatalog) }
    $seen = @{}
    foreach ($item in $Role) {
        if ($null -eq $item) { continue }
        $name = $null; $id = $null
        if ($item -is [string]) {
            if ($item -match $guid) { $id = $item } else { $name = $item }
        }
        else {
            if ($item.PSObject.Properties['roleTemplateId']) { $id = [string]$item.roleTemplateId }
            elseif ($item.PSObject.Properties['RoleTemplateId']) { $id = [string]$item.RoleTemplateId }
            elseif ($item.PSObject.Properties['roleDefinitionId']) { $id = [string]$item.roleDefinitionId }
            if ($item.PSObject.Properties['displayName']) { $name = [string]$item.displayName }
        }

        $entry = $null
        if ($id) {
            if ($id -notmatch $guid) { throw "Role template ID '$id' is not a GUID." }
            $entry = @($Catalog | Where-Object { $_.RoleTemplateId -eq $id.ToLowerInvariant() }) | Select-Object -First 1
            if ($entry -and $name -and ($entry.DisplayName -ne $name)) {
                throw "Role '$name' does not match template ID $id, which is '$($entry.DisplayName)'. Fix the access map."
            }
            if (-not $entry) {
                if (-not $AllowUnknown) {
                    throw "Role template ID $id is not in the MspGdap role catalogue, so its privilege level cannot be checked. Use a catalogue role, or pass -AllowUnknownRole if you have checked this role yourself."
                }
                $entry = [pscustomobject]@{
                    DisplayName = if ($name) { $name } else { "Unknown role $id" }
                    RoleTemplateId = $id.ToLowerInvariant(); Default = $false; HighlyPrivileged = $false; RequiresExplicitSwitch = $false; Purpose = ''; Unknown = $true
                }
            }
        }
        elseif ($name) {
            $entry = @($Catalog | Where-Object { $_.DisplayName -ieq $name.Trim() }) | Select-Object -First 1
            if (-not $entry) { throw "Unknown role name '$name'. Use the exact Microsoft Entra role name or its template ID." }
        }
        else {
            throw 'A role entry has neither a name nor a template ID.'
        }

        if (-not $seen.ContainsKey($entry.RoleTemplateId)) {
            $seen[$entry.RoleTemplateId] = $true
            $entry
        }
    }
}
