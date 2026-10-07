function Get-MspGdapRoleCatalog {
    <#
    .SYNOPSIS
        Loads the role catalogue (Data/gdap-roles.leastprivilege.json) and returns its roles.
    .PARAMETER Path
        Catalogue file. Defaults to the copy shipped in Data/.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$Path)

    if (-not $Path) { $Path = Resolve-MspDataPath -FileName 'gdap-roles.leastprivilege.json' }
    $catalog = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    foreach ($role in @($catalog.roles)) {
        [pscustomobject]@{
            DisplayName            = [string]$role.displayName
            RoleTemplateId         = ([string]$role.roleTemplateId).ToLowerInvariant()
            Default                = [bool]$role.default
            HighlyPrivileged       = [bool]$role.highlyPrivileged
            RequiresExplicitSwitch = [bool]($role.PSObject.Properties['requiresExplicitSwitch'] -and $role.requiresExplicitSwitch)
            Purpose                = [string]$role.purpose
            Unknown                = $false
        }
    }
}
