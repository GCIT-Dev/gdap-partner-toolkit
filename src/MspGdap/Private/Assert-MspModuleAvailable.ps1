function Assert-MspModuleAvailable {
    <#
    .SYNOPSIS
        Makes sure an optional Microsoft module is installed at a minimum version, imports it, and returns its version.
    #>
    [CmdletBinding()]
    [OutputType([version])]
    param(
        [Parameter(Mandatory)][string]$Name,
        [version]$MinimumVersion = '0.0'
    )
    $loaded = @(Get-Module -Name $Name | Sort-Object -Property Version -Descending) | Select-Object -First 1
    if ($loaded -and $loaded.Version -ge $MinimumVersion) { return $loaded.Version }

    $available = @(Get-Module -ListAvailable -Name $Name | Sort-Object -Property Version -Descending) | Select-Object -First 1
    if (-not $available) {
        throw "$Name is not installed. Install it with: Install-Module -Name $Name -Scope CurrentUser -MinimumVersion $MinimumVersion"
    }
    if ($available.Version -lt $MinimumVersion) {
        throw "$Name $($available.Version) is installed, but $MinimumVersion or later is required. Run: Update-Module -Name $Name"
    }
    Import-Module -Name $Name -RequiredVersion $available.Version -ErrorAction Stop -Verbose:$false | Out-Null
    return $available.Version
}
