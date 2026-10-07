function Resolve-MspDataPath {
    <#
    .SYNOPSIS
        Finds a data file shipped with the module (Data folder) or a manifest (manifests folder).
    .DESCRIPTION
        Looks in the module's own Data and manifests folders first, then in the repository's manifests
        folder (two levels above the module), so the same code works from a clone and from an installed module.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$FileName)

    $moduleRoot = Split-Path -Parent $PSScriptRoot
    $candidates = @(
        (Join-Path $moduleRoot (Join-Path 'Data' $FileName)),
        (Join-Path $moduleRoot (Join-Path 'manifests' $FileName)),
        (Join-Path (Split-Path -Parent (Split-Path -Parent $moduleRoot)) (Join-Path 'manifests' $FileName))
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).ProviderPath }
    }
    throw ("Could not find {0}. Looked in: {1}" -f $FileName, ($candidates -join '; '))
}
