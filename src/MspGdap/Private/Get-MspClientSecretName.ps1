function Get-MspClientSecretName {
    <#
    .SYNOPSIS
        Returns the vault secret name for the partner app client secret.
    .DESCRIPTION
        Only used when the partner app authenticates with a client secret,
        which is supported but discouraged. Certificates are preferred.
    .PARAMETER AppId
        Partner app (client) ID.
    .EXAMPLE
        Get-MspClientSecretName -AppId $appId
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$AppId
    )
    'MspGdap-{0}-clientsecret' -f $AppId.ToLowerInvariant()
}
