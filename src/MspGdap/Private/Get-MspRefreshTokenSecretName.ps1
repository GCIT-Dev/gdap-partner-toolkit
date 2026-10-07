function Get-MspRefreshTokenSecretName {
    <#
    .SYNOPSIS
        Returns the vault secret name for a technician's refresh token.
    .DESCRIPTION
        Pattern: MspGdap-<appId>-<first 16 hex characters of SHA-256(lower-case UPN)>.

        The UPN is hashed rather than embedded because Azure Key Vault names
        allow only letters, digits and dashes (1 to 127 characters) and Microsoft
        advises against personal information in Key Vault object names. The
        UPN is kept in the secret metadata instead. Names are lower case
        because some vaults treat names case-insensitively.
    .PARAMETER UserPrincipalName
        Technician UPN.
    .PARAMETER AppId
        Partner app (client) ID. Refresh tokens are bound to the client.
    .EXAMPLE
        Get-MspRefreshTokenSecretName -UserPrincipalName 'tech-admin@contoso.onmicrosoft.com' -AppId $appId
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
        [string]$AppId
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($UserPrincipalName.Trim().ToLowerInvariant())
    $hash = [System.Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($bytes)).Substring(0, 16).ToLowerInvariant()
    'MspGdap-{0}-{1}' -f $AppId.ToLowerInvariant(), $hash
}
