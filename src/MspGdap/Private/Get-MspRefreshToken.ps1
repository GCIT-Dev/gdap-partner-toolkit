function Get-MspRefreshToken {
    <#
    .SYNOPSIS
        Reads a technician refresh token from the configured SecretManagement vault.
    .DESCRIPTION
        Returns the token as a SecureString. The value is never cached in
        module state. It is read for each redemption and the rotated token is
        written back by Save-MspRefreshToken.
    .PARAMETER UserPrincipalName
        Technician UPN.
    .PARAMETER AppId
        Partner app (client) ID.
    .PARAMETER VaultName
        SecretManagement vault name.
    .EXAMPLE
        $rt = Get-MspRefreshToken -UserPrincipalName $upn -AppId $appId -VaultName 'MspGdapVault'
    #>
    [CmdletBinding()]
    [OutputType([securestring])]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$VaultName
    )

    $name = Get-MspRefreshTokenSecretName -UserPrincipalName $UserPrincipalName -AppId $AppId
    try {
        $secret = Get-Secret -Name $name -Vault $VaultName -ErrorAction Stop
    }
    catch {
        throw (New-MspErrorRecord -Message "No refresh token is stored for $UserPrincipalName in vault '$VaultName' (secret $name). Run Register-MspPartnerToken first. Detail: $($_.Exception.Message)" -ErrorId 'MspGdap.RefreshToken.NotFound' -Category ObjectNotFound -TargetObject $name)
    }

    if ($secret -is [securestring]) { return $secret }
    if ($secret -is [string]) { return (ConvertTo-MspSecureString -Value $secret) }
    throw (New-MspErrorRecord -Message "Secret $name in vault '$VaultName' is not a string secret." -ErrorId 'MspGdap.RefreshToken.InvalidType' -Category InvalidType -TargetObject $name)
}
