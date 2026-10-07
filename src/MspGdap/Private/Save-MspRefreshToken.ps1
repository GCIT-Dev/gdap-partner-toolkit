function Save-MspRefreshToken {
    <#
    .SYNOPSIS
        Writes a technician refresh token to the configured SecretManagement vault.
    .DESCRIPTION
        Stores the token as a SecureString under the name from
        Get-MspRefreshTokenSecretName. Non-sensitive metadata (upn, appId,
        created, lastUsed) is attached when the vault supports metadata.
        Metadata is not stored securely by SecretManagement, so it never holds
        anything secret. If the vault rejects metadata the secret is written
        without it.

        Called on registration and after every redemption, because the
        Microsoft identity platform returns a new refresh token each time and
        the newest one should always be the stored one.
    .PARAMETER UserPrincipalName
        Technician UPN.
    .PARAMETER AppId
        Partner app (client) ID.
    .PARAMETER VaultName
        SecretManagement vault name.
    .PARAMETER RefreshToken
        The refresh token as a SecureString.
    .EXAMPLE
        Save-MspRefreshToken -UserPrincipalName $upn -AppId $appId -VaultName 'MspGdapVault' -RefreshToken $secure
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$VaultName,

        [Parameter(Mandatory)]
        [securestring]$RefreshToken
    )

    $name = Get-MspRefreshTokenSecretName -UserPrincipalName $UserPrincipalName -AppId $AppId
    $now = [datetime]::UtcNow
    $created = $now
    try {
        $info = Get-SecretInfo -Name $name -Vault $VaultName -ErrorAction Stop | Select-Object -First 1
        if ($info -and $info.Metadata -and $info.Metadata.ContainsKey('created')) {
            $created = [datetime]$info.Metadata['created']
        }
    }
    catch {
        Write-Verbose "No existing metadata for secret $name in vault $VaultName."
    }

    $metadata = @{
        upn      = $UserPrincipalName.ToLowerInvariant()
        appId    = $AppId.ToLowerInvariant()
        created  = $created
        lastUsed = $now
    }

    try {
        Set-Secret -Name $name -SecureStringSecret $RefreshToken -Vault $VaultName -Metadata $metadata -ErrorAction Stop
    }
    catch {
        if ($_.Exception.Message -match 'metadata') {
            Write-Verbose "Vault $VaultName does not support secret metadata. Storing the secret without metadata."
            Set-Secret -Name $name -SecureStringSecret $RefreshToken -Vault $VaultName -ErrorAction Stop
        }
        else {
            throw (New-MspErrorRecord -Message "Could not write the refresh token for $UserPrincipalName to vault '$VaultName' (secret $name): $($_.Exception.Message)" -ErrorId 'MspGdap.RefreshToken.WriteFailed' -Category WriteError -TargetObject $name)
        }
    }
    Write-Verbose "Refresh token for $UserPrincipalName stored as $name in vault $VaultName."
    $name
}
