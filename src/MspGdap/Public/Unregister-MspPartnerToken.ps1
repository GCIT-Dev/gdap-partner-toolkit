function Unregister-MspPartnerToken {
    <#
    .SYNOPSIS
        Removes a technician's stored refresh token from the vault and clears their cached access tokens.
    .DESCRIPTION
        Use it when a technician leaves, changes role or their token may be exposed. It:
          1. removes the vault secret MspGdap-<appId>-<hash of UPN> (Remove-Secret)
          2. removes every cached access token issued to that technician in this session
          3. clears the session's selected technician if it is this one

        Removing the vault copy does not revoke tokens Microsoft already issued. Also revoke the account's
        sessions in the partner tenant (Microsoft Entra admin center, or Graph revokeSignInSessions) and
        remove the account from your GDAP security groups.
    .PARAMETER UserPrincipalName
        The technician admin account.
    .PARAMETER VaultName
        Vault to remove the secret from. Defaults to the configured vault.
    .EXAMPLE
        Unregister-MspPartnerToken -UserPrincipalName 'jane.admin@contoso-msp.onmicrosoft.com' -WhatIf
    .OUTPUTS
        MspGdap.TechnicianUnregistration (no token values)
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType('MspGdap.TechnicianUnregistration')]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
        [string]$UserPrincipalName,

        [ValidateNotNullOrEmpty()]
        [string]$VaultName
    )
    process {
        try {
            $config = Get-MspConfigurationInternal
            if (-not $config.AppId) {
                throw (New-MspErrorRecord -Message 'AppId is not configured. Run Set-MspConfiguration -AppId first.' -ErrorId 'MspGdap.Configuration.Incomplete' -Category NotSpecified)
            }
            $vault = if ($VaultName) { $VaultName } else { [string]$config.VaultName }
            if (-not $vault) {
                throw (New-MspErrorRecord -Message 'No vault given or configured. Pass -VaultName.' -ErrorId 'MspGdap.Configuration.Incomplete' -Category NotSpecified)
            }
            Assert-MspSecretManagement -VaultName $vault
            $upn = $UserPrincipalName.Trim().ToLowerInvariant()
            $appId = ([string]$config.AppId).ToLowerInvariant()
            $secretName = Get-MspRefreshTokenSecretName -UserPrincipalName $upn -AppId $appId

            $exists = @(Get-SecretInfo -Name $secretName -Vault $vault -ErrorAction Stop | Where-Object { $_.Name -eq $secretName }).Count -gt 0
            $removed = $false
            if (-not $exists) {
                Write-Warning "No stored refresh token was found for $upn in vault '$vault'."
            }
            elseif ($PSCmdlet.ShouldProcess("vault '$vault'", "Remove the refresh token of $upn ($secretName)")) {
                Remove-Secret -Name $secretName -Vault $vault -ErrorAction Stop
                $removed = $true
            }
            else {
                return
            }

            $cleared = 0
            foreach ($key in @($script:MspTokenCache.Keys)) {
                $entry = $script:MspTokenCache[$key]
                if ($entry -and $entry.PSObject.Properties['UserPrincipalName'] -and ([string]$entry.UserPrincipalName).ToLowerInvariant() -eq $upn) {
                    $script:MspTokenCache.Remove($key)
                    $cleared++
                }
            }
            if ($script:MspState.CurrentUpn -and ([string]$script:MspState.CurrentUpn).ToLowerInvariant() -eq $upn) { $script:MspState.CurrentUpn = $null }
            if ($config.TechnicianUpn -and ([string]$config.TechnicianUpn).ToLowerInvariant() -eq $upn) {
                Write-Warning "$upn is still the default technician in the configuration. Run Set-MspConfiguration -TechnicianUpn with another account."
            }
            Write-Verbose "Removed $cleared cached access token(s) for $upn."

            [pscustomobject]@{
                PSTypeName            = 'MspGdap.TechnicianUnregistration'
                UserPrincipalName     = $upn
                AppId                 = $appId
                VaultName             = $vault
                SecretName            = $secretName
                RefreshTokenRemoved   = $removed
                CachedTokensCleared   = $cleared
                NextStep              = 'Revoke the account''s sessions in the partner tenant and remove it from your GDAP security groups.'
            }
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }
}
