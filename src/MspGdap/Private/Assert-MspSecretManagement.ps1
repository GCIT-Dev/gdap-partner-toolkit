function Assert-MspSecretManagement {
    <#
    .SYNOPSIS
        Confirms Microsoft.PowerShell.SecretManagement is available and, when
        asked, that the named vault is registered.
    .PARAMETER VaultName
        Optional vault to check with Get-SecretVault.
    .EXAMPLE
        Assert-MspSecretManagement -VaultName 'MspGdapVault'
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [string]$VaultName
    )
    if (-not (Get-Command -Name 'Get-Secret' -ErrorAction SilentlyContinue)) {
        Import-Module -Name 'Microsoft.PowerShell.SecretManagement' -ErrorAction SilentlyContinue
    }
    foreach ($command in 'Get-Secret', 'Set-Secret', 'Get-SecretInfo', 'Get-SecretVault') {
        if (-not (Get-Command -Name $command -ErrorAction SilentlyContinue)) {
            throw (New-MspErrorRecord -Message 'Microsoft.PowerShell.SecretManagement is required to store technician refresh tokens. Install it with: Install-PSResource Microsoft.PowerShell.SecretManagement, then register a vault (for example Az.KeyVault or Microsoft.PowerShell.SecretStore).' -ErrorId 'MspGdap.SecretManagement.Missing' -Category NotInstalled)
        }
    }
    if ($VaultName) {
        try {
            $null = Get-SecretVault -Name $VaultName -ErrorAction Stop
        }
        catch {
            throw (New-MspErrorRecord -Message "Secret vault '$VaultName' is not registered. Register it with Register-SecretVault, then run Set-MspConfiguration -VaultName." -ErrorId 'MspGdap.SecretManagement.VaultNotFound' -Category ObjectNotFound -TargetObject $VaultName)
        }
    }
}
