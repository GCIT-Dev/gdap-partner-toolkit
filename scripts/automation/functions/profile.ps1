# Azure Functions profile.ps1 for the MspGdap timer functions in this folder.
# Runs once when a PowerShell worker starts (cold start). Pattern B from docs/08-unattended-automation.md:
# a managed identity reads the partner app certificate from one Key Vault and the automation account's
# refresh token from another. No password, client secret or key file is stored anywhere.
#
# Host: Azure Functions v4, PowerShell 7.4 or later (7.6 recommended, 7.4 support ends 10 November 2026).
# Bundle the MspGdap, Az.Accounts, Az.KeyVault, Microsoft.PowerShell.SecretManagement and (for the Exchange
# functions) ExchangeOnlineManagement modules in the app's Modules folder with Save-PSResource. Managed
# dependencies are not supported on the Flex Consumption plan.
#
# Publish layout (the function app root is this folder):
#   host.json, profile.ps1, Modules/<modules>, scripts/<the .ps1 files from the parent folder>, <FunctionName>/
#
# App settings (no secrets, only names):
#   MSPGDAP_PARTNER_TENANT_ID      partner tenant ID
#   MSPGDAP_PARTNER_APP_ID         MspGdap partner app (client) ID
#   MSPGDAP_AUTOMATION_UPN         dedicated automation admin account whose refresh token is in the token vault
#   MSPGDAP_TOKEN_KEYVAULT_NAME    Key Vault holding the refresh token (managed identity: Key Vault Secrets Officer)
#   MSPGDAP_SUBSCRIPTION_ID        subscription of the token Key Vault
#   MSPGDAP_CERT_KEYVAULT_NAME     separate Key Vault holding the partner app certificate (Key Vault Certificate User)
#   MSPGDAP_PARTNER_CERT_NAME      certificate name in that vault
#   MSPGDAP_TENANT_IDS             optional comma separated customer tenant IDs. Empty means every customer.
#   MSPGDAP_APPLY_CHANGES          optional, 'true' lets functions that can change settings do so. Default report only.
#   MSPGDAP_EXO_APP_ID, MSPGDAP_AUTOMATION_KEYVAULT_NAME, MSPGDAP_AUTOMATION_CERT_NAME
#                                  optional, app-only Exchange for the Exchange functions (see their run.ps1)
#   MspGdapStorage__queueServiceUri / MspGdapStorage__blobServiceUri
#                                  identity-based connection for the output bindings (Storage Queue Data
#                                  Message Sender and Storage Blob Data Contributor for the managed identity)

$ErrorActionPreference = 'Stop'

Connect-AzAccount -Identity | Out-Null

$vaultParameters = @{
    AZKVaultName   = $env:MSPGDAP_TOKEN_KEYVAULT_NAME
    SubscriptionId = $env:MSPGDAP_SUBSCRIPTION_ID
}
Register-SecretVault -Name 'MspGdapAutomation' -ModuleName 'Az.KeyVault' -VaultParameters $vaultParameters -AllowClobber

Import-Module MspGdap

# In-memory partner app certificate. PS256 signing works with ephemeral keys on every operating system.
$pfxBase64 = Get-AzKeyVaultSecret -VaultName $env:MSPGDAP_CERT_KEYVAULT_NAME -Name $env:MSPGDAP_PARTNER_CERT_NAME -AsPlainText
$storageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
$partnerCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String($pfxBase64), [string]::Empty, $storageFlags)
Remove-Variable -Name pfxBase64

$configParams = @{
    PartnerTenantId = $env:MSPGDAP_PARTNER_TENANT_ID
    AppId           = $env:MSPGDAP_PARTNER_APP_ID
    VaultName       = 'MspGdapAutomation'
    TechnicianUpn   = $env:MSPGDAP_AUTOMATION_UPN
}
Set-MspConfiguration @configParams
Set-MspConfiguration -Certificate $partnerCertificate
