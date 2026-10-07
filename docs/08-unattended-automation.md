# 08 Unattended automation

Everything in the earlier guides assumes a technician at the keyboard. Scheduled jobs (nightly reports, compliance checks, licence reconciliation) have no person to complete MFA. This guide covers how to run them on Azure without storing passwords, and when to choose app-only access over a delegated refresh token.

## Choose the access model first

| | App-only (preferred) | Delegated refresh token (judgement call) |
| --- | --- | --- |
| Identity in customer audit logs | The automation app | A named admin account |
| Limited by GDAP | No. Limited by the app's own permissions and roles in each customer | Yes |
| Set-up per customer | `Enable-MspExchangeAppAccess` or per-customer app role grants | Pre-consent only |
| Works for Partner Center GDAP operations (consent, relationships) | No. Microsoft only supports app-only Partner Center tokens for operations that don't need GDAP roles | Yes |
| Breaks when | Certificate expires or a customer removes the app | Token idle for 90 days, account disabled, sessions revoked, Conditional Access sign-in frequency |
| Main risk | Standing tenant-wide access in every customer where it is granted | A long-lived user credential that carries every GDAP role of that account |

**Use app-only wherever Microsoft supports it for the job**: Exchange Online (`Exchange.ManageAsApp`), Microsoft Graph application permissions for read-only reporting, the Office 365 Management API. Workload identities such as service principals and managed identities are also outside Microsoft's mandatory Azure MFA enforcement, so they don't break when MFA rules tighten.

**Use a delegated refresh token only when there is no app-only option**, for example Partner Center consent automation or an API with no application permissions. Treat it as a privileged credential and follow the controls in [the delegated section](#pattern-b-delegated-refresh-token-in-automation).

## Where to run it

| Host | Notes |
| --- | --- |
| **Azure Functions v4, PowerShell 7.6** | Generally available on every hosting plan except Linux Consumption, including Flex Consumption and Premium. Use a timer trigger so there is no public endpoint. Bundle modules in the app's `Modules` folder (`Save-PSResource`), because managed dependencies are not supported on Flex Consumption. PowerShell 7.4 is also available but leaves support on 10 November 2026. |
| **Azure Automation, PowerShell 7.6 or 7.4 runtime** | Both runtimes are supported for cloud and hybrid jobs. Import `ExchangeOnlineManagement` together with `PowerShellGet` and `PackageManagement` (Microsoft lists failures without them). |
| **Your own server** | Fine if it is managed like a tier 0 asset. Use a non-exportable certificate in the machine store and a gMSA or managed identity (Azure Arc) to reach Key Vault. |

In every case:

- Give the host a **managed identity** and use it to reach Key Vault. No client secrets or connection strings in app settings.
- Use a **dedicated Key Vault** for automation (Microsoft recommends a vault per application per environment), with the Azure RBAC permission model, soft delete, purge protection, network restrictions and `AuditEvent` logs sent to Log Analytics.
- Never expose a function key or API key that can read or write secrets on behalf of callers. If you need an HTTP trigger, protect it with Microsoft Entra authentication, not function keys, and keep it away from the vault logic.

## Pattern A: app-only with a certificate from Key Vault

1. Create the automation app and its CSP certificate ([02](02-create-partner-app.md#exchange-app-only-is-different), [05](05-exchange-access.md#app-only-exchange-for-unattended-jobs)). Import the certificate (with its private key) into the automation Key Vault as a certificate object.
2. Grant the host's managed identity **Key Vault Certificate User** on that vault ("Read entire certificate contents including secret and key portion").
3. Run `Enable-MspExchangeAppAccess` for each customer the job covers (as a technician, interactively), then `Test-MspExchangeAppAccess`.
4. Keep the job's customer list (tenant ID and initial `.onmicrosoft.com` domain) in configuration the job can read.

Timer-triggered function (`run.ps1`), PowerShell 7.6, any operating system:

```powershell
param($Timer)

$ErrorActionPreference = 'Stop'
Connect-AzAccount -Identity | Out-Null

# Load the automation app certificate.
# Linux: EphemeralKeySet keeps the private key in memory only.
# Windows: EphemeralKeySet always loads the key through CNG, which Exchange app-only does not support,
# so load it with the default flags to keep the PFX's CSP provider. Windows then holds the key in a
# temporary key container until the certificate is disposed (see the finally block at the end).
$pfxBase64 = Get-AzKeyVaultSecret -VaultName $env:MSPGDAP_KEYVAULT_NAME -Name $env:MSPGDAP_AUTOMATION_CERT_NAME -AsPlainText
$storageFlags = if ($IsWindows) {
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet
}
else {
    [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
}
$certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String($pfxBase64), [string]::Empty, $storageFlags)
Remove-Variable -Name pfxBase64

$customers = Get-Content -Path (Join-Path -Path $PSScriptRoot -ChildPath 'customers.json') -Raw | ConvertFrom-Json

foreach ($customer in $customers) {
    try {
        Connect-ExchangeOnline -AppId $env:MSPGDAP_AUTOMATION_APP_ID -Certificate $certificate -Organization $customer.InitialDomain -ShowBanner:$false
        Get-Mailbox -ResultSize Unlimited -Filter 'ForwardingSmtpAddress -ne $null' |
            Select-Object @{ Name = 'Customer'; Expression = { $customer.InitialDomain } }, UserPrincipalName, ForwardingSmtpAddress
    }
    catch {
        Write-Error -Message "Customer $($customer.InitialDomain): $($_.Exception.Message)" -ErrorAction Continue
    }
    finally {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
}
$certificate.Dispose()
```

For Microsoft Graph application permissions, use the same certificate with the Graph SDK:

```powershell
Connect-MgGraph -ClientId $env:MSPGDAP_AUTOMATION_APP_ID -TenantId $customer.TenantId -Certificate $certificate -NoWelcome
```

Notes:

- No certificate file is written by the script. On Linux `EphemeralKeySet` keeps the private key in memory only. On Windows the key sits in a temporary key container until `Dispose()`, which is the price of keeping a CSP key for Exchange.
- Exchange app-only doesn't support CNG keys, and on Windows `EphemeralKeySet` always produces a CNG key. That is why the Windows branch loads the PFX without it. The certificate must also have been created with a CSP provider in the first place ([02](02-create-partner-app.md#exchange-app-only-is-different)). Test the connection once on your host type before you rely on it. Linux hosts don't use CNG.
- The job identity in each customer's audit log is the automation app. Name it clearly (for example "Contoso MSP Automation") so customers recognise it.

## Pattern B: delegated refresh token in automation

Use this only when no app-only route exists. You are giving a scheduled job a user credential that works in every customer where that account has GDAP roles, so the controls matter more than the code.

Controls:

1. **A dedicated automation admin account** in the partner tenant, not a technician's personal admin account. Its audit trail is then clearly automation, and it doesn't break when a technician leaves.
2. **Narrow GDAP roles** for that account's groups. Read-only roles (Global Reader, Reports Reader, Security Reader) unless the job genuinely writes.
3. **Register its token interactively** with MFA (`Register-MspPartnerToken`) from a privileged access workstation, directly into the automation vault.
4. **Keep the certificate and the token apart.** Either import the partner app certificate from Key Vault into the function app as a private key certificate and load it with the `WEBSITE_LOAD_CERTIFICATES` app setting (Microsoft documents this App Service feature for Basic tier or higher, so it needs a Dedicated App Service plan), or, on any plan including Consumption and Flex Consumption, read it from a separate certificate vault at start-up and give it to MspGdap for the session only with `Set-MspConfiguration -Certificate` (shown below). On an Azure Automation Hybrid Runbook Worker, put the certificate in the machine store instead and use `-CertificateStoreLocation LocalMachine`. Keep the automation account's refresh token in its own vault, where the managed identity has **Key Vault Secrets Officer** (MspGdap writes back the rotated token after each use). Nobody else gets data roles on that vault.
5. **Conditional Access** for the account: allow sign-in only from the host's outbound IP addresses (a NAT gateway gives Azure Functions a fixed address), and alert on any sign-in from elsewhere.
6. **Run often enough** to keep the token alive (at least weekly), and alert if the job reports `AADSTS700082` (expired through inactivity) or a revoked grant.
7. **Review it quarterly**, the same as any privileged account.

`profile.ps1` (runs at cold start):

```powershell
Connect-AzAccount -Identity | Out-Null

$vaultParameters = @{
    AZKVaultName   = $env:MSPGDAP_TOKEN_KEYVAULT_NAME
    SubscriptionId = $env:MSPGDAP_SUBSCRIPTION_ID
}
Register-SecretVault -Name 'MspGdapAutomation' -ModuleName 'Az.KeyVault' -VaultParameters $vaultParameters -AllowClobber

Import-Module MspGdap
$configParams = @{
    PartnerTenantId = $env:MSPGDAP_PARTNER_TENANT_ID
    AppId           = $env:MSPGDAP_PARTNER_APP_ID
    VaultName       = 'MspGdapAutomation'
    TechnicianUpn   = $env:MSPGDAP_AUTOMATION_UPN
}

# WEBSITE_LOAD_CERTIFICATES makes the certificate available to the app:
# Windows plans load it into CurrentUser\My, Linux plans write it to /var/ssl/private/<thumbprint>.p12
if ($IsWindows) {
    $configParams.CertificateThumbprint    = $env:MSPGDAP_PARTNER_CERT_THUMBPRINT
    $configParams.CertificateStoreLocation = 'CurrentUser'
}
else {
    $configParams.CertificatePath = "/var/ssl/private/$($env:MSPGDAP_PARTNER_CERT_THUMBPRINT).p12"
}
Set-MspConfiguration @configParams
```

On a plan without `WEBSITE_LOAD_CERTIFICATES` (Consumption or Flex Consumption), replace the certificate part with an in-memory certificate. The managed identity needs **Key Vault Certificate User** on the certificate vault, which should not be the token vault:

```powershell
$pfxBase64 = Get-AzKeyVaultSecret -VaultName $env:MSPGDAP_CERT_KEYVAULT_NAME -Name $env:MSPGDAP_PARTNER_CERT_NAME -AsPlainText
$storageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
$partnerCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new([System.Convert]::FromBase64String($pfxBase64), [string]::Empty, $storageFlags)
Remove-Variable -Name pfxBase64
Set-MspConfiguration @configParams
Set-MspConfiguration -Certificate $partnerCertificate
```

`-Certificate` is kept for the session only. It is never written to the configuration file, and it takes precedence over a thumbprint or path. The partner app signs with PS256, which CNG and ephemeral keys support, so `EphemeralKeySet` is fine here on every operating system.

`Set-MspConfiguration` writes only non-secret settings, to `$HOME/.mspgdap/config.json` (or the path in the `MSPGDAP_CONFIG_PATH` app setting).

`run.ps1`:

```powershell
param($Timer)

$ErrorActionPreference = 'Stop'
$customers = Get-MspCustomer -IncludeGdapStatus

foreach ($customer in $customers) {
    try {
        Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri 'v1.0/subscribedSkus?$select=skuPartNumber,prepaidUnits,consumedUnits' |
            Select-Object @{ Name = 'Customer'; Expression = { $customer.DisplayName } }, skuPartNumber, consumedUnits
    }
    catch {
        Write-Error -Message "Customer $($customer.TenantId): $($_.Exception.Message)" -ErrorAction Continue
    }
}
```

The token cache works the same way inside a function: a warm worker reuses tokens until five minutes before expiry. Each worker process (and each runspace, if you raise `PSWorkerInProcConcurrencyUpperBound`) has its own cache.

## Things that are never acceptable

- An admin username and password in app settings, Automation variables or code, including for "just one customer".
- Resource owner password credentials (ROPC) flows. They can't do MFA.
- A shared function key or API key that returns secrets to any caller who has it.
- A partner app client secret with a multi-year lifetime in app settings.
- Writing tokens to Table storage, blob storage, logs or Application Insights.

## Microsoft Learn references

- [PowerShell developer reference for Azure Functions](https://learn.microsoft.com/en-us/azure/azure-functions/functions-reference-powershell)
- [Azure Automation runbook types](https://learn.microsoft.com/en-us/azure/automation/automation-runbook-types)
- [Use Azure Key Vault with SecretManagement in automation](https://learn.microsoft.com/en-us/powershell/utility-modules/secretmanagement/how-to/using-azure-keyvault)
- [Azure Key Vault RBAC guide](https://learn.microsoft.com/en-us/azure/key-vault/general/rbac-guide)
- [App-only authentication in Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2)
- [Mandatory MFA for Azure](https://learn.microsoft.com/en-us/entra/identity/authentication/concept-mandatory-multifactor-authentication)
- [Mandating MFA for partner tenants (app-only Partner Center tokens)](https://learn.microsoft.com/en-us/partner-center/security/partner-security-requirements-mandating-mfa)
