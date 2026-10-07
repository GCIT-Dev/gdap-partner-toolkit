# 07 Migrating from DAP, MSOnline and older Secure App Model scripts

Many partner scripts written between 2015 and 2022 no longer work. They depended on things Microsoft has retired: DAP (delegated admin privileges), the MSOnline and AzureAD PowerShell modules, Azure AD Graph, basic authentication to Exchange Online, and admin passwords stored for automation. This guide maps each retired pattern to its MspGdap replacement, with code.

## What was retired

| Retired | Status | Replacement |
| --- | --- | --- |
| DAP (admin agent with Global Administrator in every customer) | Microsoft stopped granting DAP for new customers on 25 September 2023 and replaced it with GDAP | GDAP relationships with least-privilege roles ([01](01-prerequisites.md)) |
| MSOnline module (`Connect-MsolService`, `Get-MsolPartnerContract`, `-TenantId` on Msol cmdlets) | Deprecated 30 March 2024, with Microsoft committing only to keep it working until 30 March 2025. Treat it as retired | Microsoft Graph through `Invoke-MspGraphRequest` or `Connect-MspGraph` |
| AzureAD and AzureADPreview modules | Same deprecation as MSOnline | Microsoft Graph |
| Azure AD Graph (`graph.windows.net`) | Retired. Applications can no longer call it | Microsoft Graph (`graph.microsoft.com`) |
| Exchange remote PowerShell with basic authentication (`New-PSSession ... -Authentication Basic`, `?DelegatedOrg=`) | Basic authentication and Remote PowerShell for Exchange Online are retired | `Connect-MspExchangeOnline` (modern auth, REST-based module) |
| Admin passwords in scripts, AES key files or Azure Functions app settings | Insecure, and incompatible with mandatory MFA | Per-technician refresh tokens in a vault, or app-only certificates ([03](03-register-technician-token.md), [08](08-unattended-automation.md)) |
| Community `PartnerCenter` module and `New-PartnerAccessToken` | Microsoft describes it as "an open-source project maintained by the partner community and not officially supported by Microsoft" | `Register-MspPartnerToken` and `Get-MspAccessToken` (plain REST, no dependency) |

Before you migrate a script, make sure the prerequisites are in place: a GDAP relationship and role for the work the script does, the partner app consented in the customer, and a registered technician token.

## Pattern 1: looping over customers with MSOnline

**Before (retired, do not use):**

```powershell
Connect-MsolService
$customers = Get-MsolPartnerContract -All
foreach ($customer in $customers) {
    Get-MsolUser -TenantId $customer.TenantId -All | Where-Object { $_.IsLicensed }
}
```

**After:**

```powershell
Import-Module ./src/MspGdap/MspGdap.psd1
$customers = Get-MspCustomer -IncludeGdapStatus
foreach ($customer in $customers) {
    $users = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri 'v1.0/users?$select=id,userPrincipalName,assignedLicenses'
    $users | Where-Object { $_.assignedLicenses.Count -gt 0 } |
        Select-Object @{ Name = 'Customer'; Expression = { $customer.DisplayName } }, userPrincipalName
}
```

What changed:

- `Get-MspCustomer` replaces `Get-MsolPartnerContract`. It reads your customer contracts from your partner tenant through Microsoft Graph, with their tenant IDs. Add `-IncludeGdapStatus` to include GDAP-only customers and their relationship status.
- There is no partner-wide session. Each call names its customer with `-TenantId`, and the token for that customer comes from the customer's own endpoint.
- What the loop can do in each customer depends on your GDAP roles there. If a customer has no active relationship or you have no suitable role, that customer returns an error instead of silently using another tenant.

### Common MSOnline cmdlets and their Graph equivalents

| MSOnline | Microsoft Graph request through `Invoke-MspGraphRequest` |
| --- | --- |
| `Get-MsolUser -TenantId $t -All` | `GET v1.0/users` (all pages are returned by default) |
| `Get-MsolUser -UserPrincipalName $upn` | `GET v1.0/users/{upn}` |
| `Get-MsolAccountSku -TenantId $t` | `GET v1.0/subscribedSkus` |
| `Set-MsolUserLicense -AddLicenses` | `POST v1.0/users/{id}/assignLicense` |
| `Get-MsolDomain -TenantId $t` | `GET v1.0/domains` |
| `Get-MsolCompanyInformation -TenantId $t` | `GET v1.0/organization` |
| `Get-MsolRole` and `Get-MsolRoleMember` | `GET v1.0/roleManagement/directory/roleDefinitions` and `GET v1.0/roleManagement/directory/roleAssignments?$expand=principal` |
| `Set-MsolUser -BlockCredential $true` | `PATCH v1.0/users/{id}` with `accountEnabled` set to false |

Every write supports `-WhatIf`:

```powershell
$body = @{ accountEnabled = $false }
Invoke-MspGraphRequest -TenantId 'fabrikam.onmicrosoft.com' -Method PATCH -Uri 'v1.0/users/leaver@fabrikam.com' -Body $body -WhatIf
```

## Pattern 2: DAP and the admin agent role

Under DAP, membership of AdminAgents gave Global Administrator in every customer, permanently. Scripts relied on that.

**After:** access comes from GDAP relationships and security group assignments. Inventory what you have, then create least-privilege relationships to replace anything broad:

```powershell
Get-MspGdapRelationship | Sort-Object EndDateTime | Format-Table CustomerName, DisplayName, Status, EndDateTime
New-MspGdapRelationship -TenantId 'fabrikam.onmicrosoft.com' -DisplayName 'Contoso-Fabrikam-Std-2026' -AccessMapPath ./gdap-access.json -Duration 'P730D' -AutoExtend -WhatIf
```

AdminAgents still matters, but only for calling the Partner Center API (for example `Grant-MspPartnerAppConsent`). It no longer grants access to customer data.

## Pattern 3: Exchange remote PowerShell with basic authentication

**Before (retired, do not use):**

```powershell
$credential = Get-Credential
$session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid?DelegatedOrg=$customerDomain" -Credential $credential -Authentication Basic -AllowRedirection
Import-PSSession $session
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Remove-PSSession $session
```

**After:**

```powershell
Connect-MspExchangeOnline -TenantId 'fabrikam.onmicrosoft.com'
Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
Get-AdminAuditLogConfig | Select-Object UnifiedAuditLogIngestionEnabled
Disconnect-ExchangeOnline -Confirm:$false
```

The technician needs Exchange Administrator in the customer through GDAP. See [05 Exchange access](05-exchange-access.md). Scripts that created a dedicated Exchange admin account in each customer for automation should use the app-only path instead (`Enable-MspExchangeAppAccess`), which needs no user account or password in the customer.

## Pattern 4: the AzureAD module

**Before (retired, do not use):**

```powershell
Connect-AzureAD -TenantId $customerTenantId
Get-AzureADUser -All $true | Where-Object { $_.AccountEnabled }
```

**After**, either with MspGdap's own request command:

```powershell
Invoke-MspGraphRequest -TenantId 'fabrikam.onmicrosoft.com' -Method GET -Uri 'v1.0/users?$filter=accountEnabled eq true&$select=id,userPrincipalName'
```

or with the Microsoft Graph PowerShell SDK, using MspGdap's token:

```powershell
Connect-MspGraph -TenantId 'fabrikam.onmicrosoft.com'
Get-MgUser -All -Filter 'accountEnabled eq true' -Property id, userPrincipalName
```

`Connect-MspGraph` passes a cached access token to `Connect-MgGraph -AccessToken`. The SDK does not refresh that token, so run `Connect-MspGraph` again for long sessions. `Get-MgUser` needs the `Microsoft.Graph.Users` module.

## Pattern 5: Secure App Model scripts built on Azure AD Graph

Early Secure App Model scripts used the community PartnerCenter module to mint two tokens from a stored refresh token, one of them for Azure AD Graph, and passed them to `Connect-AzureAD`.

**Before (retired, do not use):**

```powershell
$aadGraphToken = New-PartnerAccessToken -ApplicationId $appId -Credential $appCredential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $customerTenantId
$graphToken = New-PartnerAccessToken -ApplicationId $appId -Credential $appCredential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $customerTenantId
Connect-AzureAD -AadAccessToken $aadGraphToken.AccessToken -MsAccessToken $graphToken.AccessToken -AccountId $upn -TenantId $customerTenantId
```

Problems with this pattern, beyond the retired APIs: the refresh token and app secret were often read from a plain-text file or variable, the rotated refresh token was thrown away, and every call minted new tokens.

**After:**

```powershell
Register-MspPartnerToken -UserPrincipalName 'jane.admin@contoso-msp.onmicrosoft.com'   # once
Invoke-MspGraphRequest -TenantId 'fabrikam.onmicrosoft.com' -Method GET -Uri 'v1.0/organization'
```

The token is stored in your vault, rotated on every use, cached per customer and validated. The app uses a certificate rather than a secret. Your existing Secure App Model app can be retired once the MspGdap partner app is consented in every customer.

## Pattern 6: passwords in AES key files

**Before (do not use):**

```powershell
$key = Get-Content -Path .\aes.key
$password = Get-Content -Path .\password.txt | ConvertTo-SecureString -Key $key
$credential = New-Object System.Management.Automation.PSCredential ('admin@fabrikam.onmicrosoft.com', $password)
```

A key file next to an encrypted password file is a plain-text password with an extra step. Anyone who can read both files has the password. It also cannot satisfy MFA.

**After:** there is no password. Interactive work uses the technician's refresh token from the vault. Unattended work uses a certificate, either through a managed identity and Key Vault or app-only access. If an old script still needs some other secret (an API key for a third-party tool, for example), store it with SecretManagement:

```powershell
Set-Secret -Name 'ThirdPartyApiKey' -Vault 'MspGdap' -SecureStringSecret (Read-Host -AsSecureString -Prompt 'API key')
$apiKey = Get-Secret -Name 'ThirdPartyApiKey' -Vault 'MspGdap'
```

## Pattern 7: Azure Functions v1 with stored admin credentials

Old automation often ran in an Azure Functions v1 app (Windows PowerShell), with a customer or partner admin username and password in the app settings, connecting with `Connect-MsolService -Credential`.

**After:** Azure Functions v4 on PowerShell 7.6 (or Azure Automation), with a managed identity that reads a certificate or refresh token from Key Vault. App-only access is preferred where Microsoft supports it. See [08 Unattended automation](08-unattended-automation.md) for the full pattern and its trade-offs.

## Ready-made rewrites of common partner scripts

Before you rewrite a script yourself, check [scripts/README.md](../scripts/README.md). It lists working MspGdap versions of the multi-tenant scripts GCIT published between 2017 and 2020 (Exchange settings, licence and user reports, admin and MFA audits, forwarding and audit log checks, Secure Score and alert reports, and the Azure Functions jobs), each mapped to the original article and the retired method it replaced. They follow the patterns in this guide: `-TenantId` or `-AllCustomers`, no stored passwords, report only unless you add `-Apply`, and `-WhatIf` on every change. Articles about retired products or features are listed there too, with the nearest modern alternative.

## Migration checklist

1. List every script and scheduled job that uses MSOnline, AzureAD, Azure AD Graph, `New-PSSession` to Exchange, stored passwords or DAP.
2. For each, write down what it does in each customer and which GDAP role that needs.
3. Set up the prerequisites ([01](01-prerequisites.md)), create the partner app ([02](02-create-partner-app.md)) and register technician tokens ([03](03-register-technician-token.md)).
4. Pre-consent the app into every customer ([04](04-preconsent-customers.md)).
5. Rewrite each script using the patterns above, or start from a ready-made rewrite in [scripts/](../scripts/README.md). Run it with `-WhatIf` against one low-risk customer first.
6. Move unattended jobs to [08](08-unattended-automation.md).
7. Delete stored passwords and key files, rotate any credentials they contained, and remove old apps and per-customer admin accounts created for automation.

## Microsoft Learn references

- [Azure AD PowerShell deprecation notice](https://learn.microsoft.com/en-us/powershell/azure/active-directory/overview)
- [GDAP frequently asked questions (DAP to GDAP)](https://learn.microsoft.com/en-us/partner-center/customers/gdap-faq)
- [Microsoft Graph PowerShell overview](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview)
- [Connect-MgGraph](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph)
- [Connect to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)
- [Enable the Secure Application Model](https://learn.microsoft.com/en-us/partner-center/developer/enable-secure-app-model)
