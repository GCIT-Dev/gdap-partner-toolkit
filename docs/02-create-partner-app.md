# 02 Create the partner app

The partner app is a single multi-tenant app registration in **your partner tenant**. Every technician signs in to it, and it is pre-consented into each GDAP customer. You create it once.

## What gets created

| Object | Where | Detail |
| --- | --- | --- |
| Application (app registration) | Partner tenant | Multi-tenant (`AzureADMultipleOrgs`). Delegated permissions from the manifest you choose. No client secret. |
| Certificate credential | On the application object | Public key only. The private key stays on the technician workstation or in your vault. |
| Redirect URI | Web platform | `http://localhost`. Microsoft ignores the port when matching localhost redirect URIs, so the sign-in listener can use any free port. It must be the **Web** platform, not single-page application (SPA refresh tokens only last 24 hours) and not public client. |
| Service principal | Partner tenant | Required. Since March 2026, Microsoft Entra ID blocks sign-in for multi-tenant apps without a service principal in the tenant where they authenticate. |
| Admin consent | Partner tenant | Tenant-wide delegated consent for the manifest's scopes, so technicians aren't prompted and Partner Center calls work. |

What it does **not** do:

- It doesn't add application (app-only) permissions. Those belong in a separate automation app. See [05 Exchange access](05-exchange-access.md) and [manifests/README.md](../manifests/README.md).
- It doesn't turn off app instance property lock. New apps have it on by default since June 2026. Leave it on. MspGdap never adds credentials to the app's service principal in customer tenants.

## Before you start

- You need **Cloud Application Administrator** or **Application Administrator** in the partner tenant to register the app and grant tenant-wide consent to delegated permissions.
- Choose a manifest (below).
- Create a certificate (below).
- `New-MspPartnerApp` works against your partner tenant before the partner app exists, so it uses a Microsoft Graph PowerShell session for that one-off job. Install `Microsoft.Graph.Authentication` and sign in first:

```powershell
Connect-MgGraph -TenantId '<PartnerTenantId>' -Scopes 'Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All'
```

Microsoft Graph PowerShell 2.x on Windows signs in through the Web Account Manager (WAM) broker, which needs an interactive console window. Run this step in a normal PowerShell 7 window (not a background job, scheduled task, or an IDE or agent host without a console), otherwise it fails with "A window handle must be configured".

`AppRoleAssignment.ReadWrite.All` is only used for `-RestrictToGroupId` (below), but `New-MspPartnerApp` warns when it is missing. `New-MspPartnerApp` stops if the Graph session is in any tenant other than the `-PartnerTenantId` you give it.

## Choose a manifest

| Manifest | Use it when |
| --- | --- |
| `manifests/partner-app.minimal.json` | Default. 31 delegated permissions across Microsoft Graph, Partner Center, Exchange Online, the Office 365 Management API and Defender for Endpoint. Enough for this module and typical reporting and admin scripts. |
| `manifests/partner-app.full.json` | You want parity with a mature MSP worker app (Intune writes, Teams, SharePoint, Defender response actions, policy writes). 122 delegated permissions, a strict superset of the minimal manifest. Review every **High** entry in [manifests/README.md](../manifests/README.md) and remove what you don't need. |

Points to remember:

- The manifest is a ceiling. A technician can only use a permission if their GDAP roles in that customer also allow it.
- In the minimal manifest, `AppRoleAssignment.ReadWrite.All` and `RoleManagement.ReadWrite.Directory` exist only for `Enable-MspExchangeAppAccess`. Remove both if you won't use the optional automation app.
- `Connect-MspTeams` needs the Skype and Teams Tenant Admin API permission (`user_impersonation`), which is in the full manifest only. Microsoft's `Connect-MicrosoftTeams -AccessTokens` guidance also lists Graph delegated permissions such as `TeamSettings.ReadWrite.All`, `Channel.Delete.All`, `ChannelSettings.ReadWrite.All` and `ChannelMember.ReadWrite.All`. Add those to your copy of the manifest if you use the Teams module.
- The `resourceDisplayName` and `value` fields are documentation. `New-MspPartnerApp` strips them and checks every permission ID against the resource's published permission list before writing anything. A typo or a retired permission stops the run.

To make your own manifest, copy the minimal file, edit it and keep it with your own configuration.

## Create a certificate

Microsoft recommends certificate credentials over client secrets. MspGdap signs a short-lived client assertion with the certificate's private key each time it redeems a token. The assertion uses **PS256** (RSA-PSS with SHA-256) and the `x5t#S256` header, as Microsoft Learn now specifies. RS256 is kept as a fallback for keys that can't do PSS.

### Option A: one non-exportable certificate per technician workstation (recommended)

Each workstation creates its own key pair. The private key can't be exported, so it can't be copied off the machine. You upload each public key to the partner app, and you can revoke one workstation without touching the others.

```powershell
$params = @{
    Subject           = 'CN=MspGdap partner app - WS-JANE01'
    CertStoreLocation = 'Cert:\CurrentUser\My'
    Provider          = 'Microsoft Software Key Storage Provider'
    KeyAlgorithm      = 'RSA'
    KeyLength         = 3072
    HashAlgorithm     = 'SHA256'
    KeyExportPolicy   = 'NonExportable'
    KeyUsage          = 'DigitalSignature'
    NotAfter          = (Get-Date).AddMonths(12)
}
$cert = New-SelfSignedCertificate @params
$cert.Thumbprint

# Export the PUBLIC key only, to upload to the app
Export-Certificate -Cert $cert -FilePath ./partner-app-ws-jane01.cer -Type CERT
```

`New-SelfSignedCertificate` is part of the Windows PKI module, so this step is Windows only. On macOS or Linux, generate the key pair with your platform tooling or in Azure Key Vault (option B).

- `Microsoft Software Key Storage Provider` is a CNG provider and supports PS256.
- For TPM-backed keys, use `-Provider 'Microsoft Platform Crypto Provider'` instead (the key never leaves the TPM).
- The `.cer` file contains only the public key. Delete it after upload. The repository `.gitignore` blocks `*.cer`, `*.pfx` and `*.key` files.

### Option B: a certificate in Azure Key Vault

Use this for shared automation hosts (see [08 Unattended automation](08-unattended-automation.md)), or when you want central issuance and rotation. Create the certificate in Key Vault with a 12 month validity, download the public `.cer` from the vault and upload it to the app. Grant each host's managed identity only the access it needs on that vault.

### Exchange app-only is different

The optional automation app for app-only Exchange needs a certificate whose key is **not** CNG. Microsoft states "Cryptography: Next Generation (CNG) certificates aren't supported for app-only authentication with Exchange." Create that certificate with a legacy CSP, for example:

```powershell
$params = @{
    Subject           = 'CN=MspGdap automation app'
    CertStoreLocation = 'Cert:\CurrentUser\My'
    Provider          = 'Microsoft Enhanced RSA and AES Cryptographic Provider'
    KeySpec           = 'KeyExchange'
    KeyAlgorithm      = 'RSA'
    KeyLength         = 2048
    HashAlgorithm     = 'SHA256'
    KeyExportPolicy   = 'NonExportable'
    NotAfter          = (Get-Date).AddMonths(12)
}
New-SelfSignedCertificate @params
```

Do not reuse this certificate for the partner app. MspGdap falls back from PS256 to RS256 automatically when a legacy CSP key can't sign with PSS, but keep the two certificates separate anyway: they belong to different apps with different risk.

## Run New-MspPartnerApp

Preview first, then create:

```powershell
$appParams = @{
    DisplayName           = 'Contoso MSP Worker'
    PartnerTenantId       = '<PartnerTenantId>'
    ManifestPath          = './manifests/partner-app.minimal.json'
    CertificateThumbprint = '<Thumbprint>'
    RestrictToGroupId     = '<TechnicianGroupObjectId>'
}
New-MspPartnerApp @appParams -WhatIf
New-MspPartnerApp @appParams
```

`-RestrictToGroupId` is recommended. It turns on **Assignment required** for the app's service principal in your partner tenant and assigns one security group (your technicians' admin accounts). Nobody else in the partner tenant can then sign in to the app, so a phished ordinary user account can't obtain a partner app refresh token. Leave it out only if you manage the assignment yourself.

`New-MspPartnerApp`:

1. validates the manifest against each resource's published permissions,
2. creates the multi-tenant application with the Web redirect `http://localhost`,
3. uploads the certificate's public key,
4. creates the service principal in your partner tenant,
5. grants tenant-wide admin consent in the partner tenant for each resource in the manifest,
6. optionally restricts sign-in to your technician group (`-RestrictToGroupId`),
7. reads everything back and returns a result object with the app ID, the service principal ID and the status of each step, then prints the next steps. A `Failed` result is also written as an error.

If an app with the same name already exists, the command stops. `-UseExisting` reuses it instead: the manifest's permissions are **added** to the app's existing permissions and any missing redirect URI is added, keeping every other web setting. Nothing the app already has is removed, unless you also pass `-RemoveUnlisted` to make the permissions match the manifest exactly. The result warns about permissions the manifest does not list.

Record the **application (client) ID** and your **partner tenant ID**. Neither is a secret, but keep them in your own configuration, not in shared scripts.

Then configure the module on each workstation:

```powershell
Set-MspConfiguration -PartnerTenantId '<PartnerTenantId>' -AppId '<PartnerAppId>' -CertificateThumbprint '<Thumbprint>' -VaultName 'MspGdap'
Get-MspConfiguration
```

`Set-MspConfiguration` writes non-secret settings only (tenant ID, app ID, certificate thumbprint or path, vault name, signing algorithm, default technician) to `$HOME/.mspgdap/config.json`. A PFX password is kept for the current session only, and a client secret goes to your vault, never to the file.

### Adding another workstation or rotating a certificate

```powershell
Add-MspPartnerAppCertificate -AppId '<PartnerAppId>' -PartnerTenantId '<PartnerTenantId>' -CertificateThumbprint '<NewThumbprint>' -WhatIf
Add-MspPartnerAppCertificate -AppId '<PartnerAppId>' -PartnerTenantId '<PartnerTenantId>' -CertificateThumbprint '<NewThumbprint>'
```

Rotate at least a month before expiry. When every host uses the new certificate, remove the old one from the app in the Microsoft Entra admin center (App registrations, your app, Certificates and secrets). Keep the list of certificates on the app short and named so you can tell which workstation each belongs to.

### Client secrets

MspGdap accepts a client secret (as a `SecureString`) for partners who can't use certificates yet. It is discouraged: a secret can be copied anywhere and used from anywhere. If you must, keep it in your SecretManagement vault, give it the shortest practical lifetime and plan the move to a certificate.

## Manual equivalent in the Microsoft Entra admin center

If you prefer to create the app by hand, or your change process requires it:

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com) in the partner tenant as a Cloud Application Administrator.
2. Go to **Entra ID**, **App registrations**, **New registration**.
3. Name the app. Under **Supported account types**, choose accounts in any organisational directory (multi-tenant).
4. Under **Redirect URI**, choose **Web** and enter `http://localhost`. Select **Register**.
5. On the app's **Overview**, copy the **Application (client) ID** and **Directory (tenant) ID**.
6. Go to **Certificates and secrets**, **Certificates**, **Upload certificate**, and upload the `.cer` file.
7. Go to **API permissions**. Either add each delegated permission from your manifest, or open **Manifest** and replace `requiredResourceAccess` with the output of the script below.
8. Back in **API permissions**, select **Grant admin consent for** your partner tenant and confirm.
9. Go to **Enterprise applications** and confirm the app is listed (the service principal). If it is missing, create it before any technician signs in.
10. Optional but recommended: complete **publisher verification**, because customers see your app in their Enterprise applications list.

To convert a MspGdap manifest into plain `requiredResourceAccess` JSON for step 7:

```powershell
$manifest = Get-Content -Path ./manifests/partner-app.minimal.json -Raw | ConvertFrom-Json
$requiredResourceAccess = foreach ($resource in $manifest.requiredResourceAccess) {
    [ordered]@{
        resourceAppId  = $resource.resourceAppId
        resourceAccess = @($resource.resourceAccess | ForEach-Object { [ordered]@{ id = $_.id; type = $_.type } })
    }
}
ConvertTo-Json -InputObject @($requiredResourceAccess) -Depth 5
```

After a manual setup, run `Set-MspConfiguration` as above. You can check the result once a technician has registered a token:

```powershell
$partnerTenantId = (Get-MspConfiguration).PartnerTenantId
Get-MspAccessToken -PartnerTenant -Resource Graph | Test-MspAccessToken -TenantId $partnerTenantId -Resource Graph -Detailed
```

## Next steps

1. [Register a technician token](03-register-technician-token.md).
2. [Pre-consent the app into your customers](04-preconsent-customers.md).

## Microsoft Learn references

- [Microsoft identity platform certificate credentials](https://learn.microsoft.com/en-us/entra/identity-platform/certificate-credentials)
- [Authorisation code flow](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Redirect URI (reply URL) restrictions](https://learn.microsoft.com/en-us/entra/identity-platform/reply-url)
- [Retirement of service principal-less authentication](https://learn.microsoft.com/en-us/entra/identity-platform/retire-service-principal-less-authentication)
- [App instance property lock](https://learn.microsoft.com/en-us/entra/identity-platform/howto-configure-app-instance-property-locks)
- [New-SelfSignedCertificate](https://learn.microsoft.com/en-us/powershell/module/pki/new-selfsignedcertificate)
- [Grant tenant-wide admin consent](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/grant-admin-consent)
