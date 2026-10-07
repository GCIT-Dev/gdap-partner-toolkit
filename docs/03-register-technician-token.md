# 03 Register a technician token

Each technician signs in to the partner app once with their dedicated admin account. MspGdap stores the resulting refresh token in a SecretManagement vault and uses it to get access tokens for any customer and resource. This guide covers where to store it, how it rotates and how to revoke it.

## How registration works

`Register-MspPartnerToken`:

1. starts a temporary listener on `http://localhost:<free port>`,
2. opens the system browser at the partner tenant's authorise endpoint with a PKCE challenge (`S256`), a random `state` value and your UPN as a login hint,
3. you sign in with your admin account and complete MFA (Conditional Access in your partner tenant decides what is required),
4. the listener receives the authorisation code, checks `state` and redeems the code straight away (codes expire after about a minute) with the PKCE verifier and a certificate-signed client assertion. The listener only answers requests from this computer: a request from another address on the network gets a 403 and is ignored,
5. it checks that the signed-in account matches `-UserPrincipalName`, that an id_token came back carrying this sign-in's `nonce`, and that the sign-in used MFA (`amr` claim contains `mfa`). If neither token exposes `amr`, MFA can't be confirmed and nothing is stored unless you add `-SkipMfaCheck` (only do that when you are sure Conditional Access enforces MFA for the account), then
6. writes the refresh token to your vault as a `SecureString` and returns a summary object. The token itself is never shown.

Device code flow is not used. Microsoft recommends blocking it with Conditional Access, and many partners already do.

```powershell
Register-MspPartnerToken -UserPrincipalName 'jane.admin@contoso-msp.onmicrosoft.com' -SetAsDefault
```

`-SetAsDefault` saves the account as the default technician (`TechnicianUpn`) in your configuration. If the browser can't be opened (for example over remote PowerShell), add `-NoBrowser` and open the printed URL yourself on the same machine.

Check that it worked:

```powershell
$partnerTenantId = (Get-MspConfiguration).PartnerTenantId
Get-MspAccessToken -PartnerTenant -Resource Graph | Test-MspAccessToken -TenantId $partnerTenantId -Resource Graph -Detailed
Get-SecretInfo -Vault 'MspGdap' -Name 'MspGdap-*'
```

## Secret naming

The refresh token is stored as:

```text
MspGdap-<partner app ID>-<first 16 hex characters of SHA-256 of the lower-case UPN>
```

For example `MspGdap-<PartnerAppId>-3f1c9a0b7d2e4c61`. The UPN is not in the name because:

- Azure Key Vault object names allow only letters, numbers and hyphens, 1 to 127 characters. A UPN contains `@` and `.`, and can be long.
- Key Vault warns that names and identifiers may be copied globally and must not contain personal information.

The UPN, the app ID and the created and last-used times are kept as secret **metadata** where the vault supports it. Metadata is not stored securely, so it never holds anything sensitive. If a vault does not support metadata, the module stores the secret without it.

With the `Az.KeyVault` extension vault, metadata is stored as Key Vault **tags**, which are plain text and visible to anyone who can list the vault's secrets (for example with the Key Vault Reader role), even without permission to read secret values. The technician's UPN is therefore visible to those readers. That is acceptable for most partners, but keep list rights on the vault as narrow as the secret rights.

To work out a technician's secret name yourself:

```powershell
$upn = 'jane.admin@contoso-msp.onmicrosoft.com'
$bytes = [System.Text.Encoding]::UTF8.GetBytes($upn.ToLowerInvariant())
$hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
$suffix = (-join ($hash | ForEach-Object { $_.ToString('x2') })).Substring(0, 16)
"MspGdap-<PartnerAppId>-$suffix"
```

## Choose a vault

MspGdap only talks to `Microsoft.PowerShell.SecretManagement` (`Get-Secret`, `Set-Secret`, `Get-SecretInfo`, `Remove-Secret`). Any registered extension vault works. Two are documented here.

Microsoft has declared the SecretManagement and SecretStore modules feature complete and archived their repositories. They still receive security and critical bug fixes. The latest versions are SecretManagement 1.1.2 and SecretStore 1.0.6. MspGdap keeps all vault access behind one internal wrapper so another backend can be added later.

### Option 1: SecretStore on the technician's workstation

Best for individual technicians on managed devices.

```powershell
Install-Module -Name Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore -Repository PSGallery -Scope CurrentUser
Register-SecretVault -Name 'MspGdap' -ModuleName 'Microsoft.PowerShell.SecretStore' -DefaultVault
Set-SecretStoreConfiguration -Scope CurrentUser -Authentication Password -PasswordTimeout 900 -Interaction Prompt
```

- **Run these commands in an interactive PowerShell window.** The first `Register-SecretVault` or `Set-SecretStoreConfiguration` for a new SecretStore prompts for a password at the console. A background job, scheduled task or other non-interactive host cannot answer that prompt, so the command fails or waits there. Run `Register-MspPartnerToken` in the same interactive window afterwards.
- **Always use password authentication.** `-Authentication None` is meant for testing only.
- `-PasswordTimeout` is in seconds. 900 (15 minutes) is Microsoft's default. Shorter is safer, longer is more convenient. Pick a value and keep it consistent across your team.
- The store is encrypted per user, per machine. It does not roam. If a workstation is rebuilt, register again.
- MspGdap writes the rotated refresh token back every time it redeems one, so the store must be unlocked when you request a new access token. Run `Unlock-SecretStore` at the start of a session if you set `-Interaction None`.

### Option 2: Azure Key Vault through Az.KeyVault

Best when technicians work from several machines, when you need central audit logs, or for automation hosts (see [08 Unattended automation](08-unattended-automation.md)).

```powershell
Install-Module -Name Microsoft.PowerShell.SecretManagement, Az.Accounts, Az.KeyVault -Repository PSGallery -Scope CurrentUser
Connect-AzAccount -Tenant '<PartnerTenantId>'

$vaultParameters = @{
    AZKVaultName   = '<KeyVaultName>'
    SubscriptionId = '<SubscriptionId>'
}
Register-SecretVault -Name 'MspGdapKv' -ModuleName 'Az.KeyVault' -VaultParameters $vaultParameters
Set-MspConfiguration -VaultName 'MspGdapKv'
```

Use the **Azure RBAC** permission model on the vault (the default for vaults created with API version 2026-02-01 or later). Then decide how to separate technicians.

#### Per-technician vault (recommended)

Microsoft recommends a vault per application per environment with roles assigned at the vault scope. For refresh tokens, that means one small vault per technician:

```powershell
$scope = '/subscriptions/<SubscriptionId>/resourcegroups/<ResourceGroup>/providers/Microsoft.KeyVault/vaults/<TechnicianVaultName>'
New-AzRoleAssignment -RoleDefinitionName 'Key Vault Secrets Officer' -SignInName 'jane.admin@contoso-msp.onmicrosoft.com' -Scope $scope
```

**Key Vault Secrets Officer** can "perform any action on the secrets of a key vault, except managing permissions". The technician can read and rewrite their own token and nobody else's. Keep vault-level administrative roles (Owner, User Access Administrator, Key Vault Administrator, Key Vault Data Access Administrator) for a small break-glass group.

#### Shared vault with secret-scope assignments

Microsoft generally advises against role assignments on individual secrets, but lists "individual secrets require individual user access" as an exception, which is this case. An administrator creates each technician's secret first (with a placeholder value), then assigns the role at the secret's scope:

```powershell
$scope = '/subscriptions/<SubscriptionId>/resourcegroups/<ResourceGroup>/providers/Microsoft.KeyVault/vaults/<SharedVaultName>/secrets/MspGdap-<PartnerAppId>-<hash>'
New-AzRoleAssignment -RoleDefinitionName 'Key Vault Secrets Officer' -SignInName 'jane.admin@contoso-msp.onmicrosoft.com' -Scope $scope
```

Limits of this model:

- Anyone with a vault-level data role can read every technician's token. Microsoft notes that object-scope assignments do not isolate teams from vault administrators.
- Without vault-level list permission, `Get-SecretInfo` may not list secrets. Test your registration and token refresh end to end before rolling out.
- Key Vault also offers attribute-based conditions (ABAC, in preview) that can limit a principal to secret names matching a pattern. Consider them if you outgrow per-secret assignments.

#### Never use a shared key to the vault

A pattern seen in the field is an Azure Function or API in front of the vault, protected by one shared function key with read **and** write on every secret. Anyone holding that key can read or replace any technician's refresh token, and the vault's audit log shows only the function's identity. Do not build this. Each technician (or each automation host) authenticates to Key Vault as itself.

#### Harden the vault

- Turn on soft delete and purge protection.
- Restrict network access (private endpoint or firewall rules).
- Send `AuditEvent` diagnostic logs to Log Analytics and alert on secret reads by unexpected identities.

## Rotation

Microsoft Entra refresh tokens for this kind of app last **90 days** and replace themselves on every use. Microsoft does not revoke the old token when a new one is issued and asks clients to "securely delete the old refresh token after acquiring a new one."

MspGdap therefore:

- writes the new refresh token back to the vault after **every** redemption (the vault keeps the new value, Key Vault keeps the old one as a previous version),
- updates the `lastUsed` metadata, and
- warns loudly if the write-back fails. The previous token stays valid until it expires, so fix the vault before your next session.

Because the lifetime rolls forward on each use, a technician who uses the module at least once every 90 days never needs to register again. After 90 idle days, or if the token is revoked, `Get-MspAccessToken` fails with a clear message and you run `Register-MspPartnerToken` again.

If your Conditional Access sign-in frequency applies to the partner app, re-registration follows that schedule instead. See [Conditional Access for admin accounts](01-prerequisites.md#conditional-access-for-admin-accounts).

## Revocation

| Event | Effect on the stored refresh token |
| --- | --- |
| Technician changes their password | **Stays valid.** Microsoft documents that a password change does not revoke tokens for confidential clients. |
| Administrator revokes all refresh tokens (sign-in sessions) for the user | Revoked |
| Account disabled or deleted | Token can no longer be redeemed |
| 90 days without use | Expired |
| Partner app certificate removed | Token can't be redeemed from hosts that used that certificate |

To revoke a technician's access immediately, revoke their sessions in the partner tenant and remove the stored token with `Unregister-MspPartnerToken`. It works out the secret name from the UPN, removes the secret from the vault and clears that technician's cached access tokens:

```powershell
$user = Invoke-MspGraphRequest -PartnerTenant -Method GET -Uri 'v1.0/users/jane.admin@contoso-msp.onmicrosoft.com?$select=id,userPrincipalName'
Invoke-MspGraphRequest -PartnerTenant -Method POST -Uri "v1.0/users/$($user.id)/revokeSignInSessions" -Confirm
Unregister-MspPartnerToken -UserPrincipalName 'jane.admin@contoso-msp.onmicrosoft.com' -WhatIf
Unregister-MspPartnerToken -UserPrincipalName 'jane.admin@contoso-msp.onmicrosoft.com'
```

Removing the vault copy does not revoke tokens Microsoft already issued, which is why the session revocation comes first.

You can also select **Revoke sessions** on the user's page in the Microsoft Entra admin center.

## When a technician leaves

Work through this list on their last day. Record each step in your offboarding ticket.

1. Disable the technician's admin account in the partner tenant.
2. Revoke their sign-in sessions (above).
3. Remove them from every GDAP security group and from AdminAgents.
4. Remove their refresh token with `Unregister-MspPartnerToken -UserPrincipalName <their UPN>`, or delete their per-technician Key Vault (then purge it after your retention period).
5. Remove their Key Vault role assignments.
6. Remove their workstation's certificate from the partner app (App registrations, Certificates and secrets).
7. On any shared host they used, run `Disconnect-Msp` to clear cached tokens.
8. Review the partner app's sign-in logs and customer audit logs for their account over the last 30 days.

## What the module never does with your token

- Never stores it in `$global:` or any variable outside the module's private scope.
- Never writes it to disk in plain text, to the console, or to verbose, debug or information streams.
- Never sends it anywhere except `login.microsoftonline.com`.
- Never returns it from a public command.

See [SECURITY.md](../SECURITY.md) for the full list.

## Microsoft Learn references

- [Refresh tokens in the Microsoft identity platform](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens)
- [Authorisation code flow with PKCE](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [SecretManagement overview](https://learn.microsoft.com/en-us/powershell/utility-modules/secretmanagement/overview)
- [Use Azure Key Vault with SecretManagement](https://learn.microsoft.com/en-us/powershell/utility-modules/secretmanagement/how-to/using-azure-keyvault)
- [Set-SecretStoreConfiguration](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.secretstore/set-secretstoreconfiguration)
- [Set-Secret](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.secretmanagement/set-secret)
- [Azure Key Vault RBAC guide](https://learn.microsoft.com/en-us/azure/key-vault/general/rbac-guide)
- [Key Vault object naming](https://learn.microsoft.com/en-us/azure/key-vault/general/about-keys-secrets-certificates)
