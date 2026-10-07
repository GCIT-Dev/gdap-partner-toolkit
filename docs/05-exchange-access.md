# 05 Exchange access

There are two ways to run Exchange Online PowerShell against a customer with MspGdap:

| | Delegated (default) | App-only (optional) |
| --- | --- | --- |
| Who acts | The signed-in technician | A separate automation app |
| Limited by GDAP | Yes. The technician's GDAP roles apply | **No.** The app's own role in the customer applies |
| Credential | Technician's refresh token plus the partner app certificate | Automation app certificate only |
| Set-up per customer | Pre-consent (`Exchange.Manage` delegated scope) | `Exchange.ManageAsApp` app role **and** an Entra role on the app's service principal |
| Use for | Day-to-day admin, investigations, ad hoc scripts | Scheduled reports and unattended jobs |
| Command | `Connect-MspExchangeOnline -TenantId` | `Enable-MspExchangeAppAccess`, then `Connect-MspExchangeOnline -AppOnly -AppId` |

Use delegated access wherever a person is present. Use app-only only for jobs that must run unattended, and keep its roles as narrow as the job allows.

## Delegated Exchange Online

```powershell
Connect-MspExchangeOnline -TenantId 'fabrikam.onmicrosoft.com'
Get-OrganizationConfig | Select-Object DisplayName, IsDehydrated
Get-Mailbox -ResultSize 10
```

What `Connect-MspExchangeOnline` does:

1. closes any existing Exchange Online sessions (unless you pass `-KeepExistingConnections`), so a command can't land in the previous customer,
2. gets (or reuses from cache) a delegated access token for the customer with the scope `https://outlook.office365.com/.default`,
3. resolves the customer's initial `.onmicrosoft.com` domain (or uses `-Organization` if you pass it),
4. calls `Connect-ExchangeOnline` with `-AccessToken` and `-DelegatedOrganization`, and
5. checks that the session is connected to the tenant you asked for, and disconnects it with an error if not.

It refuses your own partner tenant as `-TenantId`, for delegated and app-only connections alike.

Requirements:

- `ExchangeOnlineManagement` **3.1.0 or later** (the `-AccessToken` parameter arrived in 3.1.0). Module 3.10.0 and later require PowerShell 7.6. On PowerShell 7.4 or 7.5, stay on 3.9.2 or earlier. `Connect-MspExchangeOnline` checks the installed version against your PowerShell version and explains any mismatch.
- The partner app consented in the customer with `Exchange.Manage` on Office 365 Exchange Online (`00000002-0000-0ff1-ce00-000000000000`). Both manifests include it.
- A GDAP role that allows the work: **Exchange Administrator** for changes, **Exchange Recipient Administrator** for recipient work, **Global Reader** for read-only.

About the documentation: Microsoft's parameter reference for `Connect-ExchangeOnline` says `-AccessToken` is used "with the Organization, DelegatedOrganization, or UserPrincipalName parameters" depending on the token type, and that `-DelegatedOrganization` accepts "the primary .onmicrosoft.com domain or tenant ID of the customer organization". Microsoft's current GDAP connection example uses interactive sign-in. The only Microsoft sample that combines an access token with `-DelegatedOrganization` is a retired Partner Center page. The combination is supported by the parameter documentation and works, but it is sparsely documented. If Microsoft changes it, `Connect-MspExchangeOnline` fails with a clear error rather than connecting to the wrong place.

Session notes:

- An access token lasts 60 to 90 minutes. When the Exchange session's token expires, run `Connect-MspExchangeOnline` again. MspGdap returns a cached token if it is still valid for more than five minutes, otherwise it redeems the refresh token silently.
- Exchange cmdlets have no tenant parameter, so the session decides where a change lands. That is why `Connect-MspExchangeOnline` closes other sessions first. If you use `-KeepExistingConnections` (with `-Prefix`), keep track of which prefix belongs to which customer. `Disconnect-Msp` closes the Exchange Online, Security and Compliance, Microsoft Graph SDK and Teams sessions that `Connect-Msp*` opened. Exchange sessions are closed by connection ID, so Exchange sessions you opened yourself stay open.

## App-only Exchange for unattended jobs

App-only access uses a **separate** multi-tenant automation app, never the partner app. The reasons are in [manifests/README.md](../manifests/README.md#application-permissions-use-a-separate-app): an app-only token is not limited by GDAP, so mixing it into the partner app would give the partner app standing, tenant-wide access in every customer.

### 1. Create the automation app

Create it in your partner tenant the same way as the partner app ([02 Create the partner app](02-create-partner-app.md)), with these differences:

- Permissions from `manifests/automation-app.example.json` (application permissions, `"type": "Role"`), trimmed to what your jobs need.
- No redirect URI.
- A certificate whose key is **not** CNG. Microsoft states "Cryptography: Next Generation (CNG) certificates aren't supported for app-only authentication with Exchange." See the CSP example in [02](02-create-partner-app.md#exchange-app-only-is-different).
- No client secret.

### 2. Enable it in each customer

Partner Center pre-consent can't grant application permissions under GDAP. Microsoft states "Explicit application-only consent to a customer tenant isn't supported for third-party application developers who use GDAP." Each customer needs:

1. the automation app's service principal,
2. the `Exchange.ManageAsApp` app role (`dc50a0fb-09a3-484d-be87-e023b12c6440`) on Office 365 Exchange Online, assigned to that service principal, and
3. an Entra role assigned **directly** to that service principal. Microsoft says "With GDAP, you can no longer add Microsoft Entra roles to an application by using relationships or through inclusion in a security group."

`Enable-MspExchangeAppAccess` does all three as the technician, through Microsoft Graph:

```powershell
$enableParams = @{
    TenantId = 'fabrikam.onmicrosoft.com'
    AppId    = '<AutomationAppId>'
    Role     = 'Exchange Administrator'
}
Enable-MspExchangeAppAccess @enableParams -WhatIf
Enable-MspExchangeAppAccess @enableParams

Test-MspExchangeAppAccess -TenantId 'fabrikam.onmicrosoft.com' -AppId '<AutomationAppId>'
Test-MspExchangeAppAccess -TenantId 'fabrikam.onmicrosoft.com' -AppId '<AutomationAppId>' -TestConnection -CertificateThumbprint '<Thumbprint>'
```

| Step | Graph call | GDAP role the technician needs |
| --- | --- | --- |
| Service principal | Read, or create from the app ID. If it can't be created, the result includes the admin consent URL for a Cloud Application Administrator to open | Cloud Application Administrator |
| `Exchange.ManageAsApp` | `POST /servicePrincipals/{exchange-sp-id}/appRoleAssignedTo` | Cloud Application Administrator (it can grant app roles for any API except Microsoft Graph app roles) |
| Entra role | `POST /roleManagement/directory/roleAssignments` with the role template ID, the service principal's object ID and `directoryScopeId` `/` | **Privileged Role Administrator** |

Design rules that `Enable-MspExchangeAppAccess` follows:

- It uses the unified role management API. It does not activate roles through the legacy `POST /directoryRoles` endpoint, which Microsoft says to replace with the unified RBAC API.
- Every step is read back. The result has `Success`, `Outcome` and one entry per step (service principal, app role assignment, role assignment). A step is `Passed` when it was already in place, `Changed` only when the change was made **and** confirmed by readback, and `Failed` or `Unknown` otherwise. `Success` is false if any step is `Failed` or `Unknown`. It never reports a pass for a failed step.
- It never adds credentials to the service principal in the customer tenant. Credentials live only on the application object in your partner tenant.
- The default `-Role` is **Exchange Administrator** (`29232cdf-9323-42fd-ade2-1d097af3e4de`). For narrower jobs use **Exchange Recipient Administrator** (`31392ffb-586c-42d1-9346-e59415a2cc4e`) or **Global Reader** (`f2ef992c-3afb-46b9-b7cf-a126ee74c451`). Global Administrator is refused unless you add `-AllowGlobalAdministrator`, and Microsoft says it should be limited to emergencies.
- It refuses to run against the MspGdap partner app. App-only permissions stay in their own app.
- Before any write it checks that the app is yours: either its service principal in the customer is owned by your partner tenant (`appOwnerOrganizationId`), or, when there is no service principal yet, the app is registered in your partner tenant. A typo or a third-party app ID is refused. `-AllowExternalApp` overrides this for an automation app you own in another tenant.
- Roles are limited to the ones Microsoft lists for Exchange Online app-only access. Exchange Administrator, Exchange Recipient Administrator, Global Reader and Security Reader are accepted as they are. Compliance Administrator, Security Administrator and Helpdesk Administrator need `-AllowPrivilegedRole`, because they give the app standing access beyond Exchange. Any other role (for example Privileged Role Administrator) is refused, because it would do nothing for Exchange and a lot elsewhere. Role template IDs that are not in the role catalogue are refused.
- A `Failed` result is also written as a non-terminating error, so `-ErrorAction Stop` and `try`/`catch` work in automation.

Keep Privileged Role Administrator out of standing GDAP access. Put it in its own security group with just-in-time membership, and use it only while running `Enable-MspExchangeAppAccess`.

If a customer won't allow these roles, their own Global Administrator or Privileged Role Administrator can grant the app role through the admin consent URL Microsoft documents for Exchange app-only, then assign the role in their Microsoft Entra admin center:

```text
https://login.microsoftonline.com/<customer tenant ID>/adminconsent?client_id=<AutomationAppId>&scope=https://outlook.office365.com/.default
```

### 3. Connect app-only

```powershell
Connect-MspExchangeOnline -TenantId 'fabrikam.onmicrosoft.com' -AppOnly -AppId '<AutomationAppId>' -CertificateThumbprint '<Thumbprint>'
```

This resolves the customer's initial `.onmicrosoft.com` domain (Microsoft requires it for `-Organization`) and runs the documented app-only connection, which you can also run directly:

```powershell
Connect-ExchangeOnline -AppId '<AutomationAppId>' -CertificateThumbprint '<Thumbprint>' -Organization 'fabrikam.onmicrosoft.com' -ShowBanner:$false
```

`-CertificateThumbprint` works on Windows only. On Linux and macOS (including Azure Functions on Linux), pass an `X509Certificate2` object with `-Certificate` instead, to either `Connect-MspExchangeOnline -AppOnly` or `Connect-ExchangeOnline`. See [08 Unattended automation](08-unattended-automation.md).

App-only limits to know:

- Some Microsoft 365 Group cmdlets don't work app-only (`New-UnifiedGroup`, `Remove-UnifiedGroup`, `Add-UnifiedGroupLinks`, `Remove-UnifiedGroupLinks`). Use the delegated path or Microsoft Graph for those.
- The app's role, not GDAP, decides what it can do. Review the role assignments in each customer as part of your quarterly access review (`Test-MspExchangeAppAccess` lists them).

### Removing app-only access from a customer

Delete the Entra role assignment and the app role assignment in the customer (or delete the automation app's service principal from their Enterprise apps). Then run `Test-MspExchangeAppAccess` and confirm every step reports absent.

## Security and Compliance PowerShell

`Connect-MspSecurityCompliance` connects Security and Compliance PowerShell (Microsoft Purview) with the technician's delegated GDAP token. A live test on 7 October 2026 confirmed that it works: it connected to a GDAP customer with the default settings below and ran `Get-RetentionCompliancePolicy`, with no consent beyond the normal pre-consent and no extra manifest entry. Microsoft documents this combination only briefly, so try one customer after you upgrade `ExchangeOnlineManagement`.

```powershell
Connect-MspSecurityCompliance -TenantId 'fabrikam.onmicrosoft.com'
Get-RetentionCompliancePolicy | Select-Object Name, Enabled
```

What is and isn't documented:

- `Connect-IPPSSession -AccessToken` exists from `ExchangeOnlineManagement` 3.8.0-Preview1, and is used "with the Organization, DelegatedOrganization, or UserPrincipalName parameters" depending on the token type.
- For `-DelegatedOrganization`, Microsoft says to use the primary `.onmicrosoft.com` domain and that "You must use the AzureADAuthorizationEndpointUri parameter with this parameter."
- Microsoft does **not** document which token audience `Connect-IPPSSession -AccessToken` expects. Security and Compliance PowerShell is a different resource from Exchange Online (Microsoft's app-only article uses Microsoft Exchange Online Protection, not Office 365 Exchange Online). By default `Connect-MspSecurityCompliance` passes a delegated token for `https://ps.compliance.protection.outlook.com`, the audience the module itself requests, with `-Organization` set to the customer's initial `.onmicrosoft.com` domain. Learn only documents `-Organization` for certificate connections, but this pairing is the one confirmed in the live test. `-TokenResource` changes the audience (for example to `https://outlook.office365.com`), and `-UseDelegatedOrganization` switches to `-DelegatedOrganization` (you then supply `-AzureADAuthorizationEndpointUri`, which must be a `https://login.microsoftonline.com` or `https://login.microsoftonline.us` address). Those two alternatives were not part of the live test. If a connection fails, connect interactively with `Connect-IPPSSession -UserPrincipalName <your admin UPN> -DelegatedOrganization <customer>.onmicrosoft.com -AzureADAuthorizationEndpointUri 'https://login.microsoftonline.com/organizations'` instead.

For unattended Security and Compliance jobs, use Microsoft's documented app-only connection (`Connect-IPPSSession -AppId <id> -CertificateThumbprint <thumbprint> -Organization <customer>.onmicrosoft.com`) with the permission and role Microsoft lists for Security and Compliance PowerShell in its app-only article. MspGdap does not automate that role in this release.

Avoid certificate files with passwords in scripts. Microsoft's own note on `-CertificatePassword` says "there's really no automated *and* secure way to connect using a local certificate." Use a certificate from the store or load it from Key Vault at run time.

## Microsoft Learn references

- [Connect-ExchangeOnline](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-exchangeonline)
- [Connect to Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell)
- [App-only authentication in Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2)
- [About the Exchange Online PowerShell module (versions)](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2)
- [Connect-IPPSSession](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-ippssession)
- [Create unifiedRoleAssignment](https://learn.microsoft.com/en-us/graph/api/rbacapplication-post-roleassignments?view=graph-rest-1.0)
- [GDAP and the Secure Application Model](https://learn.microsoft.com/en-us/partner-center/developer/gdap-and-secure-application-model)
