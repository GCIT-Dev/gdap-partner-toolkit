# 04 Pre-consent customers

Before a technician can get a token for a customer, the partner app must exist in that customer's tenant (as a service principal) with tenant-wide consent for its delegated scopes. Partner Center can do this for you through GDAP, without anyone in the customer clicking a consent prompt.

## What pre-consent does

`Grant-MspPartnerAppConsent` calls the Partner Center application consent API:

```http
POST https://api.partnercenter.microsoft.com/v1/customers/{customer_id}/applicationconsents
```

with a body like this, **one resource per request**, scopes joined with a comma and a space. That follows the [GDAP FAQ](https://learn.microsoft.com/en-us/partner-center/customers/gdap-faq): "Make individual calls for each resource. When making a single POST request, pass only one resource and its corresponding scopes" and "Concatenate the required scopes using a comma followed by a space." (The older developer page shows several grants in one array and scopes without spaces. MspGdap follows the newer FAQ.)

```json
{
  "applicationId": "<PartnerAppId>",
  "displayName": "Contoso MSP Worker",
  "applicationGrants": [
    {
      "enterpriseApplicationId": "00000003-0000-0000-c000-000000000000",
      "scope": "User.ReadWrite.All, Group.ReadWrite.All, Directory.Read.All"
    }
  ]
}
```

In the customer's tenant this creates the partner app's service principal (if missing) and one tenant-wide (`AllPrincipals`) delegated permission grant per resource. Nothing else is created. The consent gives no access by itself: every request still runs as a technician, inside that technician's GDAP roles.

## Requirements

| Requirement | Detail |
| --- | --- |
| Active GDAP relationship | Status `active` with the customer. GDAP-only customers (no reseller relationship) are supported. |
| Consent role in the customer | The technician must hold **Cloud Application Administrator** or **Application Administrator** in that customer through a GDAP security group. Microsoft says Privileged Role Administrator "is no longer recommended" for this. |
| AdminAgents | The technician must be a member of AdminAgents in the partner tenant to call Partner Center. |
| MFA | The technician's refresh token must come from an MFA sign-in. Partner Center rejects App+User calls without an MFA claim (`401 Unauthorized - MFA required`). |
| Same app | The Partner Center token must be issued to the same app that is being consented ("the appId claim in the access token should also be" the application ID in the body). MspGdap always uses the configured partner app, and refuses an `-AppId` that is not that app. |
| Resource service principals | Each resource in the manifest (for example Office 365 Management APIs or Microsoft Defender for Endpoint) must already exist as a service principal in the customer's tenant. A customer without a service may not have it. |
| Delegated only | Application (app-only) permissions can't be consented this way under GDAP. See [05 Exchange access](05-exchange-access.md) for the app-only path. |

Check the GDAP side first:

```powershell
Test-MspGdapAccess -TenantId 'fabrikam.onmicrosoft.com' -AnyRole 'Cloud Application Administrator', 'Application Administrator'
```

## Grant consent to one customer

```powershell
$consentParams = @{
    TenantId     = 'fabrikam.onmicrosoft.com'
    ManifestPath = './manifests/partner-app.minimal.json'
}
Grant-MspPartnerAppConsent @consentParams -WhatIf
Grant-MspPartnerAppConsent @consentParams
Test-MspPartnerAppConsent @consentParams
```

How `Grant-MspPartnerAppConsent` behaves:

1. It reads what already exists in the customer through Microsoft Graph: the partner app's service principal, its `AllPrincipals` grants and the resource service principals. A customer that has never been consented can't be read yet, see [First-time consent](#first-time-consent).
2. For each resource in the manifest it decides: **already granted** (skip), **missing** (POST), **scopes missing from an existing grant** (report, change only with `-Force`) or **resource not present in the tenant** (report and skip).
3. It sends one POST per missing resource, each under `-WhatIf` and `-Confirm`, with the `ValidateMfa: true` header and an `ms-requestid` idempotency key that stays the same if the call is retried.
4. It treats `200` and `201` as success. A `409` (seen in the field, not documented by Microsoft) is only treated as "already exists" after a Graph readback confirms the grant. A `500` is retried with back-off. A `401 MFA required` stops the run and tells you to register again with MFA.
5. It reads the result back through Graph and returns one result object per customer, with a step per resource. A step is only `Changed` when Graph confirms the new grant. If Partner Center accepted a request but Graph can't confirm it, the step is `Failed` or `Unknown`, never a pass. A `Failed` result is also written as a non-terminating error, so `-ErrorAction Stop` works.
6. With `-ManifestPath`, every scope is checked before anything is sent: the permission `id` must be a published delegated scope of the resource, the manifest's `value` must match the published name, and the scope must be in the partner app registration's `requiredResourceAccess`. A scope that is only in the manifest is refused unless you add `-Force`.

### First-time consent

The Graph reads run as the partner app in the customer. Before any consent exists, the app can't get a token there at all, and Microsoft Entra ID answers with `AADSTS65001` (the user or administrator has not consented to use the application). `Grant-MspPartnerAppConsent` treats that answer as "not consented yet", not as an error:

1. The `Read current consent` step reports `Passed` with "Not consented yet (AADSTS65001)".
2. Every resource in the manifest is posted to Partner Center, one POST per resource. The resources can't be checked for a service principal first, so a resource the customer doesn't have fails at Partner Center instead.
3. Once Partner Center has accepted the consent, the app can read the customer, and each grant is confirmed through Graph as usual before it is reported as `Changed`.

Consent takes a short while to reach Microsoft Graph, so the readback keeps trying for `-ReadbackTimeoutSeconds` (default 60). In a live test the default 60 second wait was enough. If a step still comes back `Unknown`, wait a minute and run `Test-MspPartnerAppConsent`.

Any other error on the first read (for example a `403`) still reports `Unknown` and posts nothing, because the current consent really is unknown. `Test-MspPartnerAppConsent` reports `AADSTS65001` as a `Failed` step, "Not consented".

### Adding scopes to an existing consent

Microsoft's documented way to add scopes (GDAP FAQ, "How can I add new scopes into the resource of an application that is already consented") is to DELETE the existing application consent and POST it again with the extra scopes, one POST per resource. Deleting first briefly removes the app from the customer, so `Grant-MspPartnerAppConsent` never does it silently. It reports the missing scopes as a failed step and only runs the DELETE and POST when you add `-Force` (and confirm, unless you also pass `-Confirm:$false`).

With `-Force` nothing that is consented now is lost. Before the DELETE it reads every tenant-wide grant the partner app holds in that customer, including resources you excluded with `-ExcludeResource`, resources that are not in your manifest, and extra scopes, and it posts all of them again together with the missing scopes. The confirmation message (and the `-WhatIf` output) lists exactly what will be re-posted. If it cannot identify an existing grant, it stops before the DELETE.

```powershell
Grant-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com' -ManifestPath './manifests/partner-app.minimal.json' -Force -WhatIf
```

## Grant consent to many customers

`Get-MspCustomer -IncludeGdapStatus` includes GDAP-only customers (without it, only customers with a reseller contract are listed). `Grant-MspPartnerAppConsent -AllCustomers` is a shortcut that targets every customer with an active GDAP relationship. Preview the whole run, then run it in small batches. A failed customer writes a non-terminating error as well as returning `Success = $false`, so keep the default error action in a loop like the one below. With `-ErrorAction Stop`, the first failed customer stops the whole run. Partner Center allows 50 requests per second per application, and MspGdap paces its calls well below that.

```powershell
$customers = Get-MspCustomer -IncludeGdapStatus
$customers | ForEach-Object {
    Grant-MspPartnerAppConsent -TenantId $_.TenantId -ManifestPath './manifests/partner-app.minimal.json' -WhatIf
}

$results = $customers | ForEach-Object {
    Grant-MspPartnerAppConsent -TenantId $_.TenantId -ManifestPath './manifests/partner-app.minimal.json'
}
$results | Where-Object { -not $_.Success } | ForEach-Object {
    $tenant = $_.TenantId
    $_.Steps | Where-Object { $_.Status -in 'Failed', 'Unknown' } | Select-Object @{ Name = 'TenantId'; Expression = { $tenant } }, Step, Status, Detail
} | Format-Table -Wrap
```

## Check consent

`Test-MspPartnerAppConsent` compares the customer's tenant with the manifest. It returns one result per customer with `TenantId`, `Success`, `Outcome` and a `Steps` list:

| Step | Passes when |
| --- | --- |
| Partner app service principal | The partner app's service principal exists in the customer |
| One step per resource | The resource's service principal exists and every manifest scope is in the tenant-wide (`AllPrincipals`) grant. A failure lists the missing scopes |
| Extra scopes (warning only) | Scopes are consented that the manifest doesn't list |
| Application permissions on partner app (warning only) | The partner app holds no app role assignments in the customer. It is meant to be delegated only |

```powershell
$check = Test-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com' -ManifestPath './manifests/partner-app.minimal.json'
$check.Success
$check.Steps | Format-Table Step, Status, Detail -Wrap
```

If a customer doesn't have a resource (for example Defender for Endpoint), exclude it for that customer with `-ExcludeResource` on both commands.

You can make the same check by hand with Graph:

```powershell
$tenant = 'fabrikam.onmicrosoft.com'
$sp = Invoke-MspGraphRequest -TenantId $tenant -Method GET -Uri "v1.0/servicePrincipals(appId='<PartnerAppId>')?`$select=id,displayName,appId"
Invoke-MspGraphRequest -TenantId $tenant -Method GET -Uri "v1.0/oauth2PermissionGrants?`$filter=clientId eq '$($sp.id)' and consentType eq 'AllPrincipals'" |
    Select-Object resourceId, scope
```

### In the customer's Microsoft Entra admin center

A technician with Global Reader (or the customer's own admin) can confirm it in the portal:

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) in the customer's directory.
2. Go to **Entra ID**, **Enterprise apps**, **All applications**.
3. Search for the partner app's name or application ID.
4. Open it and select **Permissions**. The **Admin consent** tab lists every API and scope granted.

Customers often ask what your app can do. Point them at this list and at [manifests/README.md](../manifests/README.md), and explain that the app can only act through your technicians' GDAP roles.

## Troubleshooting

| Error | Likely cause | Fix |
| --- | --- | --- |
| `AADSTS65001` (the user or administrator has not consented to use the application) | No consent in this customer, or no grant for the resource you asked for | Run `Grant-MspPartnerAppConsent`, then `Test-MspPartnerAppConsent`. Make sure the resource is in your manifest. |
| `AADSTS700016` (application not found in the directory) | The partner app's service principal does not exist in the customer | Run `Grant-MspPartnerAppConsent`. Since March 2026 an app can't authenticate in a tenant without its service principal. |
| `AADSTS500011` (resource principal not found in the tenant) | The resource (for example Defender for Endpoint) is not provisioned in the customer | Skip that resource for this customer, or ask the customer to enable the service. |
| `AADSTS53003` (blocked by Conditional Access) | A customer or partner Conditional Access policy blocked the sign-in | Check which tenant's policy applied in the sign-in logs. |
| `AADSTS700082` (refresh token expired due to inactivity) | No use for 90 days | Run `Register-MspPartnerToken` again. |
| Partner Center `401 MFA required` | Refresh token was not issued from an MFA sign-in | Run `Register-MspPartnerToken` again and complete MFA. |
| Partner Center `403 Forbidden` | Technician not in AdminAgents, or no Cloud Application Administrator or Application Administrator in this customer through GDAP | Fix group membership and the access assignment, then sign in again so the token carries the new roles. |
| Partner Center `404 Not Found` | Customer is not in your Partner Center customer list, or the customer ID differs from the tenant ID | Check that the customer appears in `Get-MspCustomer`. Microsoft describes the path value as the "ID of the customer generated in Partner Center", which is normally the tenant ID. |
| Partner Center `400 Bad Request` | A wrong `enterpriseApplicationId`, a scope name that doesn't exist, or a `displayName` or `applicationId` that doesn't match the app the token was issued to | Use the manifest. MspGdap sends one resource per request, validates scope names and always uses the configured app. |

Role and group changes in GDAP can take a while to apply. If you have just been added to a group, wait and request a fresh token (`Clear-MspTokenCache -TenantId <customer>`).

## Remove consent

When you offboard a customer, or to roll back a consent, remove the partner app's consent with `Remove-MspPartnerAppConsent`. It calls the Partner Center remove consent API

```http
DELETE https://api.partnercenter.microsoft.com/v1/customers/{customer_id}/applicationconsents/{application_id}
```

with the same MFA and role requirements as granting, then reads the customer through Graph until the partner app has no tenant-wide grants left:

```powershell
Remove-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com' -WhatIf
Remove-MspPartnerAppConsent -TenantId 'fabrikam.onmicrosoft.com'
```

The DELETE removes the tenant-wide grants but leaves the partner app's service principal in the customer tenant (seen in a live test on 7 October 2026). With no grants it can't be used, but for a full offboarding delete it too, from **Enterprise apps** in the customer tenant or with a Graph `DELETE servicePrincipals/{id}` as a technician whose GDAP roles include Cloud Application Administrator.

If Partner Center has no record of the consent (for example it was granted by a customer admin, not through Partner Center), the result says so and fails while grants remain. The customer's admin can then delete the partner app from **Enterprise apps** in their tenant. Then:

1. terminate or let expire the GDAP relationship,
2. remove the optional automation app from the customer if you enabled it,
3. confirm with `Test-MspPartnerAppConsent` that the app is gone (the partner app service principal step fails as not present).

## Microsoft Learn references

- [Control Panel Vendor APIs for customer consent](https://learn.microsoft.com/en-us/partner-center/developer/control-panel-vendor-apis)
- [GDAP and the Secure Application Model](https://learn.microsoft.com/en-us/partner-center/developer/gdap-and-secure-application-model)
- [GDAP frequently asked questions](https://learn.microsoft.com/en-us/partner-center/customers/gdap-faq)
- [Enable the Secure Application Model (ValidateMfa)](https://learn.microsoft.com/en-us/partner-center/developer/enable-secure-app-model)
- [Grant tenant-wide admin consent](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/grant-admin-consent)
