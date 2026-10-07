# Permission manifests

These files describe the API permissions that `New-MspPartnerApp` puts on the multi-tenant partner app it creates in your partner tenant, and the permissions `Grant-MspPartnerAppConsent` pre-consents into each GDAP customer.

| File | Use it when |
| --- | --- |
| `partner-app.minimal.json` | Default. Least privilege for the MspGdap toolkit and typical admin and reporting scripts. Start here. |
| `partner-app.full.json` | You want parity with a mature MSP worker app (Intune, Teams, SharePoint, Defender response actions, policy writes and more). It is a strict superset of `partner-app.minimal.json`. Review every line before you use it. |
| `automation-app.example.json` | Optional. Application (app-only) permissions for a **separate** unattended automation app. Never merge it into the partner app. |

## How delegated permissions and GDAP combine

Every permission in the two partner app manifests is **delegated** (`"type": "Scope"`). A delegated token can only do what both of these allow:

1. the scopes consented for the app in the customer tenant, and
2. the Microsoft Entra roles the signed-in technician holds in that customer through GDAP security groups.

The manifest is a ceiling, not a grant of access. A technician whose GDAP groups only carry Helpdesk Administrator cannot reset a Global Administrator's password, even though the app holds `User.ReadWrite.All`. Keep GDAP role assignments tight and the manifest can stay practical.

## File format

```json
{
  "manifestVersion": 1,
  "name": "partner-app.minimal",
  "description": "...",
  "requiredResourceAccess": [
    {
      "resourceAppId": "00000003-0000-0000-c000-000000000000",
      "resourceDisplayName": "Microsoft Graph",
      "resourceAccess": [
        { "id": "e1fe6dd8-ba31-4d61-89e7-88639da4683d", "type": "Scope", "value": "User.Read" }
      ]
    }
  ]
}
```

- `resourceAppId` and `id` are Microsoft's public first-party application IDs and permission IDs. They are the same in every tenant.
- `resourceDisplayName` and `value` are documentation. Microsoft Graph rejects unknown properties, so `New-MspPartnerApp` removes them before it writes `requiredResourceAccess` to the application object.
- `New-MspPartnerApp` also checks every entry before writing: it reads the resource's service principal in your partner tenant (`oauth2PermissionScopes` for `Scope`, `appRoles` for `Role`) and stops if the `id` does not resolve to the stated `value`. A typo or a retired permission fails loudly instead of being written.
- There are no redirect URIs in the manifests. `New-MspPartnerApp` adds the loopback redirect (`http://localhost`) itself.

To build your own manifest, copy `partner-app.minimal.json`, add or remove entries, and look up IDs with a read-only Graph call such as `GET /servicePrincipals(appId='00000003-0000-0000-c000-000000000000')?$select=oauth2PermissionScopes`.

## Resources used

| Resource | Application ID |
| --- | --- |
| Microsoft Graph | `00000003-0000-0000-c000-000000000000` |
| Microsoft Partner Center | `fa3d9a0c-3fb0-42cc-9193-47c7ecd2edbd` |
| Office 365 Exchange Online | `00000002-0000-0ff1-ce00-000000000000` |
| Office 365 SharePoint Online | `00000003-0000-0ff1-ce00-000000000000` |
| Office 365 Management APIs | `c5393580-f805-4401-95e8-94b7a6ef2fc2` |
| WindowsDefenderATP (Microsoft Defender for Endpoint) | `fc780465-2017-40d4-a0c5-307022471b92` |
| Skype and Teams Tenant Admin API | `48ac35b8-9aa8-4d74-927d-1f4a14a0b239` |
| Azure Resource Manager | `797f4846-ba00-4fd7-ba43-dac1f8f63013` |
| Dataverse | `00000007-0000-0000-c000-000000000000` |
| Microsoft Forms | `c9a559d2-7aab-4f13-a6ed-e7e9c52aec87` |

## partner-app.minimal.json

31 delegated permissions across 5 resources.

| Resource | Permission | Purpose |
| --- | --- | --- |
| Graph | `openid` | Sign the technician in with OpenID Connect. |
| Graph | `profile` | Read the technician's basic profile claims at sign-in. |
| Graph | `offline_access` | Issue the refresh token that MspGdap stores in SecretManagement. |
| Graph | `User.Read` | Read the signed-in technician's own profile (UPN shown in token registration). |
| Graph | `User.ReadWrite.All` | Read and update user accounts (offboarding, licence and profile changes). |
| Graph | `Group.ReadWrite.All` | Read groups and manage membership (security groups, Microsoft 365 groups). |
| Graph | `Directory.Read.All` | Read directory objects, including service principals and consent grants checked by `Test-MspPartnerAppConsent`. |
| Graph | `Domain.Read.All` | Read verified domains, including the initial onmicrosoft.com domain used for Exchange connections. |
| Graph | `Organization.Read.All` | Read tenant details and subscribed SKUs. |
| Graph | `AuditLog.Read.All` | Read sign-in and directory audit logs. |
| Graph | `Reports.Read.All` | Read Microsoft 365 usage and MFA registration reports. |
| Graph | `SecurityEvents.Read.All` | Read Secure Score and legacy security events. |
| Graph | `SecurityAlert.Read.All` | Read Microsoft Defender XDR alerts. |
| Graph | `SecurityIncident.Read.All` | Read Microsoft Defender XDR incidents. |
| Graph | `Policy.Read.All` | Read Conditional Access, authentication method and other tenant policies. |
| Graph | `UserAuthenticationMethod.Read.All` | Read users' registered MFA methods for reporting. |
| Graph | `DeviceManagementManagedDevices.Read.All` | Read Intune managed devices. |
| Graph | `DeviceManagementConfiguration.Read.All` | Read Intune configuration and compliance policies. |
| Graph | `DelegatedAdminRelationship.ReadWrite.All` | Partner tenant only: read, create and assign GDAP relationships (`*-MspGdap*` functions). |
| Graph | `Application.Read.All` | Find the partner and automation app service principals in a customer. |
| Graph | `AppRoleAssignment.ReadWrite.All` | Grant `Exchange.ManageAsApp` to the automation app in a customer (`Enable-MspExchangeAppAccess`). |
| Graph | `RoleManagement.ReadWrite.Directory` | Add the automation app's service principal to Exchange Administrator in a customer (`Enable-MspExchangeAppAccess`). |
| Partner Center | `user_impersonation` | List customers and pre-consent the app through `applicationconsents` (`Grant-MspPartnerAppConsent`). |
| Exchange Online | `Exchange.Manage` | Delegated Exchange Online PowerShell (`Connect-MspExchangeOnline`, `Connect-MspSecurityCompliance`). |
| Office 365 Management APIs | `ActivityFeed.Read` | Read the unified audit log through the Management Activity API. See the note on Management API permission IDs below. |
| Defender for Endpoint | `Machine.Read` | Read onboarded devices. |
| Defender for Endpoint | `Alert.Read` | Read Defender for Endpoint alerts. |
| Defender for Endpoint | `Vulnerability.Read` | Read vulnerability (TVM) data. |
| Defender for Endpoint | `Software.Read` | Read software inventory. |
| Defender for Endpoint | `SecurityRecommendation.Read` | Read security recommendations. |
| Defender for Endpoint | `Score.Read` | Read exposure and secure configuration scores. |

Note on Office 365 Management API permission IDs: the manifests use the IDs that a working partner app holds as delegated scopes. Some public sources list the same IDs as application roles, with different IDs for the delegated scopes. `New-MspPartnerApp` checks every ID against the resource's published scopes in your tenant before it writes anything, so a wrong ID stops the run rather than adding the wrong permission. If it stops on these entries, look up the published IDs with `Invoke-MspGraphRequest -PartnerTenant -Uri "v1.0/servicePrincipals(appId='c5393580-f805-4401-95e8-94b7a6ef2fc2')?`$select=oauth2PermissionScopes,appRoles"` and correct your copy of the manifest.

The last three Graph entries are the most privileged in this file. They exist because `Enable-MspExchangeAppAccess` has to assign an app role and a directory role inside the customer. If you do not plan to use the optional automation app, remove `AppRoleAssignment.ReadWrite.All` and `RoleManagement.ReadWrite.Directory` before running `New-MspPartnerApp`.

## partner-app.full.json

122 delegated permissions across 10 resources. It is a strict superset of `partner-app.minimal.json`: every entry in the minimal manifest is also in this one, with the same resource, `id` and `type`, so moving from minimal to full never removes a permission. It also holds the delegated permissions of the worker app in the reference implementation, the four sign-in scopes (`openid`, `profile`, `offline_access`, `User.Read`) and four Graph permissions the bundled scripts use (`Domain.ReadWrite.All`, `ServiceMessage.Read.All`, `IdentityRiskEvent.Read.All`, `IdentityRiskyUser.Read.All`). Entries marked **High** can change security posture, run code on devices or act as users. Only keep them if your scripts need them.

`User.DeleteRestore.All` (permanently delete or restore users) is deliberately left out: a permanent delete cannot be undone, is rarely needed because Microsoft Entra purges deleted users after 30 days, and is safer done by hand in the Microsoft Entra admin center than held as a standing permission in every customer.

### Microsoft Graph (79)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `openid` | | Sign the technician in with OpenID Connect. |
| `profile` | | Read basic profile claims at sign-in. |
| `offline_access` | | Issue the refresh token stored in SecretManagement. |
| `User.Read` | | Read the signed-in technician's own profile. |
| `AppCatalog.ReadWrite.All` | | Publish and manage apps in the Teams app catalogue. |
| `AppCatalog.Submit` | | Submit Teams app packages for approval. |
| `Application.Read.All` | | Find the partner and automation app service principals in a customer. |
| `Application.ReadWrite.All` | **High** | Create and change app registrations and service principals. |
| `AppRoleAssignment.ReadWrite.All` | **High** | Grant app roles and delegated grants to service principals. |
| `AuditLog.Read.All` | | Read sign-in and directory audit logs. |
| `BitlockerKey.Read.All` | **High** | Read BitLocker recovery keys. |
| `Channel.ReadBasic.All` | | Read Teams channel names and descriptions. |
| `ChannelMessage.Read.All` | | Read Teams channel messages. |
| `Chat.Create` | | Create Teams chats. |
| `Chat.ReadWrite` | | Read and send Teams chat messages as the technician. |
| `CopilotPackages.Read.All` | | Read Microsoft 365 Copilot agent and app packages. |
| `CopilotPackages.ReadWrite.All` | | Manage Microsoft 365 Copilot agent and app packages. |
| `DelegatedAdminRelationship.ReadWrite.All` | | Partner tenant only: manage GDAP relationships and access assignments. |
| `DeviceManagementApps.ReadWrite.All` | | Manage Intune apps and app assignments. |
| `DeviceManagementConfiguration.Read.All` | | Read Intune configuration and compliance policies. |
| `DeviceManagementConfiguration.ReadWrite.All` | | Manage Intune configuration and compliance policies. |
| `DeviceManagementManagedDevices.Read.All` | | Read Intune managed devices. |
| `DeviceManagementManagedDevices.ReadWrite.All` | | Manage Intune devices (sync, rename, retire). |
| `DeviceManagementServiceConfig.ReadWrite.All` | | Manage Intune service settings, enrolment and Autopilot. |
| `Directory.Read.All` | | Read directory objects, including service principals and consent grants checked by `Test-MspPartnerAppConsent`. |
| `Directory.ReadWrite.All` | **High** | Read and write all directory objects. |
| `Domain.Read.All` | | Read verified domains, including the initial onmicrosoft.com domain used for Exchange connections. |
| `Domain.ReadWrite.All` | **High** | Change domain settings such as password expiry and the default domain. |
| `Files.ReadWrite.All` | | Read and write files the technician can reach in OneDrive and SharePoint. |
| `Group.ReadWrite.All` | | Read groups and manage membership (security groups, Microsoft 365 groups). |
| `IdentityRiskEvent.Read.All` | | Read Identity Protection risk detections. |
| `IdentityRiskyUser.Read.All` | | Read Identity Protection risky users. |
| `Mail.ReadWrite` | | Read and write mail the technician can access. |
| `Mail.Send` | | Send mail as the technician. |
| `MailboxSettings.ReadWrite` | | Change mailbox settings such as automatic replies. |
| `Organization.Read.All` | | Read tenant details and subscribed SKUs. |
| `Policy.Read.All` | | Read Conditional Access and other tenant policies. |
| `Policy.ReadWrite.AccessReview` | **High** | Change the access review default policy. |
| `Policy.ReadWrite.ApplicationConfiguration` | **High** | Change app management and token lifetime policies. |
| `Policy.ReadWrite.AuthenticationFlows` | **High** | Change authentication flow policies (self-service sign-up). |
| `Policy.ReadWrite.AuthenticationMethod` | **High** | Change authentication methods policy (MFA methods, passkeys, TAP). |
| `Policy.ReadWrite.Authorization` | **High** | Change the authorisation policy (user consent, guest restrictions). |
| `Policy.ReadWrite.ConditionalAccess` | **High** | Create and change Conditional Access policies and named locations. |
| `Policy.ReadWrite.ConsentRequest` | **High** | Change admin consent request settings. |
| `Policy.ReadWrite.CrossTenantAccess` | **High** | Change cross-tenant access settings. |
| `Policy.ReadWrite.CrossTenantCapability` | **High** | Change Microsoft 365 cross-tenant capability settings. |
| `Policy.ReadWrite.DeviceConfiguration` | **High** | Change device registration policy. |
| `Policy.ReadWrite.ExternalIdentities` | **High** | Change external identities (B2B) policy. |
| `Policy.ReadWrite.FeatureRollout` | **High** | Change staged rollout policies. |
| `Policy.ReadWrite.FedTokenValidation` | **High** | Change federated token validation policy. |
| `Policy.ReadWrite.IdentityProtection` | **High** | Change Identity Protection policies. |
| `Policy.ReadWrite.MobilityManagement` | **High** | Change MDM and MAM enrolment scope policies. |
| `Policy.ReadWrite.PermissionGrant` | **High** | Change consent and permission grant policies. |
| `Policy.ReadWrite.SecurityDefaults` | **High** | Turn security defaults on or off. |
| `Policy.ReadWrite.TrustFramework` | **High** | Change trust framework (custom policy) settings. |
| `Presence.Read.All` | | Read Teams presence for users. |
| `Reports.Read.All` | | Read Microsoft 365 usage and MFA registration reports. |
| `RoleManagement.ReadWrite.Directory` | **High** | Assign and remove Entra directory roles. |
| `SecurityAlert.Read.All` | | Read Microsoft Defender XDR alerts. |
| `SecurityAlert.ReadWrite.All` | | Read and update Defender XDR alerts. |
| `SecurityEvents.Read.All` | | Read Secure Score and legacy security events. |
| `SecurityEvents.ReadWrite.All` | | Read and update Secure Score and security events. |
| `SecurityIncident.Read.All` | | Read Microsoft Defender XDR incidents. |
| `ServiceMessage.Read.All` | | Read Microsoft 365 Message center posts. |
| `Sites.FullControl.All` | **High** | Full control of SharePoint sites the technician can reach through Graph. |
| `Sites.ReadWrite.All` | | Read and write SharePoint site content through Graph. |
| `Tasks.ReadWrite` | | Read and write the technician's Planner and To Do tasks. |
| `Team.ReadBasic.All` | | List teams. |
| `TeamsAppInstallation.ReadWriteAndConsentForChat` | | Install Teams apps in chats and consent to their permissions. |
| `TeamsAppInstallation.ReadWriteForChat` | | Install and remove Teams apps in chats. |
| `TeamsAppInstallation.ReadWriteForTeam` | | Install and remove Teams apps in teams. |
| `TeamsAppInstallation.ReadWriteForUser` | | Install and remove Teams apps for users. |
| `TeamsAppInstallation.ReadWriteSelfForChat` | | Let a Teams app manage its own installation in chats. |
| `TeamsAppInstallation.ReadWriteSelfForTeam` | | Let a Teams app manage its own installation in teams. |
| `ThreatHunting.Read.All` | | Run Defender XDR advanced hunting queries. |
| `ThreatIntelligence.Read.All` | | Read Defender threat intelligence. |
| `User.ReadWrite.All` | | Read and update user accounts. |
| `UserAuthenticationMethod.Read.All` | | Read users' registered MFA methods for reporting. |
| `UserAuthenticationMethod.ReadWrite.All` | **High** | Read, add and remove users' MFA methods (reset MFA, issue a TAP). |

### Microsoft Partner Center (1)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `user_impersonation` | | List customers and pre-consent the app through `applicationconsents`. |

### Office 365 Exchange Online (1)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `Exchange.Manage` | | Delegated Exchange Online and Security and Compliance PowerShell. |

### Office 365 SharePoint Online (7)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `AllSites.FullControl` | **High** | SharePoint admin and PnP operations across all sites. |
| `AllSites.Manage` | | Create and change lists and libraries in all sites. |
| `AllSites.Write` | | Read and write list items and documents in all sites. |
| `MyFiles.Read` | | Read the technician's own OneDrive files. |
| `MyFiles.Write` | | Write the technician's own OneDrive files. |
| `Sites.Search.All` | | Run SharePoint search queries as the technician. |
| `TermStore.ReadWrite.All` | | Read and write managed metadata. |

### Office 365 Management APIs (3)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `ActivityFeed.Read` | | Read the unified audit log through the Management Activity API. |
| `ActivityFeed.ReadDlp` | | Read DLP events, including detected sensitive data. |
| `ServiceHealth.Read` | | Read service health through the Management API. |

### WindowsDefenderATP, Microsoft Defender for Endpoint (21)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `Alert.Read` | | Read alerts. |
| `Alert.ReadWrite` | | Read and update alerts. |
| `File.Read.All` | | Read file profiles. |
| `IntegrationConfiguration.ReadWrite` | | Read and change integration settings. |
| `Ip.Read.All` | | Read IP address profiles. |
| `Machine.CollectForensics` | **High** | Collect an investigation package from a device. |
| `Machine.Isolate` | **High** | Isolate and release a device. |
| `Machine.LiveResponse` | **High** | Run live response sessions (remote shell) on a device. |
| `Machine.Read` | | Read onboarded devices. |
| `Machine.ReadWrite` | | Read devices and update tags and metadata. |
| `Machine.RestrictExecution` | **High** | Restrict app execution on a device. |
| `Machine.Scan` | | Start an antivirus scan. |
| `Machine.StopAndQuarantine` | **High** | Stop a process and quarantine a file. |
| `RemediationTasks.Read` | | Read TVM remediation tasks. |
| `Score.Read` | | Read exposure and secure configuration scores. |
| `SecurityBaselinesAssessment.Read` | | Read security baseline assessments. |
| `SecurityConfiguration.ReadWrite` | | Read and change security configuration. |
| `SecurityRecommendation.Read` | | Read security recommendations. |
| `Software.Read` | | Read software inventory. |
| `Url.Read.All` | | Read URL profiles. |
| `Vulnerability.Read` | | Read vulnerability data. |

### Skype and Teams Tenant Admin API (4)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `user_impersonation` | | Delegated Microsoft Teams PowerShell and admin operations. |
| `Team.ReadBasic.All` | | List the technician's teams through the Teams admin API. |
| `MailboxSettings.Read` | | Mailbox settings read exposed by the Teams admin API (Microsoft labels it for internal service use). |
| `sharepoint_service` | | Lets the Teams admin service connect to SharePoint on the technician's behalf. |

### Azure Resource Manager (1)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `user_impersonation` | | Manage Azure resources in customer subscriptions as the technician (needs Azure RBAC as well as GDAP). |

### Dataverse (1)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `user_impersonation` | | Access Dataverse and Power Platform environments as the technician. |

### Microsoft Forms (4)

| Permission | Impact | Purpose |
| --- | --- | --- |
| `Forms.Read` | | Read the technician's forms. |
| `Forms.ReadWrite` | | Create and change the technician's forms. |
| `Responses.Read.All` | | Read responses to the technician's forms. |
| `Responses.ReadWrite` | | Submit form responses. |

## Application permissions: use a separate app

The reference implementation also held application (app-only) permissions on the same app as its delegated permissions. MspGdap deliberately does not repeat that, and neither manifest above contains a `"type": "Role"` entry. Reasons:

- **No GDAP ceiling.** An app-only token is not tied to a technician, so GDAP roles do not limit it. A role such as `Directory.ReadWrite.All` applies to the whole customer tenant.
- **Bigger blast radius.** If the credential of a mixed app leaks, the attacker gets both every technician's delegated reach (with a stolen refresh token) and standing tenant-wide access in every customer where the app roles were granted.
- **Different consent path.** Partner Center `applicationconsents` pre-consents delegated scopes. App roles have to be granted per customer through Graph (`appRoleAssignedTo`), so they belong in their own, separately reviewed step.
- **Cleaner audit.** Sign-in logs show service principal sign-ins for the automation app and user sign-ins for the partner app. One app per purpose makes both easier to read and to alert on.

If you need unattended jobs (for example nightly Exchange reports), create a second multi-tenant app from `automation-app.example.json`:

| Resource | Permission (application) | Purpose |
| --- | --- | --- |
| Exchange Online | `Exchange.ManageAsApp` | App-only Exchange Online PowerShell with a certificate. The service principal also needs the Exchange Administrator role (or a narrower Exchange RBAC role) in each customer. |
| Graph | `Organization.Read.All` | Read tenant details and SKUs. |
| Graph | `Domain.Read.All` | Read domains, including the initial onmicrosoft.com domain. |
| Graph | `User.Read.All` | Read users. |
| Graph | `Group.Read.All` | Read groups and membership. |
| Graph | `AuditLog.Read.All` | Read sign-in and directory audit logs. |
| Graph | `Reports.Read.All` | Read usage reports. |
| Graph | `Policy.Read.All` | Read Conditional Access and other policies. |
| Office 365 Management APIs | `ActivityFeed.Read` | Read the unified audit log. |

Rules for the automation app:

- Certificate credential only. No client secrets.
- Keep it read-mostly. Add write roles one at a time, each with a written reason.
- Grant it per customer with `Enable-MspExchangeAppAccess`, then confirm with `Test-MspExchangeAppAccess`. Both return structured results and report a failure if any step fails.
- Remove it from customers you offboard.
