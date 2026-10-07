# 01 Prerequisites

Before you create the partner app or register a technician, get the partner side right: Partner Center, GDAP relationships, security groups, admin accounts, MFA and Conditional Access. Most failures later in the process trace back to one of these.

## Checklist

- [ ] Partner Center account with admin access to Customers and Administer.
- [ ] An active GDAP relationship with each customer you will manage (status `active`).
- [ ] Security groups in the partner tenant, one per job function, assigned to each relationship with a least-privilege role set.
- [ ] A dedicated admin account per technician, in the right security groups.
- [ ] MFA registered and enforced for every admin account (phishing-resistant methods preferred).
- [ ] Conditional Access policies for admin accounts in the partner tenant.
- [ ] The technicians who will pre-consent the partner app are members of **AdminAgents**.
- [ ] PowerShell 7.6 (or 7.4 until 10 November 2026) and `Microsoft.PowerShell.SecretManagement` with a vault.

## Partner Center

MspGdap uses two Microsoft APIs on the partner side:

- **Microsoft Graph** (`tenantRelationships/delegatedAdminRelationships`) for GDAP relationships and access assignments. This needs the delegated permission `DelegatedAdminRelationship.ReadWrite.All` on the partner app (it is in both manifests).
- **Partner Center REST API** (`api.partnercenter.microsoft.com`) for customer lists and application consent. Microsoft requires the Secure Application Model for any App+User integration with Partner Center, which is what this toolkit implements.

Partner Center points to know:

- Calling the Partner Center API as a technician requires membership of the **AdminAgents** security group in the partner tenant. Microsoft states that the partner user "must also be a member of the AdminAgents security group, which is required for calling the Partner Center APIs." Keep AdminAgents small: only the people who pre-consent the app or manage relationships.
- Since **1 April 2026**, every App+User call to Partner Center must carry an MFA claim. A token without one is rejected with `401 Unauthorized - MFA required`. MspGdap sends the `ValidateMfa: true` header and reports `isMfaCompliant` in its results.
- GDAP does **not** need a reseller (CSP) relationship. A customer that buys licences elsewhere can still approve a GDAP request from Partner Center (Administer, then Request admin relationship).

## GDAP relationships

A GDAP relationship is time-bound, least-privilege access that the customer must approve. MspGdap can create relationships for you (`New-MspGdapRelationship`) or you can create them in Partner Center. Either way, these rules from Microsoft apply:

| Rule | Detail |
| --- | --- |
| Duration | Between 1 day (`P1D`) and 2 years (`P2Y`). MspGdap defaults to `P730D`. |
| Auto-extend | Only `PT0S` (off), `P0D` or `P180D` (extend by six months). MspGdap's `-AutoExtend` switch sets `P180D`. |
| Global Administrator | A relationship that contains Global Administrator **cannot** auto-extend. Microsoft also recommends against replacing Global Administrator with every other role. |
| Roles are fixed | You can't add Microsoft Entra roles to a relationship after it is created. To add roles, create a second relationship. |
| Approval | The partner locks a new relationship for approval (`lockForApproval`). The customer then approves it. Unapproved requests expire after 90 days. |
| Display name | Unique across all your relationships, 50 characters or fewer. |
| Security groups | Up to 100 security groups per customer. Access assignments are made one group at a time and start as `pending`. |

Preview a relationship before you create it. The role list in the `-WhatIf` output is final once the relationship exists:

```powershell
New-MspGdapRelationship -TenantId 'fabrikam.onmicrosoft.com' -DisplayName 'Contoso-Fabrikam-Std-2026' -AccessMapPath ./gdap-access.json -Duration 'P730D' -AutoExtend -WhatIf
```

`New-MspGdapRelationship` returns the relationship ID. Send the customer the approval link from Partner Center (Customers, select the customer, Admin relationships). MspGdap does not build the approval link itself because Microsoft does not document its format.

### Default GDAP relationships

Since 25 September 2023, new customers created through Partner Center receive a **Default GDAP** relationship instead of DAP. Microsoft's Default GDAP role set includes highly privileged roles such as Privileged Role Administrator and Privileged Authentication Administrator. List them and decide whether to keep, replace or terminate them:

```powershell
Get-MspGdapRelationship -Status active | Where-Object { $_.ContainsPrivilegedRoles } | Format-Table CustomerName, DisplayName, EndDateTime, AutoExtendDuration
```

## Security groups and the least-privilege role map

Create security groups in your **partner tenant** (cloud-only, role-assignable is not required) and add technicians to them. Then assign each group to each customer relationship with only the roles that job needs. Effective access in a customer is always the combination of:

1. the scopes the partner app holds in that customer (the manifest), and
2. the roles the technician's groups hold in that customer through GDAP.

The default map below is a starting point. It deliberately leaves out Global Administrator and Privileged Authentication Administrator.

| Partner group (example name) | Microsoft Entra roles | Template IDs | Typical jobs |
| --- | --- | --- | --- |
| `GDAP-Readers` | Global Reader, Reports Reader, Security Reader | `f2ef992c-3afb-46b9-b7cf-a126ee74c451`, `4a5d8f65-41da-4de4-8968-e035b65339cf`, `5d6b6bb7-de71-4623-b4af-96380a352509` | Reporting, audits, read-only scripts |
| `GDAP-Helpdesk` | Helpdesk Administrator, Service Support Administrator, Directory Readers | `729827e3-9c14-49f7-bb1b-9608f156bbb8`, `f023fd81-a637-4b56-95fd-791ac0226033`, `88d8e3e3-8f55-4a1e-953a-9b9898b8876b` | Password resets for non-admins, Microsoft support requests |
| `GDAP-UserAdmin` | User Administrator, Groups Administrator, License Administrator, Authentication Administrator | `fe930be7-5e62-47db-91af-98c3a49a38b1`, `fdd7a751-b60b-444a-984c-02652fe8fa1c`, `4d6ac14f-3453-41d0-bef9-a3e0c569773a`, `c4e39bd9-1100-46d3-8c65-fb160da0071f` | Joiners and leavers, licences, MFA resets for non-admins |
| `GDAP-Exchange` | Exchange Administrator | `29232cdf-9323-42fd-ade2-1d097af3e4de` | Mailboxes, `Connect-MspExchangeOnline` |
| `GDAP-Endpoint` | Intune Administrator | `3a2c62db-5318-420d-8d74-23affee5d9d5` | Devices, apps, compliance |
| `GDAP-Collaboration` | Teams Administrator, SharePoint Administrator | `69091246-20e8-4a56-aa4d-066075b2a7a8`, `f28a1f50-f6e7-4571-818b-6a12f2af6b6c` | Teams, SharePoint, OneDrive |
| `GDAP-Security` | Security Administrator, Conditional Access Administrator | `194ae4cb-b126-40b2-bd5b-6091b380977d`, `b1be1c3e-b65d-4f19-8427-f6fa0d97feb9` | Defender, Conditional Access changes |
| `GDAP-AppConsent` | Cloud Application Administrator | `158c047a-c907-4556-b7ef-446551a6b5f7` | `Grant-MspPartnerAppConsent` only |
| `GDAP-PrivilegedRoles` (optional, just in time) | Privileged Role Administrator | `e8611ab8-c189-46e8-94e1-60213ab1f814` | `Enable-MspExchangeAppAccess` only |

Notes on the map:

- **Consent.** Microsoft's current guidance is that the relationship and group used for application consent must include Application Administrator or Cloud Application Administrator. It states "Use of Privileged Role Administrator is no longer recommended" for consent. Microsoft also suggests a short consent-only relationship if you do not want Cloud Application Administrator in your standard relationship.
- **Privileged Role Administrator** is only needed to assign the Exchange Administrator role to the optional automation app's service principal. Keep it in its own group, make membership just in time with Microsoft Entra Privileged Identity Management for groups, or leave it out entirely if you don't use app-only Exchange.
- **Authentication Administrator** can only reset methods for non-admin users. Resetting MFA for customer admins needs Privileged Authentication Administrator, which is not in the default map on purpose. Handle those cases with the customer.
- Role template IDs are the same in every tenant. Check them against Microsoft's [built-in roles reference](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference).

### The access map file

`New-MspGdapRelationship` and `Set-MspGdapAccessAssignment` read the same JSON file, which you keep with your own configuration (not in this repository). Start from `src/MspGdap/Data/gdap-access.example.json`, which contains the full default map above. The relationship gets the union of all roles in the file, and each group gets only its own roles. Replace every `groupId` placeholder with the object ID of a security group in your partner tenant.

```json
{
  "accessMapVersion": 1,
  "assignments": [
    {
      "groupDisplayName": "GDAP-Readers",
      "groupId": "<PartnerSecurityGroupObjectId>",
      "roles": [
        { "displayName": "Global Reader", "roleTemplateId": "f2ef992c-3afb-46b9-b7cf-a126ee74c451" },
        { "displayName": "Reports Reader", "roleTemplateId": "4a5d8f65-41da-4de4-8968-e035b65339cf" },
        { "displayName": "Security Reader", "roleTemplateId": "5d6b6bb7-de71-4623-b4af-96380a352509" }
      ]
    },
    {
      "groupDisplayName": "GDAP-Exchange",
      "groupId": "<PartnerSecurityGroupObjectId>",
      "roles": [
        { "displayName": "Exchange Administrator", "roleTemplateId": "29232cdf-9323-42fd-ade2-1d097af3e4de" }
      ]
    }
  ]
}
```

Once the customer has approved the relationship, assign the groups and confirm:

```powershell
$relationship = Get-MspGdapRelationship -TenantId 'fabrikam.onmicrosoft.com' -Status active
Set-MspGdapAccessAssignment -RelationshipId $relationship.Id -AccessMapPath ./gdap-access.json -WhatIf
Set-MspGdapAccessAssignment -RelationshipId $relationship.Id -AccessMapPath ./gdap-access.json
Test-MspGdapAccess -TenantId 'fabrikam.onmicrosoft.com'
Test-MspGdapAccess -TenantId 'fabrikam.onmicrosoft.com' -AnyRole 'Cloud Application Administrator', 'Application Administrator'
```

`Set-MspGdapAccessAssignment` makes one call per group, reads each assignment back until it is `active` (or `-WaitSeconds` runs out) and reports each group as its own step. `Test-MspGdapAccess` works out the technician's effective roles from the partner side (active relationships, their access assignments and the technician's group memberships), then calls the customer with the technician's token to prove it reaches the right tenant. It does not ask the customer for your roles: a GDAP technician is not an object in the customer directory, so `/me/memberOf` fails there, and the `Customer-side role view` step is reported as `Skipped`. `-RequiredRole` and `-AnyRole` turn it into a pass or fail check.

## Dedicated admin accounts

Each technician needs their **own** admin account in the partner tenant for this toolkit. Do not use day-to-day accounts and never share an account between people.

- Cloud-only account (not synchronised from on-premises Active Directory), for example `jane.admin@contoso-msp.onmicrosoft.com`.
- No mailbox and no productivity licence unless there is a specific need. Less to phish, less to compromise.
- Member of the GDAP security groups for the technician's job, and of AdminAgents only if they pre-consent apps or manage relationships.
- MFA registered with phishing-resistant methods (passkeys, FIDO2 security keys or Windows Hello for Business) where possible.
- Disabled on the technician's last day, with sessions revoked. See [03 Register a technician token](03-register-technician-token.md#when-a-technician-leaves).

One account per person means one refresh token per person, so every action in a customer's audit log maps to a named technician.

## MFA

Microsoft requires MFA for partner users in Partner Center, the Partner Center API and partner delegated administration. The refresh token MspGdap stores must also come from an MFA sign-in: Microsoft states "the refresh token must be created through MFA to work for OBO scenarios."

`Register-MspPartnerToken` checks the `amr` claim of the first token it receives and refuses to store a token that does not show MFA, unless you override it (not recommended).

Microsoft's mandatory MFA for Azure (phase 2) covers Azure Resource Manager requests only. Microsoft Graph, Exchange Online and Partner Center calls are not in its scope, so do not rely on it to protect this toolkit. Enforce MFA yourself with Conditional Access.

## Conditional Access for admin accounts

Create these policies in the **partner tenant**, targeted at a group that contains every technician admin account (and excluding your break-glass accounts):

| Policy | Setting | Why |
| --- | --- | --- |
| Require phishing-resistant MFA | Grant: require authentication strength "Phishing-resistant MFA" for all resources | Stops replayed passwords and most adversary-in-the-middle phishing |
| Require a managed device | Grant: require a compliant or Microsoft Entra hybrid joined device | Keeps token registration and use on devices you control |
| Block device code flow | Conditions: authentication flows, device code flow, block | Microsoft recommends getting "as close as possible to a unilateral block on device code flow". MspGdap never uses it. |
| Block legacy authentication | Conditions: client apps, legacy clients, block | No password-only protocols |
| Restrict locations (optional) | Block outside your named locations or countries | Reduces exposure if credentials leak |

**Sign-in frequency trade-off.** A sign-in frequency session control applies when a refresh token is redeemed. If you set, for example, 12 hours for all resources, each technician will need to run `Register-MspPartnerToken` again after 12 hours, because the stored refresh token can no longer be redeemed silently. That is a valid, more secure choice for interactive use. If you want the token to stay usable for its full lifetime, scope the sign-in frequency policy so it does not target the partner app, and compensate with the device and authentication strength policies above. Decide this deliberately and record the decision.

Customers can also apply their own Conditional Access policies to your technicians as service provider users. If a customer's policy blocks a technician, the token request for that customer fails with an AADSTS error that names the policy outcome. That is the customer's control to change, not yours.

## Workstation

- PowerShell 7.6 LTS (`$PSVersionTable.PSVersion`).
- `Install-Module -Name Microsoft.PowerShell.SecretManagement -Repository PSGallery -Scope CurrentUser` plus a vault module. See [03 Register a technician token](03-register-technician-token.md).
- Optional: `ExchangeOnlineManagement` (3.10.x on PowerShell 7.6, 3.9.2 or earlier on 7.4 and 7.5), `Microsoft.Graph.Authentication`, `MicrosoftTeams`.
- A default browser that can reach `login.microsoftonline.com` and a free local TCP port for the loopback redirect (`http://localhost:<port>`).

## Microsoft Learn references

- [Introduction to GDAP](https://learn.microsoft.com/en-us/partner-center/customers/gdap-introduction)
- [GDAP frequently asked questions](https://learn.microsoft.com/en-us/partner-center/customers/gdap-faq)
- [Least privileged roles by task](https://learn.microsoft.com/en-us/partner-center/customers/gdap-least-privileged-roles-by-task)
- [Expiring GDAP relationships and auto extend](https://learn.microsoft.com/en-us/partner-center/customers/expiring-gdap-relationships-and-auto-extend-gdap)
- [GDAP and the Secure Application Model](https://learn.microsoft.com/en-us/partner-center/developer/gdap-and-secure-application-model)
- [Mandating MFA for partner tenants](https://learn.microsoft.com/en-us/partner-center/security/partner-security-requirements-mandating-mfa)
- [Block authentication flows with Conditional Access](https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-block-authentication-flows)
- [Create delegatedAdminRelationship (Graph)](https://learn.microsoft.com/en-us/graph/api/tenantrelationship-post-delegatedadminrelationships?view=graph-rest-1.0)
