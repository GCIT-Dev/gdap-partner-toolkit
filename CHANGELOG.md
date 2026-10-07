# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.1] - Unreleased (pre-release)

Fixes from a live test on 7 October 2026 against a Microsoft partner tenant and a GDAP customer: app creation, technician token registration, pre-consent, Graph, Exchange Online and Security and Compliance with delegated GDAP tokens, and four rewritten scripts.

### Fixed

- `New-MspPartnerApp` and `Add-MspPartnerAppCertificate` did not recognise a certificate that was already on the app when Microsoft Graph returned `customKeyIdentifier` as the hex thumbprint. A 40 character hex thumbprint is also valid Base64, so the hex form is now compared first.
- `Grant-MspPartnerAppConsent` reported a never-consented customer as `Unknown` and posted nothing, because the partner app can't read a customer before consent exists (`AADSTS65001`). It now treats that error as "not consented yet", posts every resource and confirms each grant through Graph afterwards. Other read errors still report `Unknown` and post nothing.
- `Test-MspPartnerAppConsent` reports `AADSTS65001` as a `Failed` step, "Not consented", instead of `Unknown`.
- `Connect-MspExchangeOnline` failed with a validation error when it could not read the customer's initial domain, because the tenant ID fallback was assigned to the validated `-Organization` parameter. It now passes the tenant ID to `-DelegatedOrganization` as intended.
- `Invoke-MspGraphRequest -NoPaging` returned the raw page object (`@odata.context`, `@odata.nextLink`, `value`) for collections. It now writes the items of the first page, the same as paged mode, and writes a verbose message when more pages exist. Responses that are not collections (a single entity, report CSV content, `$count` text) are returned unchanged.
- `Test-MspGdapAccess` no longer calls `/me/memberOf` in the customer. GDAP technicians are not objects in the customer directory, so that call always failed with HTTP 400. The `Customer-side role view` step is reported as `Skipped`, and the roles come from the partner-side group assignments shown in `Effective roles`.

### Changed

- `Connect-MspSecurityCompliance` is no longer marked experimental. Its default (a delegated token for `https://ps.compliance.protection.outlook.com` with `-Organization`) was confirmed working in the live test, with no extra consent or manifest entry. It no longer writes a warning, and its result has `Experimental = $false`.
- The scripts that use `Connect-MspSecurityCompliance` (`connect-security-compliance.ps1`, `set-default-retention-policy.ps1`, `elevation-of-privilege-alert-policy.ps1`, `remove-email-from-mailboxes.ps1`), `scripts/README.md`, `scripts/MAPPING.json` and guide 05 no longer call that connection unverified.
- `Domain.ReadWrite.All`, `ServiceMessage.Read.All`, `IdentityRiskEvent.Read.All` and `IdentityRiskyUser.Read.All` are in `partner-app.full.json`, so script help, `scripts/README.md` and `scripts/MAPPING.json` no longer list them as missing. If your partner app was created from an older copy of the manifest, add them and run `Grant-MspPartnerAppConsent -Force` in each customer.
- `User.DeleteRestore.All` is removed from `partner-app.full.json` (now 122 delegated permissions), and `scripts/users-and-licensing/remove-deleted-user.ps1` no longer permanently deletes users. It lists soft-deleted users with the deletion date, the purge date and the days left until Microsoft Entra purges them (30 days), and returns guidance to permanently delete in the Microsoft Entra admin center when that is really required. A permanent delete cannot be undone and is rarely needed, so the toolkit does not hold a standing permission for it. Its `-Apply` switch is gone. If your partner app has `User.DeleteRestore.All` from an earlier copy of the full manifest, remove it from the app.
- Guide 03 explains that the first `Register-SecretVault` or `Set-SecretStoreConfiguration` for a new SecretStore prompts for a password, so it must run in an interactive PowerShell window.
- Guide 04 explains first-time consent and the readback wait (the default 60 seconds was enough in testing).
- `README.md` has a Tested section.
## [0.2.0] - Unreleased (pre-release)

KB script rewrites. The MspGdap module itself is unchanged.

### Added

- `scripts/`: 69 working GDAP rewrites of the PowerShell scripts GCIT published in its knowledge base and blog between 2017 and 2020, which relied on MSOnline, DAP, Exchange basic authentication remote PowerShell, the AzureAD and AzureRM modules, Azure AD Graph, Azure Functions v1 and stored passwords or AES key files. They cover 74 original articles in five folders: `exchange`, `users-and-licensing`, `delegated-reporting`, `automation` and `security-and-graph-apps`.
- Every script authenticates only through MspGdap (`Invoke-MspGraphRequest`, `Connect-MspExchangeOnline`, `Connect-MspSecurityCompliance`), takes `-TenantId` (or `-AllCustomers`), records per-customer failures and carries on, returns rows with `CustomerTenantId` and `CustomerName`, and can write a CSV with `-OutputPath`. Scripts that change anything are report only unless you add `-Apply`, and every change goes through `ShouldProcess`.
- Unsafe originals are replaced, not copied: temporary Global Administrator and Exchange admin accounts with shared passwords, MFA trusted IP bypasses and IP allow lists are replaced by the technician's GDAP role, app-only Exchange access or a report-only audit.
- Azure Functions v4 timer functions (PowerShell 7.6, managed identity and Key Vault, report only unless `MSPGDAP_APPLY_CHANGES` is `true`) in `scripts/automation/functions` and `scripts/security-and-graph-apps/functions`, replacing the Azure Functions v1 jobs.
- `scripts/MAPPING.json`, one entry per original article (83 in all) with its status, the method it replaced, the script or the reason there is none, and any permission the full manifest does not cover. `scripts/README.md` presents the same information with prerequisites and how to run the scripts.
- `tests/Scripts.Tests.ps1`: static checks that every script parses, requires PowerShell 7.4 and MspGdap, has complete help with links, exposes `-TenantId`, guards every change with `ShouldProcess` and uses none of the retired methods, and that `MAPPING.json` and `scripts/README.md` match the scripts on disk.
- `tests/scripts`: mocked smoke tests that run each script against `contoso.onmicrosoft.com`, including `-WhatIf` runs that must not write.

### Known gaps

- Five scripts needed a delegated Microsoft Graph permission that `partner-app.full.json` did not include at first (`Domain.ReadWrite.All`, `User.DeleteRestore.All`, `ServiceMessage.Read.All`, `IdentityRiskEvent.Read.All`, `IdentityRiskyUser.Read.All`). Four are now in the full manifest, and `remove-deleted-user.ps1` no longer needs `User.DeleteRestore.All` (see 0.2.1).
- The four scripts that use `Connect-MspSecurityCompliance` needed a live test. It passed on 7 October 2026 (see 0.2.1).

## [0.1.0] - Unreleased (pre-release)

First public pre-release.

### Added

- `MspGdap` module for PowerShell 7.4 or later (7.6 recommended, 7.4 and 7.5 supported until 10 November 2026). Windows PowerShell 5.1 is not supported.
- Technician token registration with the authorisation code flow, PKCE (`S256`) and a loopback redirect. Refresh tokens are stored only through SecretManagement and rewritten on every rotation (`Register-MspPartnerToken`).
- Certificate client assertions signed with PS256 and the `x5t#S256` header, with an RS256 fallback. Client secrets supported as `SecureString` but discouraged.
- Module-scoped token cache keyed by tenant, resource and app, with a five-minute expiry buffer and validation that treats access tokens as opaque (`Get-MspAccessToken`, `Get-MspAuthHeader`, `Test-MspAccessToken`, `Clear-MspTokenCache`, `Disconnect-Msp`).
- Configuration commands (`Set-MspConfiguration`, `Get-MspConfiguration`).
- Microsoft Graph requests with paging, `Retry-After` handling and `-WhatIf` on writes (`Invoke-MspGraphRequest`).
- Customer discovery and tenant resolution (`Get-MspCustomer`, `Resolve-MspTenantId`).
- Partner app creation from a validated permission manifest, certificate upload and partner tenant admin consent (`New-MspPartnerApp`, `Add-MspPartnerAppCertificate`).
- Partner Center application consent with one resource per request, the `ValidateMfa` header, an `ms-requestid` idempotency key and Graph readback (`Grant-MspPartnerAppConsent`, `Test-MspPartnerAppConsent`). Adding scopes with `-Force` re-posts everything already consented, so nothing is dropped.
- Consent removal with Graph readback (`Remove-MspPartnerAppConsent`) and technician offboarding (`Unregister-MspPartnerToken`).
- Write commands raise a non-terminating error when their outcome is `Failed`, as well as returning `Success = $false`.
- Customer-only commands refuse the partner tenant. `Disconnect-Msp` closes the Exchange, Security and Compliance, Graph SDK and Teams sessions that `Connect-Msp*` opened.
- Session-only certificate for automation hosts (`Set-MspConfiguration -Certificate`).
- GDAP relationship and access assignment management with a least-privilege default role map that excludes Global Administrator (`Get-MspGdapRelationship`, `New-MspGdapRelationship`, `Set-MspGdapAccessAssignment`, `Test-MspGdapAccess`).
- App-only Exchange enablement for a separate automation app using the unified role management API, with per-step readback, a check that the app belongs to your partner tenant, and roles limited to those Exchange app-only supports (`Enable-MspExchangeAppAccess`, `Test-MspExchangeAppAccess`).
- Connections for Exchange Online, Microsoft Graph PowerShell and Microsoft Teams, plus experimental Security and Compliance (`Connect-MspExchangeOnline`, `Connect-MspGraph`, `Connect-MspTeams`, `Connect-MspSecurityCompliance`).
- Permission manifests: `partner-app.minimal.json`, `partner-app.full.json` and `automation-app.example.json`.
- Guides 01 to 08, threat model, security policy, contributing guide.
- CI with PSScriptAnalyzer (module and tests), Pester on Windows and Linux (including a check that every documentation example matches the module), and gitleaks with custom rules.

