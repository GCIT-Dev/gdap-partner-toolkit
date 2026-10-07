# Security policy

MspGdap handles credentials that can reach many customer tenants. We take reports seriously and would rather hear about a possible problem early than read about it later.

## Reporting a vulnerability

**Do not open a public issue for a security problem.**

Report it privately through GitHub: open the repository's **Security** tab and select **Report a vulnerability**. Include:

- what you found and where (file, function, line or commit),
- how to reproduce it, ideally offline or against a test tenant,
- the impact you expect (for example token disclosure, wrong-tenant change, privilege escalation), and
- whether you have shared it with anyone else.

Never include real tokens, secrets, tenant IDs or customer data in a report. Redact them or use placeholders such as `<PartnerTenantId>`.

What to expect:

- An acknowledgement within 5 business days (Australian Eastern time).
- An initial assessment within 10 business days.
- A fix, a mitigation or a clear explanation of why we don't consider it a vulnerability. We will agree a disclosure date with you and credit you in the release notes unless you prefer otherwise.

This is a community project with no paid support. The timings above are goals, not a service level agreement.

## Supported versions

| Version | Supported |
| --- | --- |
| Latest release on `main` | Yes |
| Older releases | No. Upgrade to the latest release |

## Scope

In scope:

- The `MspGdap` module in `src/`, the scripts in this repository, the manifests, the CI configuration and the documentation (for example guidance that would lead partners into an unsafe configuration).

Out of scope:

- Vulnerabilities in Microsoft services, Microsoft PowerShell modules or SecretManagement vault extensions. Report those to Microsoft through the [Microsoft Security Response Center](https://msrc.microsoft.com/).
- Issues that need an already compromised technician workstation with an unlocked vault, unless MspGdap makes the impact meaningfully worse.
- Findings that depend on a partner ignoring the documented prerequisites (for example storing tokens outside a vault).

## What MspGdap never does

The module is designed so that each of these is a bug if it ever happens:

1. **Never stores refresh tokens outside SecretManagement.** No plain-text files, no registry values, no environment variables, no `$global:` variables.
2. **Never logs secrets.** Access tokens, refresh tokens, client assertions, authorisation codes, PKCE verifiers, client secrets and private keys never appear in output, verbose, debug, information or warning streams, or error messages.
3. **Never returns a refresh token** from a public command. Access tokens are returned as `SecureString` by default. Plain text is only returned when you ask for it: `Get-MspAccessToken -AsPlainText`, and the bearer header from `Get-MspAuthHeader`, which exists because HTTP needs it. Internally the `Connect-Msp*` commands convert a token to plain text only to hand it to the Microsoft module that needs it, and drop the reference straight after.
4. **Never sends a token** to any host other than `login.microsoftonline.com` (refresh tokens) or the resource it was issued for (access tokens).
5. **Never defaults to the partner tenant.** Tenant-scoped commands need `-TenantId`, and the partner tenant needs an explicit `-PartnerTenant`. Commands that only act on customers (consent, Exchange, GDAP, `Connect-Msp*`) refuse the partner tenant outright.
6. **Never exports a private key** or writes certificates to disk.
7. **Never adds credentials** to the partner app's or automation app's service principal in a customer tenant.
8. **Never reports success for a failed step.** Write and test commands read changes back and return per-step status, and write commands also raise a non-terminating error when the outcome is `Failed`.
9. **Never adds Global Administrator** to a GDAP role map by default.
10. **Never uses device code flow or resource owner password credentials.**
11. **Never makes a change without `ShouldProcess`.** Every write supports `-WhatIf` and `-Confirm`, and `-WhatIf` changes nothing, not even session state.
12. **Never stores a technician token without proof of the sign-in.** The id_token nonce must match the request, and MFA must be visible in the token claims unless `-SkipMfaCheck` is given. The sign-in listener ignores requests that do not come from this computer.
13. **Never gives an automation app more than Exchange needs by default.** `Enable-MspExchangeAppAccess` only assigns roles that Exchange Online app-only supports, needs a switch for roles beyond Exchange, and refuses apps that are not registered in your partner tenant.
14. **Never repeats a write blindly.** POST and PATCH requests are not retried on 503 or 504, and Partner Center calls reuse one `ms-requestid` across retries.

## Secret hygiene for contributors and users

- Use placeholders in code, tests, issues and pull requests: `<PartnerTenantId>`, `<PartnerAppId>`, `<CustomerTenantId>`, `contoso`, `fabrikam`.
- Test with mocked HTTP responses. Test JWTs must be clearly fake (for example signed with a throwaway key generated in the test and with `00000000-0000-0000-0000-000000000000` style IDs).
- Never commit `.pfx`, `.p12`, `.cer`, `.pem` or `.key` files, `config.json`, `*.secret` files or anything under `.mspgdap/`. The `.gitignore` blocks these, and CI runs gitleaks with the rules in `.gitleaks.toml` on every push and pull request.
- If you accidentally commit a secret, **rotate it first**, then remove it from history. Removing a commit does not un-leak a credential.
- When you share logs for troubleshooting, share only the AADSTS code, the trace ID and the correlation ID.

## If you think a token or certificate has leaked

1. Revoke the affected technician's sign-in sessions in your partner tenant and delete their refresh token secret ([docs/03](docs/03-register-technician-token.md#revocation)).
2. Remove the affected certificate from the partner app (or automation app) and issue a new one.
3. Review sign-in logs for the app and the account in the partner tenant, and audit logs in affected customers.
4. Tell affected customers in line with your contracts and any notification obligations that apply to you.
