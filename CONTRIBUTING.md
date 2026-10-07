# Contributing to MspGdap

Thanks for helping. This toolkit is used against real customer tenants, so the bar for correctness and security is high. Please read this page before you open a pull request.

## Ground rules

- **No secrets, ever.** No tokens, client secrets, certificates, real tenant IDs, real app IDs, real customer names or domains, or internal hostnames. Use placeholders such as `<PartnerTenantId>`, `contoso` and `fabrikam`. CI runs gitleaks on every push and will fail the build.
- **No live tenant calls in tests.** Every test mocks HTTP. Nothing in `tests/` may reach `login.microsoftonline.com`, `graph.microsoft.com`, `api.partnercenter.microsoft.com` or any other real endpoint.
- **Microsoft Learn is the source of truth.** When you add or change behaviour that depends on a Microsoft API, link the Learn page in the pull request and, where useful, in the comment-based help.
- **Least privilege by default.** New permissions in a manifest, new roles in the default GDAP map and new write operations need a written reason in the pull request.
- Be kind in reviews and issues.

## Development set-up

1. Install PowerShell 7.6 (7.4 also works until 10 November 2026).
2. Install the development modules:

   ```powershell
   Install-Module -Name Pester -MinimumVersion 5.5.0 -Repository PSGallery -Scope CurrentUser -Force -SkipPublisherCheck
   Install-Module -Name PSScriptAnalyzer -Repository PSGallery -Scope CurrentUser -Force
   ```

3. Run the same checks CI runs:

   ```powershell
   Invoke-ScriptAnalyzer -Path ./src -Recurse -Settings ./PSScriptAnalyzerSettings.psd1
   Invoke-ScriptAnalyzer -Path ./tests -Recurse -Settings ./tests/PSScriptAnalyzerSettings.psd1
   Invoke-Pester -Path ./tests -Output Detailed
   ```

4. Optionally install [gitleaks](https://github.com/gitleaks/gitleaks) and scan before you push:

   ```text
   gitleaks git --config .gitleaks.toml --verbose
   ```

## Code conventions

- One public function per file in `src/MspGdap/Public/`, one private helper per file in `src/MspGdap/Private/`. The file name matches the function name.
- Approved verbs only (`Get-Verb`). Public functions use the `Msp` noun prefix.
- `[CmdletBinding()]` on every function. `SupportsShouldProcess` on every function that changes anything, with `ConfirmImpact` set to `Medium` or `High` for destructive changes.
- Every tenant-scoped public function has a **mandatory** `-TenantId` (GUID or verified domain). Only an explicit `-PartnerTenant` switch may target the partner tenant.
- Return objects, not formatted text. Write and test commands return one result per target with a per-step `Status`, and `Success` is false if any step is `Failed`, `Unknown` or `WhatIf`. Build them with `New-MspOperationResult`. Write commands pass `-Cmdlet $PSCmdlet` so a `Failed` outcome is also written as a non-terminating error.
- Commands that only act on customers resolve the tenant with `Resolve-MspCustomerTenant`, which refuses the partner tenant. Graph and Partner Center calls go through the existing wrappers, which pin the host and handle retries. Do not retry POST or PATCH on 503 or 504.
- Comment-based help on every public function, with a `.PARAMETER` description for every parameter and at least one example that uses placeholders. `tests/Core.Hygiene.Tests.ps1` fails otherwise.
- PowerShell 7.4 or later only. Windows PowerShell 5.1 is not supported.
- No `Write-Host` in module code. Use `Write-Verbose`, `Write-Warning`, `Write-Error` and `Write-Information`, and never pass a token, secret or private key to any of them.
- Module state lives in `$script:` scope only. Never `$global:`.
- Australian English in documentation and help text. No em dashes and no semicolons in prose.
- GUIDs in `src/`, `scripts/` and `examples/` must be placeholders made of one repeated character (for example `00000000-0000-0000-0000-000000000000`) or public Microsoft IDs. Put a public ID on a line that names it (`roleTemplateId`, `resourceAppId`, `appRoleId` and similar) or end the line with `# public-id`. The `mspgdap-guid-in-code` gitleaks rule fails the build otherwise.
- Tests use the same repeated-character placeholder GUIDs for tenant and app IDs.

## Tests

- Pester 5. Mock `Invoke-RestMethod`, `Invoke-WebRequest`, the SecretManagement cmdlets and the Microsoft module cmdlets you call.
- Cover the failure paths: throttling (`429` with `Retry-After`), `401 MFA required`, expired tokens, opaque (non-JWT) tokens, wrong tenant, wrong audience, partial failure in multi-step operations.
- Test JWTs are built inside the test with fake claims. Do not paste tokens from anywhere, and build token-shaped strings at run time (for example `('e30', 'e30', 'x') -join '.'`) so secret scanners do not flag them.
- Every PowerShell example in `README.md` and `docs/` is parsed by `tests/Docs.Tests.ps1`, which fails when an example calls an MspGdap command or parameter that does not exist. Update the docs in the same change as the code.

## Pull requests

- Keep each pull request focused on one change.
- Update the documentation and `CHANGELOG.md` (under **Unreleased**) in the same pull request.
- Fill in the pull request description: what changed, why, the Microsoft Learn links you relied on and how you tested it.
- CI must pass: PSScriptAnalyzer (no errors), Pester (all tests pass on Windows and Linux), gitleaks (no findings).
- A maintainer reviews every change. Changes to authentication, token handling, secret storage, manifests or GDAP role maps need a second reviewer.

## GitHub organisation repositories and gitleaks

The CI workflow downloads a pinned gitleaks release, checks it against the release checksums and scans the full history with `.gitleaks.toml`. No licence key or secret is needed. Run the same scan locally with `gitleaks git --config .gitleaks.toml --redact .` before you push.

## Reporting security issues

Don't open a public issue. Follow [SECURITY.md](SECURITY.md).

## Licence

By contributing, you agree that your contributions are licensed under the [MIT licence](LICENSE).
