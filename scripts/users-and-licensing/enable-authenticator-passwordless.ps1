#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally enables, Microsoft Authenticator passwordless phone sign-in in the Authentication methods policy of one or more customer tenants.

.DESCRIPTION
    The original article turned on Authenticator phone sign-in with a preview
    Azure AD policy (New-AzureADPolicy -Type AuthenticatorAppSignInPolicy). That
    mechanism is gone. Authenticator is now configured in the Authentication
    methods policy, where each target group has an authenticationMode of push,
    deviceBasedPush (passwordless phone sign-in) or any (both).

    For each customer the script reads the Microsoft Authenticator configuration
    through Microsoft Graph and returns its state and targets.
    PasswordlessAllowed is true when the method is enabled and every include
    target allows passwordless (any or deviceBasedPush).

    With -Apply:
    - If Authenticator is disabled, it is enabled for all users with
      authenticationMode any.
    - If it is enabled, the existing include targets keep their groups and are
      switched from push to any. Targets that already allow passwordless are
      left as they are.
    Each change goes through ShouldProcess (supports -WhatIf) and is read back.
    Exclude targets and feature settings (number matching, app and location
    context) are not touched.

    Consider passkeys in Microsoft Authenticator (the Passkey (FIDO2) method) for
    phishing-resistant passwordless sign-in. The original article's advice to
    bypass MFA by adding your IP address to the legacy MFA trusted IPs list must
    not be followed. MspGdap tokens come from an MFA sign-in, so no bypass is
    needed.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER Apply
    Enables passwordless phone sign-in as described above. Without it the script
    only reports.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./enable-authenticator-passwordless.ps1 -AllCustomers | Format-Table CustomerName, State, Targets, PasswordlessAllowed

    Reports the Authenticator configuration in every active GDAP customer.

.EXAMPLE
    ./enable-authenticator-passwordless.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows the change for one customer. Remove -WhatIf to make it.

.NOTES
    Replaces the original 2018 method: AzureADPreview module (retired), Connect-AzureAD -Credential with a delegated admin account without MFA, Get-AzureADContract (DAP) and New-AzureADPolicy -Type AuthenticatorAppSignInPolicy, plus advice to bypass MFA with trusted IPs (removed as unsafe).
    Required GDAP roles: Global Reader to report, Authentication Policy Administrator for -Apply.
    Required partner app permissions: Microsoft Graph delegated Policy.Read.AuthenticationMethod to report and Policy.ReadWrite.AuthenticationMethod for -Apply (covered in manifests/partner-app.full.json by Policy.ReadWrite.AuthenticationMethod).

.LINK
    https://gcit.com.au/knowledge-base/allow-passwordless-authentication-for-all-delegated-office-365-tenants/

.LINK
    https://learn.microsoft.com/en-us/graph/api/microsoftauthenticatorauthenticationmethodconfiguration-update

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [switch]$Apply,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $policyUri = 'v1.0/policies/authenticationMethodsPolicy/authenticationMethodConfigurations/microsoftAuthenticator'
    $columns = @('CustomerTenantId', 'CustomerName', 'State', 'Targets', 'PasswordlessAllowed', 'Action', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()

    function Get-ResultRow {
        param(
            [Parameter(Mandatory)][string[]]$Column,
            [Parameter(Mandatory)][System.Collections.IDictionary]$Value
        )
        $row = [ordered]@{}
        foreach ($name in $Column) {
            $row[$name] = if ($Value.Contains($name)) { $Value[$name] } else { $null }
        }
        [pscustomobject]$row
    }

    function Get-PolicySummary {
        param($Policy)
        $includes = @($Policy.includeTargets | Where-Object { $_ })
        $allowed = $Policy.state -eq 'enabled' -and $includes.Count -gt 0 -and @($includes | Where-Object { $_.authenticationMode -notin 'any', 'deviceBasedPush' }).Count -eq 0
        [pscustomobject]@{
            State               = $Policy.state
            Targets             = ($includes | ForEach-Object { '{0}:{1}' -f $_.id, $_.authenticationMode }) -join ', '
            PasswordlessAllowed = $allowed
            Includes            = $includes
        }
    }
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
        foreach ($item in $TenantId) { $targets.Add($item) }
    }
}

end {
    if ($AllCustomers) {
        foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus | Where-Object { $_.GdapStatus -eq 'active' })) {
            $targets.Add($customer.TenantId)
            $knownNames[$customer.TenantId] = $customer.DisplayName
        }
    }

    foreach ($target in $targets) {
        $tenant = $target
        $values = [ordered]@{ CustomerTenantId = $target; Action = 'None'; Status = 'OK' }
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $values.CustomerTenantId = $tenant
            $values.CustomerName = $knownNames[$tenant]
            if (-not $values.CustomerName) {
                $values.CustomerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $summary = Get-PolicySummary -Policy (Invoke-MspGraphRequest -TenantId $tenant -Uri $policyUri)
            $values.State = $summary.State
            $values.Targets = $summary.Targets
            $values.PasswordlessAllowed = $summary.PasswordlessAllowed

            if ($summary.PasswordlessAllowed) {
                $values.Action = 'AlreadyAllowed'
            }
            else {
                $body = @{ '@odata.type' = '#microsoft.graph.microsoftAuthenticatorAuthenticationMethodConfiguration' }
                if ($summary.State -ne 'enabled' -or $summary.Includes.Count -eq 0) {
                    $body.state = 'enabled'
                    $body.includeTargets = @(@{ targetType = 'group'; id = 'all_users'; authenticationMode = 'any' })
                    $description = 'Enable Microsoft Authenticator for all users with passwordless phone sign-in (mode any)'
                }
                else {
                    $body.includeTargets = @($summary.Includes | ForEach-Object {
                            $mode = if ($_.authenticationMode -in 'any', 'deviceBasedPush') { $_.authenticationMode } else { 'any' }
                            $includeTarget = @{ targetType = $_.targetType; id = $_.id; authenticationMode = $mode }
                            # PATCH replaces the whole include list, so carry the registration flag over.
                            if ($_.PSObject.Properties['isRegistrationRequired'] -and $null -ne $_.isRegistrationRequired) { $includeTarget.isRegistrationRequired = [bool]$_.isRegistrationRequired }
                            $includeTarget
                        })
                    $description = 'Allow passwordless phone sign-in (mode any) for the existing Authenticator targets'
                }

                if (-not $Apply) {
                    $values.Action = 'WouldChange'
                }
                elseif ($PSCmdlet.ShouldProcess("$($values.CustomerName) ($tenant)", $description)) {
                    Invoke-MspGraphRequest -TenantId $tenant -Method PATCH -Uri $policyUri -Body $body -Confirm:$false | Out-Null
                    $after = Get-PolicySummary -Policy (Invoke-MspGraphRequest -TenantId $tenant -Uri $policyUri)
                    $values.State = $after.State
                    $values.Targets = $after.Targets
                    $values.PasswordlessAllowed = $after.PasswordlessAllowed
                    $values.Action = if ($after.PasswordlessAllowed) { 'Changed' } else { 'Unconfirmed' }
                    if (-not $after.PasswordlessAllowed) { $values.Status = 'Warning' }
                }
                else {
                    $values.Action = 'WhatIf'
                }
            }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $values.Status = 'Failed'
            $values.Error = $_.Exception.Message
        }

        $row = Get-ResultRow -Column $columns -Value $values
        $results.Add($row)
        $row
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
