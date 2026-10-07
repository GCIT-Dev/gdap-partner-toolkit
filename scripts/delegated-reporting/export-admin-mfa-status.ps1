#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports every customer's administrators with their MFA registration and how MFA is enforced, and optionally creates a report-only Conditional Access policy that requires MFA for administrators.

.DESCRIPTION
    The original article switched on legacy per-user MFA for unlicensed Global
    Administrators, and included scripts to switch it off again. Per-user MFA is
    replaced by Conditional Access and security defaults, and Microsoft now
    enforces MFA for admin portals itself. What matters today is whether each
    admin has registered strong methods and whether the tenant enforces MFA for
    admin roles.

    Report (default): for each customer the script reads, through Microsoft Graph,
    - the authentication method registration of every user flagged as an admin
      (reports/authenticationMethods/userRegistrationDetails, enabled users only),
    - whether security defaults are on,
    - the enabled Conditional Access policies that require MFA (or an
      authentication strength) for all users or for directory roles.
    It returns one row per admin with IsMfaRegistered, the registered methods and
    AdminMfaEnforcedBy (security defaults, the policy names, or nothing).

    Create (-CreateConditionalAccessPolicy -Apply): in customers where nothing
    enforces MFA for admins, it creates a Conditional Access policy named
    "Require MFA for administrators" for Microsoft's recommended list of admin
    roles, in report-only state (enabledForReportingButNotEnforced). Exclude your
    break-glass accounts with -ExcludeUserId. Review the sign-in logs, then switch
    the policy on in the Microsoft Entra admin center. Each creation goes through
    ShouldProcess and supports -WhatIf. Conditional Access needs Microsoft Entra ID
    P1 (included in Microsoft 365 Business Premium). Tenants without it should use
    security defaults.

    If Conditional Access or the registration report cannot be read (for example
    in a tenant without Microsoft Entra ID P1), the customer is still reported,
    with the reason in the Error column, and no policy is created there.

    The original article also blocked the admins it listed until MFA could be
    registered. To block reviewed admin accounts, use
    export-global-admins.ps1 -BlockFromCsv, which never blocks the last enabled
    Global Administrator.

    This script never turns MFA off and never relies on shared admin accounts.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER CreateConditionalAccessPolicy
    Plans a report-only Conditional Access policy in customers where nothing
    enforces MFA for admins. Needs -Apply to create it.

.PARAMETER ExcludeUserId
    Object IDs of break-glass accounts to exclude from the new policy. Each one
    applies to every tenant in the run, so use it with a single -TenantId.

.PARAMETER Apply
    Creates the policy planned by -CreateConditionalAccessPolicy. Without it the
    script only reports.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-admin-mfa-status.ps1 -AllCustomers -OutputPath ./admin-mfa-status.csv

    Reports admin MFA registration and enforcement in every active GDAP customer.

.EXAMPLE
    ./export-admin-mfa-status.ps1 -TenantId 'contoso.onmicrosoft.com' -CreateConditionalAccessPolicy -ExcludeUserId '00000000-0000-0000-0000-000000000001' -Apply -WhatIf

    Shows the report-only policy that would be created for one customer. Remove -WhatIf to create it.

.NOTES
    Replaces the original 2018 method: Connect-MsolService, Get-MsolPartnerContract (DAP), Get-MsolRole 'Company Administrator', Get-MsolRoleMember and Set-MsolUser -StrongAuthenticationRequirements (legacy per-user MFA, MSOnline module retired 30 May 2025), plus scripts that removed MFA from admins (removed as unsafe).
    Required GDAP roles: Global Reader (or Reports Reader and Security Reader) to report, Conditional Access Administrator for -CreateConditionalAccessPolicy -Apply.
    Required partner app permissions: Microsoft Graph delegated AuditLog.Read.All and Policy.Read.All to report, Policy.Read.All and Policy.ReadWrite.ConditionalAccess to create the policy (all in manifests/partner-app.full.json).

.LINK
    https://gcit.com.au/knowledge-base/enable-mfa-on-all-global-admins-in-customers-office-365-tenants/

.LINK
    https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-old-require-mfa-admin

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [switch]$CreateConditionalAccessPolicy,

    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string[]]$ExcludeUserId,

    [switch]$Apply,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    if ($Apply -and -not $CreateConditionalAccessPolicy) {
        Write-Warning -Message '-Apply only applies to -CreateConditionalAccessPolicy. Reporting only.'
    }

    # Role template IDs from Microsoft's "Require MFA for administrators" Conditional Access template.
    $adminRoleIds = @(
        '62e90394-69f5-4237-9190-012177145e10' # public-id Global Administrator
        '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' # public-id Application Administrator
        'c4e39bd9-1100-46d3-8c65-fb160da0071f' # public-id Authentication Administrator
        'b0f54661-2d74-4c50-afa3-1ec803f12efe' # public-id Billing Administrator
        '158c047a-c907-4556-b7ef-446551a6b5f7' # public-id Cloud Application Administrator
        'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' # public-id Conditional Access Administrator
        '29232cdf-9323-42fd-ade2-1d097af3e4de' # public-id Exchange Administrator
        '729827e3-9c14-49f7-bb1b-9608f156bbb8' # public-id Helpdesk Administrator
        '966707d0-3269-4727-9be2-8c3a10f19b9d' # public-id Password Administrator
        '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' # public-id Privileged Authentication Administrator
        'e8611ab8-c189-46e8-94e1-60213ab1f814' # public-id Privileged Role Administrator
        '194ae4cb-b126-40b2-bd5b-6091b380977d' # public-id Security Administrator
        'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' # public-id SharePoint Administrator
        'fe930be7-5e62-47db-91af-98c3a49a38b1' # public-id User Administrator
    )
    $policyName = 'Require MFA for administrators'
    $columns = @('CustomerTenantId', 'CustomerName', 'UserPrincipalName', 'DisplayName', 'IsMfaRegistered', 'IsMfaCapable', 'MethodsRegistered', 'SecurityDefaultsEnabled', 'AdminMfaEnforcedBy', 'PolicyAction', 'Status', 'Error')
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

    function Test-PolicyCoversAdmin {
        param($Policy)
        if ($Policy.state -ne 'enabled') { return $false }
        $grant = $Policy.grantControls
        $requiresMfa = $grant -and ((@($grant.builtInControls) -contains 'mfa') -or $grant.authenticationStrength)
        if (-not $requiresMfa) { return $false }
        $users = $Policy.conditions.users
        $apps = @($Policy.conditions.applications.includeApplications)
        return ((@($users.includeUsers) -contains 'All') -or @($users.includeRoles | Where-Object { $_ }).Count -gt 0) -and ($apps -contains 'All')
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
        $customerName = $null
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $customerName = $knownNames[$tenant]
            if (-not $customerName) {
                $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $securityDefaults = [bool](Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy').isEnabled
            $enforcedBy = [System.Collections.Generic.List[string]]::new()
            $readProblems = [System.Collections.Generic.List[string]]::new()
            $caReadable = $true
            if ($securityDefaults) { $enforcedBy.Add('Security defaults') }
            else {
                try {
                    foreach ($policy in @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/identity/conditionalAccess/policies')) {
                        if (Test-PolicyCoversAdmin -Policy $policy) { $enforcedBy.Add([string]$policy.displayName) }
                    }
                }
                catch {
                    # Usually a tenant without Microsoft Entra ID P1. Report it and carry on.
                    $caReadable = $false
                    $readProblems.Add("Conditional Access not readable: $($_.Exception.Message)")
                }
            }

            $policyAction = 'None'
            if ($CreateConditionalAccessPolicy) {
                if ($enforcedBy.Count -gt 0) {
                    $policyAction = 'NotNeeded'
                }
                elseif (-not $caReadable) {
                    $policyAction = 'SkippedConditionalAccessNotReadable'
                }
                elseif (-not $Apply) {
                    $policyAction = 'WouldCreateReportOnly'
                }
                elseif ($PSCmdlet.ShouldProcess("$customerName ($tenant)", "Create Conditional Access policy '$policyName' in report-only state")) {
                    $body = @{
                        displayName   = $policyName
                        state         = 'enabledForReportingButNotEnforced'
                        conditions    = @{
                            clientAppTypes = @('all')
                            applications   = @{ includeApplications = @('All') }
                            users          = @{
                                includeRoles = $adminRoleIds
                                excludeUsers = @($ExcludeUserId | Where-Object { $_ })
                            }
                        }
                        grantControls = @{ operator = 'OR'; builtInControls = @('mfa') }
                    }
                    try {
                        $created = Invoke-MspGraphRequest -TenantId $tenant -Method POST -Uri 'v1.0/identity/conditionalAccess/policies' -Body $body -Confirm:$false
                        $policyAction = if ($created.id) { "CreatedReportOnly:$($created.id)" } else { 'Unconfirmed' }
                    }
                    catch {
                        $policyAction = "CreateFailed: $($_.Exception.Message)"
                        Write-Warning -Message "Could not create the policy in $tenant`: $($_.Exception.Message)"
                    }
                }
                else {
                    $policyAction = 'WhatIf'
                }
            }

            $admins = @()
            $registrationReadable = $true
            try {
                $admins = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/reports/authenticationMethods/userRegistrationDetails' | Where-Object { $_.isAdmin })
            }
            catch {
                $registrationReadable = $false
                $readProblems.Add("Registration details not readable: $($_.Exception.Message)")
            }
            $enforcedText = if ($enforcedBy.Count -gt 0) { $enforcedBy -join ', ' } elseif ($caReadable) { 'Nothing' } else { 'Unknown' }
            $problemText = if ($readProblems.Count -gt 0) { $readProblems -join ' | ' } else { $null }
            if ($admins.Count -eq 0) {
                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId        = $tenant
                    CustomerName            = $customerName
                    SecurityDefaultsEnabled = $securityDefaults
                    AdminMfaEnforcedBy      = $enforcedText
                    PolicyAction            = $policyAction
                    Status                  = if ($registrationReadable) { 'NoEnabledAdminsFound' } else { 'Partial' }
                    Error                   = $problemText
                }
                $results.Add($row)
                $row
                continue
            }

            foreach ($admin in $admins) {
                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId        = $tenant
                    CustomerName            = $customerName
                    UserPrincipalName       = $admin.userPrincipalName
                    DisplayName             = $admin.userDisplayName
                    IsMfaRegistered         = $admin.isMfaRegistered
                    IsMfaCapable            = $admin.isMfaCapable
                    MethodsRegistered       = @($admin.methodsRegistered) -join ', '
                    SecurityDefaultsEnabled = $securityDefaults
                    AdminMfaEnforcedBy      = $enforcedText
                    PolicyAction            = $policyAction
                    Status                  = if ($admin.isMfaRegistered -and $enforcedBy.Count -gt 0) { 'OK' } else { 'Attention' }
                    Error                   = $problemText
                }
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $row = Get-ResultRow -Column $columns -Value @{
                CustomerTenantId = $tenant
                CustomerName     = $customerName
                Status           = 'Failed'
                Error            = $_.Exception.Message
            }
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
