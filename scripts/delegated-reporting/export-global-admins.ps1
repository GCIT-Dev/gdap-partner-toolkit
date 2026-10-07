#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports every customer's Global Administrators and, after you have reviewed the list, blocks sign-in for the accounts you choose.

.DESCRIPTION
    Report (default): for each customer the script reads the active Global
    Administrator role assignments through Microsoft Graph
    (roleManagement/directory/roleAssignments) and returns one row per member,
    with whether the account is enabled, whether it has a licence and, where the
    tenant's licences allow it, its last sign-in. Groups and service principals
    that hold the role are listed too. Use -UnlicensedOnly to return only
    unlicensed user accounts, which are often the default admin created when the
    tenant was set up.

    Block (-BlockFromCsv): pass a copy of the report that you have reviewed and cut
    down to the accounts that should be blocked. For each row the script checks
    that the user is still an enabled Global Administrator in that tenant, and that
    blocking it would still leave at least -MinimumEnabledAdmins enabled Global
    Administrator user accounts, so a customer is never locked out of their own
    tenant. Without -Apply it reports what it would block. With -Apply it sets
    accountEnabled to false, one ShouldProcess confirmation per account (supports
    -WhatIf), and reads the account back.

    -RoleName reports on another directory role, as the original's $RoleName did.
    The column EnabledGlobalAdminCount then counts enabled members of that role.

    Only active assignments are read. Eligible assignments in Privileged Identity
    Management are not included. Review break-glass accounts with the customer
    before you block anything.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER BlockFromCsv
    Path of a reviewed CSV with CustomerTenantId and UserId columns (as produced by
    this script). Every row is an account to block.

.PARAMETER MinimumEnabledAdmins
    The number of enabled Global Administrator user accounts that must remain in a
    tenant after blocking. Default 1.

.PARAMETER UnlicensedOnly
    In report mode, returns only user accounts without a licence.

.PARAMETER RoleName
    In report mode, the directory role to list instead of Global Administrator,
    such as 'Exchange Administrator' or 'User Administrator' (the display name, as
    in the original article's $RoleName). Blocking always works on Global
    Administrators.

.PARAMETER Apply
    With -BlockFromCsv, blocks sign-in for the listed accounts. Without it the
    script only reports.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-global-admins.ps1 -AllCustomers -UnlicensedOnly -OutputPath ./unlicensed-global-admins.csv

    Exports unlicensed Global Administrators in every active GDAP customer for review.

.EXAMPLE
    ./export-global-admins.ps1 -BlockFromCsv ./reviewed-admins-to-block.csv -Apply -WhatIf

    Shows which reviewed accounts would be blocked. Remove -WhatIf to block them.

.NOTES
    Replaces the original 2017 method: Connect-MsolService -Credential (no MFA), Get-MsolPartnerContract -All (DAP), Get-MsolRole 'Company Administrator', Get-MsolRoleMember and Set-MsolUser -BlockCredential $true (MSOnline module, retired 30 May 2025).
    Required GDAP roles: Global Reader to report, Privileged Authentication Administrator to block Global Administrators.
    Required partner app permissions: Microsoft Graph delegated RoleManagement.Read.Directory, User.Read.All and AuditLog.Read.All to report, User.EnableDisableAccount.All to block (covered in manifests/partner-app.full.json by RoleManagement.ReadWrite.Directory, User.ReadWrite.All and AuditLog.Read.All).

.LINK
    https://gcit.com.au/knowledge-base/get-list-every-customers-office-365-administrators-via-powershell-delegated-administration/

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

    [Parameter(Mandatory, ParameterSetName = 'Block')]
    [ValidateNotNullOrEmpty()]
    [string]$BlockFromCsv,

    [Parameter(ParameterSetName = 'Block')]
    [ValidateRange(1, 10)]
    [int]$MinimumEnabledAdmins = 1,

    [Parameter(ParameterSetName = 'Tenant')]
    [Parameter(ParameterSetName = 'AllCustomers')]
    [switch]$UnlicensedOnly,

    [Parameter(ParameterSetName = 'Tenant')]
    [Parameter(ParameterSetName = 'AllCustomers')]
    [ValidatePattern('^[A-Za-z0-9 ._()-]+$')]
    [string]$RoleName = 'Global Administrator',

    [Parameter(ParameterSetName = 'Block')]
    [switch]$Apply,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    # Global Administrator role template ID (the same in every tenant).
    $globalAdminRoleId = '62e90394-69f5-4237-9190-012177145e10' # public-id (Global Administrator role template)
    $columns = @('CustomerTenantId', 'CustomerName', 'RoleName', 'PrincipalType', 'DisplayName', 'UserPrincipalName', 'UserId', 'AccountEnabled', 'IsLicensed', 'LastSignInDateTime', 'EnabledGlobalAdminCount', 'Action', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()
    $blockList = @()

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

    function Get-GlobalAdminMember {
        param([string]$Tenant, [string]$RoleId)
        $uri = "v1.0/roleManagement/directory/roleAssignments?`$filter=roleDefinitionId eq '$RoleId'&`$expand=principal"
        foreach ($assignment in @(Invoke-MspGraphRequest -TenantId $Tenant -Uri $uri)) {
            $principal = $assignment.principal
            $type = if ($principal) { ([string]$principal.'@odata.type') -replace '^#microsoft\.graph\.', '' } else { 'unknown' }
            $member = [ordered]@{
                PrincipalType      = $type
                Id                 = if ($principal) { $principal.id } else { $assignment.principalId }
                DisplayName        = if ($principal) { $principal.displayName } else { $null }
                UserPrincipalName  = $null
                AccountEnabled     = $null
                IsLicensed         = $null
                LastSignInDateTime = $null
            }
            if ($type -eq 'user') {
                $user = $null
                try {
                    $user = Invoke-MspGraphRequest -TenantId $Tenant -Uri ('v1.0/users/{0}?$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses,signInActivity' -f $member.Id)
                }
                catch {
                    # signInActivity needs Microsoft Entra ID P1 or P2. Read the user again without it.
                    $user = Invoke-MspGraphRequest -TenantId $Tenant -Uri ('v1.0/users/{0}?$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses' -f $member.Id)
                }
                $member.DisplayName = $user.displayName
                $member.UserPrincipalName = $user.userPrincipalName
                $member.AccountEnabled = $user.accountEnabled
                $member.IsLicensed = @($user.assignedLicenses | Where-Object { $_ }).Count -gt 0
                if ($user.PSObject.Properties['signInActivity'] -and $user.signInActivity) {
                    $member.LastSignInDateTime = $user.signInActivity.lastSignInDateTime
                }
            }
            [pscustomobject]$member
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
    if ($PSCmdlet.ParameterSetName -eq 'Block') {
        $blockList = @(Import-Csv -Path $BlockFromCsv | Where-Object { $_.CustomerTenantId -and $_.UserId })
        if ($blockList.Count -eq 0) {
            throw "No rows with CustomerTenantId and UserId were found in $BlockFromCsv."
        }
        foreach ($tenantValue in @($blockList.CustomerTenantId | Sort-Object -Unique)) { $targets.Add($tenantValue) }
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

            $roleId = $globalAdminRoleId
            $roleLabel = 'Global Administrator'
            if ($PSCmdlet.ParameterSetName -ne 'Block' -and $RoleName -ne 'Global Administrator') {
                $definitionUri = "v1.0/roleManagement/directory/roleDefinitions?`$filter=displayName eq '{0}'&`$select=id,displayName" -f ($RoleName -replace "'", "''")
                $definition = @(Invoke-MspGraphRequest -TenantId $tenant -Uri $definitionUri) | Select-Object -First 1
                if (-not $definition) { throw "No role named '$RoleName' was found in this tenant." }
                $roleId = [string]$definition.id
                $roleLabel = [string]$definition.displayName
            }

            $members = @(Get-GlobalAdminMember -Tenant $tenant -RoleId $roleId)
            $enabledUsers = @($members | Where-Object { $_.PrincipalType -eq 'user' -and $_.AccountEnabled })
            $remaining = $enabledUsers.Count

            if ($PSCmdlet.ParameterSetName -ne 'Block') {
                foreach ($member in $members) {
                    if ($UnlicensedOnly -and ($member.PrincipalType -ne 'user' -or $member.IsLicensed)) { continue }
                    $row = Get-ResultRow -Column $columns -Value @{
                        CustomerTenantId        = $tenant
                        CustomerName            = $customerName
                        RoleName                = $roleLabel
                        PrincipalType           = $member.PrincipalType
                        DisplayName             = $member.DisplayName
                        UserPrincipalName       = $member.UserPrincipalName
                        UserId                  = $member.Id
                        AccountEnabled          = $member.AccountEnabled
                        IsLicensed              = $member.IsLicensed
                        LastSignInDateTime      = $member.LastSignInDateTime
                        EnabledGlobalAdminCount = $enabledUsers.Count
                        Action                  = 'None'
                        Status                  = 'OK'
                    }
                    $results.Add($row)
                    $row
                }
                continue
            }

            foreach ($entry in @($blockList | Where-Object { $_.CustomerTenantId -eq $target -or $_.CustomerTenantId -eq $tenant })) {
                $member = $members | Where-Object { $_.Id -eq $entry.UserId } | Select-Object -First 1
                $values = [ordered]@{
                    CustomerTenantId        = $tenant
                    CustomerName            = $customerName
                    RoleName                = $roleLabel
                    PrincipalType           = 'user'
                    DisplayName             = if ($member) { $member.DisplayName } else { $entry.DisplayName }
                    UserPrincipalName       = if ($member) { $member.UserPrincipalName } else { $entry.UserPrincipalName }
                    UserId                  = $entry.UserId
                    AccountEnabled          = if ($member) { $member.AccountEnabled } else { $null }
                    IsLicensed              = if ($member) { $member.IsLicensed } else { $null }
                    LastSignInDateTime      = if ($member) { $member.LastSignInDateTime } else { $null }
                    EnabledGlobalAdminCount = $enabledUsers.Count
                    Status                  = 'OK'
                }

                if (-not $member -or $member.PrincipalType -ne 'user') {
                    $values.Action = 'SkippedNotGlobalAdmin'
                }
                elseif (-not $member.AccountEnabled) {
                    $values.Action = 'AlreadyBlocked'
                }
                elseif (($remaining - 1) -lt $MinimumEnabledAdmins) {
                    $values.Action = 'SkippedLastEnabledAdmin'
                    $values.Status = 'Warning'
                }
                elseif (-not $Apply) {
                    $values.Action = 'WouldBlock'
                    $remaining--
                }
                elseif ($PSCmdlet.ShouldProcess("$($member.UserPrincipalName) in $customerName ($tenant)", 'Block sign-in (accountEnabled = false)')) {
                    try {
                        Invoke-MspGraphRequest -TenantId $tenant -Method PATCH -Uri "v1.0/users/$($member.Id)" -Body @{ accountEnabled = $false } -Confirm:$false | Out-Null
                        $check = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/users/{0}?$select=accountEnabled' -f $member.Id)
                        $values.AccountEnabled = $check.accountEnabled
                        $values.Action = if ($check.accountEnabled -eq $false) { 'Blocked' } else { 'Unconfirmed' }
                        if ($check.accountEnabled -ne $false) { $values.Status = 'Warning' }
                        $remaining--
                    }
                    catch {
                        $values.Action = 'BlockFailed'
                        $values.Status = 'Failed'
                        $values.Error = $_.Exception.Message
                    }
                }
                else {
                    $values.Action = 'WhatIf'
                    $remaining--
                }

                $row = Get-ResultRow -Column $columns -Value $values
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
