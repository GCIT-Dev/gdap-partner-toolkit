#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft Entra admin role additions and removals in customer tenants.

.DESCRIPTION
    Reads the Microsoft Entra directory audit log (GET /auditLogs/directoryAudits) in each customer for
    role management events in the last -Hours hours, and returns one row per role member added, removed,
    made eligible or activated. Unlike the original, removals are reported as well as additions, and no
    state needs to be stored between runs because the audit log is the record of change.

    With -IncludeCurrentAssignments the script also returns every current active role assignment
    (GET /roleManagement/directory/roleAssignments) as rows with Activity "CurrentAssignment", which is
    useful as a baseline.

    Directory audit logs are kept for 7 days in Microsoft Entra ID Free and 30 days with Microsoft Entra ID
    P1 or P2, so run the script at least daily. The AdminRoleChangeAlert Azure Function in the functions
    folder runs it on a timer and queues one message per change.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER Hours
    How far back to look in the audit log, in hours. Default 25 (one day plus a margin).

.PARAMETER IncludeCurrentAssignments
    Also return every current active directory role assignment.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-admin-role-changes.ps1 -AllCustomers -Hours 168 -OutputPath ./role-changes.csv

    Lists every admin role change across all customers in the last week.

.EXAMPLE
    ./get-admin-role-changes.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeCurrentAssignments

    Lists recent role changes and every current role holder in one customer.

.NOTES
    Replaces the original 2018 method: an Azure Functions v1 timer function using MSOnline and DAP
    (Get-MsolPartnerContract, Get-MsolRole, Get-MsolRoleMember) with an AES-encrypted stored password,
    comparing role members against Azure Table storage accessed with the storage account key.
    Required GDAP roles: Security Reader or Reports Reader (audit log), plus Global Reader for
    -IncludeCurrentAssignments.
    Required partner app permissions: Microsoft Graph delegated AuditLog.Read.All, and
    RoleManagement.Read.Directory or RoleManagement.ReadWrite.Directory for -IncludeCurrentAssignments
    (AuditLog.Read.All and RoleManagement.ReadWrite.Directory are in the full manifest).

.LINK
    https://gcit.com.au/knowledge-base/monitor-office-365-admin-role-changes-in-all-customer-tenants/

.LINK
    docs/07-migrating-from-dap-msonline.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateRange(1, 720)]
    [int]$Hours = 25,

    [switch]$IncludeCurrentAssignments,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-ResultRow {
        param($Customer, $When, $Activity, $Result, $Role, $TargetType, $Target, $InitiatedBy, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId = $Customer.TenantId
            CustomerName     = $Customer.Name
            ActivityDateTime = $When
            Activity         = $Activity
            Result           = $Result
            RoleName         = $Role
            TargetType       = $TargetType
            Target           = $Target
            InitiatedBy      = $InitiatedBy
            Error            = $ErrorMessage
        }
    }

    function Get-ModifiedValue {
        param($Resource, [string]$Name)
        $property = @($Resource.modifiedProperties | Where-Object { $_.displayName -eq $Name }) | Select-Object -First 1
        if (-not $property) { return $null }
        # Removals carry the role in oldValue, and newValue is an empty JSON string ("").
        $value = ([string]$property.newValue).Trim('"')
        if (-not $value) { $value = ([string]$property.oldValue).Trim('"') }
        $value
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }
    $since = [datetime]::UtcNow.AddHours(-$Hours).ToString('yyyy-MM-ddTHH:mm:ssZ')

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $filter = "category eq 'RoleManagement' and activityDateTime ge $since"
            $events = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "auditLogs/directoryAudits?`$filter=$filter" |
                    Where-Object { $_.activityDisplayName -match 'member (to|from) role' })

            foreach ($auditEvent in $events) {
                $roleResource = @($auditEvent.targetResources | Where-Object { $_.type -eq 'Role' }) | Select-Object -First 1
                $member = @($auditEvent.targetResources | Where-Object { $_.type -ne 'Role' }) | Select-Object -First 1
                $roleName = Get-ModifiedValue -Resource $member -Name 'Role.DisplayName'
                if (-not $roleName -and $roleResource) { $roleName = $roleResource.displayName }
                $initiatedBy = if ($auditEvent.initiatedBy.user.userPrincipalName) {
                    $auditEvent.initiatedBy.user.userPrincipalName
                }
                elseif ($auditEvent.initiatedBy.app.displayName) {
                    "App: $($auditEvent.initiatedBy.app.displayName)"
                }
                else {
                    $null
                }
                $target = if ($member.userPrincipalName) { $member.userPrincipalName } else { $member.displayName }
                $row = ConvertTo-ResultRow -Customer $customer -When $auditEvent.activityDateTime -Activity $auditEvent.activityDisplayName -Result $auditEvent.result -Role $roleName -TargetType $member.type -Target $target -InitiatedBy $initiatedBy
                $results.Add($row)
                $row
            }

            if ($IncludeCurrentAssignments) {
                $definitions = @{}
                Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'roleManagement/directory/roleDefinitions?$select=id,displayName' |
                    ForEach-Object { $definitions[[string]$_.id] = $_.displayName }
                $assignments = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'roleManagement/directory/roleAssignments?$expand=principal')
                foreach ($assignment in $assignments) {
                    $principal = $assignment.principal
                    $principalType = if ($principal.'@odata.type') { ([string]$principal.'@odata.type') -replace '^#microsoft\.graph\.', '' } else { $null }
                    $target = if ($principal.userPrincipalName) { $principal.userPrincipalName } else { $principal.displayName }
                    $row = ConvertTo-ResultRow -Customer $customer -Activity 'CurrentAssignment' -Role $definitions[[string]$assignment.roleDefinitionId] -TargetType $principalType -Target $target
                    $results.Add($row)
                    $row
                }
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Activity 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
