#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Prepares a departing user's OneDrive for handover by making the successor the user's manager, so OneDrive access is delegated to them when the account is deleted.

.DESCRIPTION
    The original article made a Global Administrator a site collection admin on
    two OneDrives and copied every file with the legacy PnP module, using a
    password credential and a temporary admin account without MFA. None of that
    works or is acceptable today. A GDAP technician is not a user in the customer
    tenant and cannot be made a site collection admin of a user's OneDrive.

    The supported route is OneDrive's automatic access delegation. When a user
    account is deleted, the user's manager in Microsoft Entra ID is given access to
    the OneDrive and emailed a link, and the OneDrive is kept for the retention
    period (30 days by default). The successor then opens the OneDrive and moves
    or copies the files they need.

    This script:
    1. Reads the departing user and the successor through Microsoft Graph.
    2. Reports the departing user's current manager.
    3. With -Apply, sets the successor as the departing user's manager
       (PUT /users/{id}/manager/$ref), then reads it back.
    4. With -Apply and -RemoveDepartingUser, deletes the departing user once the
       manager is confirmed, which starts the OneDrive handover. This is the same
       soft delete as the Microsoft 365 admin center (restorable for 30 days).

    Without -Apply it only reports. Every change goes through ShouldProcess and
    supports -WhatIf.

    Check in the SharePoint admin center that "Enable access delegation" is
    selected under My Site Cleanup (it is on by default), and that the OneDrive
    retention period is long enough for the handover. Deleting the user from the
    Microsoft 365 admin center, which offers "Give another user access to this
    user's OneDrive", is an equivalent manual route.

.PARAMETER TenantId
    The customer tenant ID (GUID) or verified domain, such as
    contoso.onmicrosoft.com. Accepts pipeline input. Normally one tenant.

.PARAMETER AllCustomers
    Accepted for consistency with the other scripts. Not useful here, because the
    users only exist in one tenant, and it cannot be combined with -Apply.

.PARAMETER DepartingUserPrincipalName
    User principal name or object ID of the user who is leaving.

.PARAMETER SuccessorUserPrincipalName
    User principal name or object ID of the user who should receive the OneDrive
    files.

.PARAMETER Apply
    Sets the successor as manager. Without it the script only reports.

.PARAMETER RemoveDepartingUser
    With -Apply, deletes the departing user after the manager change is confirmed.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./set-onedrive-successor.ps1 -TenantId 'contoso.onmicrosoft.com' -DepartingUserPrincipalName 'leaver@contoso.com' -SuccessorUserPrincipalName 'manager@contoso.com'

    Reports the departing user's current manager and what would change.

.EXAMPLE
    ./set-onedrive-successor.ps1 -TenantId 'contoso.onmicrosoft.com' -DepartingUserPrincipalName 'leaver@contoso.com' -SuccessorUserPrincipalName 'manager@contoso.com' -Apply -RemoveDepartingUser -WhatIf

    Shows the manager change and the deletion without making them. Remove -WhatIf to run them.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Connect-SPOService and Set-SPOUser -IsSiteCollectionAdmin with a Global Administrator password, then Connect-PnPOnline -Credentials and Copy-PnPFile with the retired SharePointPnPPowerShellOnline module (and advice to use an admin account without MFA, removed as unsafe).
    Required GDAP roles: User Administrator (Privileged Authentication Administrator if the departing user holds an admin role).
    Required partner app permissions: Microsoft Graph delegated User.ReadWrite.All (in manifests/partner-app.full.json).

.LINK
    https://gcit.com.au/knowledge-base/transfer-users-onedrive-files-another-user-via-powershell/

.LINK
    https://learn.microsoft.com/en-us/sharepoint/retention-and-deletion

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

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$DepartingUserPrincipalName,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$SuccessorUserPrincipalName,

    [switch]$Apply,

    [switch]$RemoveDepartingUser,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    if ($Apply -and $AllCustomers) {
        throw '-Apply cannot be used with -AllCustomers. Name the customer with -TenantId.'
    }
    if ($RemoveDepartingUser -and -not $Apply) {
        Write-Warning -Message '-RemoveDepartingUser has no effect without -Apply. Reporting only.'
    }

    $columns = @('CustomerTenantId', 'CustomerName', 'DepartingUser', 'DepartingUserId', 'Successor', 'SuccessorId', 'PreviousManager', 'ManagerAction', 'RemoveAction', 'NextStep', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()
    $userSelect = 'id,displayName,userPrincipalName,accountEnabled'

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
        $values = [ordered]@{
            CustomerTenantId = $target
            DepartingUser    = $DepartingUserPrincipalName
            Successor        = $SuccessorUserPrincipalName
            ManagerAction    = 'None'
            RemoveAction     = 'None'
        }
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $values.CustomerTenantId = $tenant
            $customerName = $knownNames[$tenant]
            if (-not $customerName) {
                $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }
            $values.CustomerName = $customerName

            $departing = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/users/{0}?$select={1}' -f [uri]::EscapeDataString($DepartingUserPrincipalName), $userSelect)
            $successor = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/users/{0}?$select={1}' -f [uri]::EscapeDataString($SuccessorUserPrincipalName), $userSelect)
            if ($departing.id -eq $successor.id) {
                throw 'The departing user and the successor are the same account.'
            }
            $values.DepartingUser = $departing.userPrincipalName
            $values.DepartingUserId = $departing.id
            $values.Successor = $successor.userPrincipalName
            $values.SuccessorId = $successor.id

            $manager = $null
            try {
                $manager = Invoke-MspGraphRequest -TenantId $tenant -Uri "v1.0/users/$($departing.id)/manager?`$select=id,userPrincipalName"
            }
            catch {
                # Graph returns 404 when no manager is set. Treat that as no manager.
                Write-Verbose -Message "No manager returned for $($departing.userPrincipalName): $($_.Exception.Message)"
            }
            $values.PreviousManager = if ($manager) { $manager.userPrincipalName } else { $null }
            $managerReady = $manager -and $manager.id -eq $successor.id

            if ($managerReady) {
                $values.ManagerAction = 'AlreadySet'
            }
            elseif (-not $Apply) {
                $values.ManagerAction = 'WouldSet'
            }
            elseif ($PSCmdlet.ShouldProcess("$($departing.userPrincipalName) in $customerName ($tenant)", "Set manager to $($successor.userPrincipalName)")) {
                $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($successor.id)" }
                Invoke-MspGraphRequest -TenantId $tenant -Method PUT -Uri "v1.0/users/$($departing.id)/manager/`$ref" -Body $body -Confirm:$false | Out-Null
                $check = Invoke-MspGraphRequest -TenantId $tenant -Uri "v1.0/users/$($departing.id)/manager?`$select=id"
                if ($check.id -eq $successor.id) {
                    $values.ManagerAction = 'Set'
                    $managerReady = $true
                }
                else {
                    $values.ManagerAction = 'Unconfirmed'
                }
            }
            else {
                $values.ManagerAction = 'WhatIf'
            }

            if ($RemoveDepartingUser -and $Apply) {
                if ($values.ManagerAction -eq 'WhatIf') {
                    # The manager change was only previewed, so preview the deletion as well.
                    $null = $PSCmdlet.ShouldProcess("$($departing.userPrincipalName) in $customerName ($tenant)", 'Delete user (OneDrive access passes to the manager)')
                    $values.RemoveAction = 'WhatIf'
                }
                elseif (-not $managerReady) {
                    $values.RemoveAction = 'SkippedManagerNotConfirmed'
                }
                elseif ($PSCmdlet.ShouldProcess("$($departing.userPrincipalName) in $customerName ($tenant)", 'Delete user (OneDrive access passes to the manager)')) {
                    Invoke-MspGraphRequest -TenantId $tenant -Method DELETE -Uri "v1.0/users/$($departing.id)" -Confirm:$false | Out-Null
                    $values.RemoveAction = 'Deleted'
                }
                else {
                    $values.RemoveAction = 'WhatIf'
                }
            }

            $values.NextStep = switch ($true) {
                ($values.RemoveAction -eq 'Deleted') { "$($successor.userPrincipalName) will be emailed a link to the OneDrive. Copy the files before the OneDrive retention period ends."; break }
                ($managerReady) { "Delete $($departing.userPrincipalName) (or rerun with -Apply -RemoveDepartingUser) to hand the OneDrive to $($successor.userPrincipalName)."; break }
                default { 'Rerun with -Apply to set the successor as manager.' }
            }
            $values.Status = if ($values.ManagerAction -eq 'Unconfirmed') { 'Warning' } else { 'OK' }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $values.CustomerName = $customerName
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
