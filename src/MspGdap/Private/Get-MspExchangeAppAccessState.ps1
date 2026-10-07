function Get-MspExchangeAppAccessState {
    <#
    .SYNOPSIS
        Reads, through Microsoft Graph in the customer, whether an automation app can use Exchange Online app-only.
    .DESCRIPTION
        Checks the Office 365 Exchange Online service principal, the automation app's service principal, the
        Exchange.ManageAsApp app role assignment, and directory role assignments to the app's service principal
        (unified RBAC: roleManagement/directory/roleAssignments, scope "/").
        The Exchange.ManageAsApp role ID is read from the Exchange service principal's published appRoles.
        Throws when Graph cannot be read.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AppId,
        [Parameter(Mandatory)][object[]]$Role
    )
    $exchangeAppId = '00000002-0000-0ff1-ce00-000000000000'
    $exoSp = Get-MspServicePrincipalByAppId -TenantId $TenantId -AppId $exchangeAppId -Select 'id,appId,displayName,appRoles'
    $manageAsAppRoleId = $null
    if ($exoSp) {
        $published = @($exoSp.appRoles | Where-Object { $_.value -eq 'Exchange.ManageAsApp' }) | Select-Object -First 1
        if ($published) { $manageAsAppRoleId = [string]$published.id }
    }
    $appSp = Get-MspServicePrincipalByAppId -TenantId $TenantId -AppId $AppId -Select 'id,appId,displayName,appOwnerOrganizationId'

    $hasManageAsApp = $false
    $roleStates = @()
    if ($appSp) {
        if ($exoSp -and $manageAsAppRoleId) {
            $assignments = @(Invoke-MspGraphCall -TenantId $TenantId -Path ("servicePrincipals/{0}/appRoleAssignments" -f $appSp.id))
            $hasManageAsApp = @($assignments | Where-Object { $_.resourceId -eq $exoSp.id -and $_.appRoleId -eq $manageAsAppRoleId }).Count -gt 0
        }
        $roleStates = foreach ($wanted in $Role) {
            $path = "roleManagement/directory/roleAssignments?`$filter=principalId eq '{0}' and roleDefinitionId eq '{1}'" -f $appSp.id, $wanted.RoleTemplateId
            $found = @(Invoke-MspGraphCall -TenantId $TenantId -Path $path | Where-Object { -not $_.directoryScopeId -or $_.directoryScopeId -eq '/' }) | Select-Object -First 1
            [pscustomobject]@{
                DisplayName    = $wanted.DisplayName
                RoleTemplateId = $wanted.RoleTemplateId
                Assigned       = [bool]$found
                AssignmentId   = if ($found) { [string]$found.id } else { $null }
            }
        }
    }
    else {
        $roleStates = foreach ($wanted in $Role) {
            [pscustomobject]@{ DisplayName = $wanted.DisplayName; RoleTemplateId = $wanted.RoleTemplateId; Assigned = $false; AssignmentId = $null }
        }
    }

    [pscustomobject]@{
        TenantId                 = $TenantId
        ExchangeServicePrincipal = $exoSp
        ManageAsAppRoleId        = $manageAsAppRoleId
        AppServicePrincipal      = $appSp
        HasManageAsApp           = $hasManageAsApp
        Roles                    = @($roleStates)
    }
}
