function Enable-MspExchangeAppAccess {
    <#
    .SYNOPSIS
        Gives a SEPARATE automation app app-only access to Exchange Online in one customer, and proves it by readback.
    .DESCRIPTION
        For unattended jobs (Connect-ExchangeOnline -AppId -CertificateThumbprint -Organization). In the customer:
          1. checks the Office 365 Exchange Online service principal exists
          2. ensures the automation app's service principal exists (created through Graph if allowed, otherwise
             the admin consent URL is returned for a Cloud Application Administrator to open)
          3. grants the Exchange.ManageAsApp app role (Graph appRoleAssignedTo). Partner Center cannot consent
             application permissions under GDAP, so this is done directly in the customer.
          4. assigns each directory role (default Exchange Administrator) to the app's service principal with the
             unified RBAC API. Roles cannot be given to apps through GDAP. This needs Privileged Role Administrator.
        Every step is read back. A step that cannot be confirmed is Failed, and the result's Success is false.
        A Failed result is also written as a non-terminating error (MspGdap.Enable-MspExchangeAppAccess.Failed).
        Guardrails, all checked before any write:
          - it refuses to run against the MspGdap partner app: keep app-only permissions in their own app
          - the app must be registered in YOUR partner tenant (or its service principal in the customer must
            be owned by your partner tenant), so a typo or a third-party app ID is never granted access.
            -AllowExternalApp overrides this.
          - roles are limited to the ones Exchange Online app-only supports, Global Administrator needs
            -AllowGlobalAdministrator and roles beyond Exchange need -AllowPrivilegedRole
          - your partner tenant is refused as -TenantId
        It never adds credentials to any service principal.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .PARAMETER AppId
        Application ID of your automation app (created from manifests/automation-app.example.json).
    .PARAMETER Role
        Directory roles for the app. Default 'Exchange Administrator'. 'Exchange Recipient Administrator',
        'Global Reader' and 'Security Reader' are also allowed without a switch.
    .PARAMETER AllowGlobalAdministrator
        Permit Global Administrator in -Role. Not recommended.
    .PARAMETER AllowPrivilegedRole
        Permit Compliance Administrator, Security Administrator or Helpdesk Administrator in -Role.
    .PARAMETER AllowExternalApp
        Skip the check that the app belongs to your partner tenant. Only for an automation app you own
        that is registered in another tenant.
    .PARAMETER ReadbackTimeoutSeconds
        How long to keep reading back each change before reporting it as Failed. Default 60.
    .EXAMPLE
        Enable-MspExchangeAppAccess -TenantId 'fabrikam.onmicrosoft.com' -AppId '<AutomationAppId>' -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [ValidateNotNullOrEmpty()][string[]]$Role = @('Exchange Administrator'),
        [switch]$AllowGlobalAdministrator,
        [switch]$AllowPrivilegedRole,
        [switch]$AllowExternalApp,
        [ValidateRange(0, 600)][int]$ReadbackTimeoutSeconds = 60
    )
    process {
        $AppId = $AppId.ToLowerInvariant()
        $steps = New-Object System.Collections.Generic.List[object]
        $info = [ordered]@{ AppId = $AppId; ServicePrincipalId = $null; InitialDomain = $null; AdminConsentUrl = $null; Roles = @() }
        $tenant = $TenantId
        $finish = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Enable-MspExchangeAppAccess' -TenantId $tenant -Target $AppId -Steps $steps.ToArray() -Property $info }

        try {
            $roles = @(Resolve-MspExchangeAppRole -Role $Role -AllowGlobalAdministrator:$AllowGlobalAdministrator -AllowPrivilegedRole:$AllowPrivilegedRole)
            $info.Roles = @($roles | ForEach-Object { $_.DisplayName })
            $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Prepare' -Status Failed -Detail $_.Exception.Message))
            return (& $finish)
        }
        $info.AdminConsentUrl = 'https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&scope=https://outlook.office365.com/.default' -f $tenant, $AppId

        $partnerAppId = $null
        try { $partnerAppId = Resolve-MspPartnerAppId } catch { $partnerAppId = $null }
        if ($partnerAppId -and $partnerAppId -eq $AppId) {
            $steps.Add((New-MspStepResult -Step 'Check app' -Status Failed -Detail 'This is the MspGdap partner app. App-only Exchange access belongs in a separate automation app (see manifests/automation-app.example.json).'))
            return (& $finish)
        }

        try { $state = Get-MspExchangeAppAccessState -TenantId $tenant -AppId $AppId -Role $roles }
        catch {
            $steps.Add((New-MspStepResult -Step 'Read current state' -Status Unknown -Detail "Could not read the customer through Graph. Nothing was changed. $($_.Exception.Message)"))
            return (& $finish)
        }

        # The app must belong to the partner. Checked before any write.
        $ownerStep = $null
        $partnerTenantId = $null
        try { $partnerTenantId = ([string](Get-MspConfiguration -ErrorAction Stop).PartnerTenantId).ToLowerInvariant() } catch { $partnerTenantId = $null }
        if ($state.AppServicePrincipal -and $state.AppServicePrincipal.PSObject.Properties['appOwnerOrganizationId'] -and $state.AppServicePrincipal.appOwnerOrganizationId) {
            $owner = ([string]$state.AppServicePrincipal.appOwnerOrganizationId).ToLowerInvariant()
            if ($partnerTenantId -and $owner -eq $partnerTenantId) {
                $ownerStep = New-MspStepResult -Step 'Check app owner' -Status Passed -Detail "The service principal in the customer belongs to your partner tenant ($owner)."
            }
            else {
                $ownerStep = New-MspStepResult -Step 'Check app owner' -Status Failed -Detail "App $AppId is owned by tenant $owner, not your partner tenant. Nothing was changed. Use -AllowExternalApp only if you own that app."
            }
        }
        else {
            try {
                $registered = @(Invoke-MspGraphCall -PartnerTenant -Path ("applications?`$filter=appId eq '{0}'&`$select=id,appId,displayName" -f $AppId)) | Select-Object -First 1
                if ($registered) {
                    $ownerStep = New-MspStepResult -Step 'Check app owner' -Status Passed -Detail "App '$($registered.displayName)' is registered in your partner tenant."
                }
                else {
                    $ownerStep = New-MspStepResult -Step 'Check app owner' -Status Failed -Detail "App $AppId is not registered in your partner tenant. Nothing was changed. Check -AppId, or use -AllowExternalApp only if you own that app in another tenant."
                }
            }
            catch {
                $ownerStep = New-MspStepResult -Step 'Check app owner' -Status Unknown -Detail "Could not read the partner tenant to confirm the app is yours. Nothing was changed. $($_.Exception.Message)"
            }
        }
        if ($ownerStep.Status -ne 'Passed' -and $AllowExternalApp) {
            $ownerStep = New-MspStepResult -Step 'Check app owner' -Status Warning -Detail "Not confirmed as a partner app, continuing because -AllowExternalApp was used. $($ownerStep.Detail)"
        }
        $steps.Add($ownerStep)
        if ($ownerStep.Status -in 'Failed', 'Unknown') { return (& $finish) }

        if (-not $state.ExchangeServicePrincipal) {
            $steps.Add((New-MspStepResult -Step 'Exchange Online service principal' -Status Failed -Detail 'Office 365 Exchange Online has no service principal in this customer (no Exchange licence?). Nothing was changed.'))
            return (& $finish)
        }
        if (-not $state.ManageAsAppRoleId) {
            $steps.Add((New-MspStepResult -Step 'Exchange Online service principal' -Status Failed -Detail 'The Exchange Online service principal does not publish Exchange.ManageAsApp. Nothing was changed.'))
            return (& $finish)
        }
        $exoSp = $state.ExchangeServicePrincipal
        $steps.Add((New-MspStepResult -Step 'Exchange Online service principal' -Status Passed -Detail "Present ($($exoSp.id))."))

        # Automation app service principal
        $appSp = $state.AppServicePrincipal
        if ($appSp) {
            $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status Passed -Detail "Present ($($appSp.id))."))
        }
        elseif ($PSCmdlet.ShouldProcess("customer $tenant", "Create service principal for automation app $AppId")) {
            try {
                $null = Invoke-MspGraphCall -TenantId $tenant -Method POST -Path 'servicePrincipals' -Body @{ appId = $AppId }
                $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition { Get-MspServicePrincipalByAppId -TenantId $tenant -AppId $AppId }
                if ($check.Satisfied) {
                    $appSp = $check.Value
                    $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status Changed -Detail "Created and confirmed ($($appSp.id))."))
                }
                else {
                    $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status Failed -Detail "Created but not confirmed. $($check.LastError)"))
                }
            }
            catch {
                $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status Failed -Detail "Could not create it ($($_.Exception.Message)). The partner app may lack Application.ReadWrite.All. A Cloud Application Administrator can open the AdminConsentUrl in this result instead, then run this command again."))
            }
        }
        else {
            $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status WhatIf -Detail 'Would create the service principal.'))
        }
        if ($appSp) { $info.ServicePrincipalId = [string]$appSp.id }

        if (-not $appSp) {
            $status = if ($steps | Where-Object { $_.Status -eq 'WhatIf' }) { 'WhatIf' } else { 'Skipped' }
            $steps.Add((New-MspStepResult -Step 'Grant Exchange.ManageAsApp' -Status $status -Detail 'Needs the service principal first.'))
            foreach ($wanted in $roles) { $steps.Add((New-MspStepResult -Step "Assign role: $($wanted.DisplayName)" -Status $status -Detail 'Needs the service principal first.')) }
            return (& $finish)
        }
        $appSpId = [string]$appSp.id
        $exoSpId = [string]$exoSp.id
        $manageAsAppRoleId = $state.ManageAsAppRoleId

        # Exchange.ManageAsApp
        $hasRole = $state.HasManageAsApp
        if (-not $state.AppServicePrincipal) {
            try { $hasRole = (Get-MspExchangeAppAccessState -TenantId $tenant -AppId $AppId -Role $roles).HasManageAsApp } catch { $hasRole = $false }
        }
        if ($hasRole) {
            $steps.Add((New-MspStepResult -Step 'Grant Exchange.ManageAsApp' -Status Passed -Detail 'Already granted.'))
        }
        elseif ($PSCmdlet.ShouldProcess("customer $tenant", "Grant Exchange.ManageAsApp to automation app $AppId")) {
            try {
                $null = Invoke-MspGraphCall -TenantId $tenant -Method POST -Path "servicePrincipals/$exoSpId/appRoleAssignedTo" -Body @{ principalId = $appSpId; resourceId = $exoSpId; appRoleId = $manageAsAppRoleId }
                $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
                    @(Invoke-MspGraphCall -TenantId $tenant -Path "servicePrincipals/$appSpId/appRoleAssignments" | Where-Object { $_.resourceId -eq $exoSpId -and $_.appRoleId -eq $manageAsAppRoleId }).Count -gt 0
                }
                if ($check.Satisfied) { $steps.Add((New-MspStepResult -Step 'Grant Exchange.ManageAsApp' -Status Changed -Detail 'Granted and confirmed.')) }
                else { $steps.Add((New-MspStepResult -Step 'Grant Exchange.ManageAsApp' -Status Failed -Detail "The grant was sent but is not visible on readback. $($check.LastError)")) }
            }
            catch {
                $steps.Add((New-MspStepResult -Step 'Grant Exchange.ManageAsApp' -Status Failed -Detail "$($_.Exception.Message) You need Cloud Application Administrator (or Application Administrator) in this customer through GDAP, and AppRoleAssignment.ReadWrite.All on the partner app."))
            }
        }
        else {
            $steps.Add((New-MspStepResult -Step 'Grant Exchange.ManageAsApp' -Status WhatIf -Detail 'Would grant the app role.'))
        }

        # Directory roles (unified RBAC)
        $currentRoles = if ($state.AppServicePrincipal) { $state.Roles } else { @() }
        foreach ($wanted in $roles) {
            $name = "Assign role: $($wanted.DisplayName)"
            $existing = @($currentRoles | Where-Object { $_.RoleTemplateId -eq $wanted.RoleTemplateId -and $_.Assigned })
            if ($existing.Count -gt 0) {
                $steps.Add((New-MspStepResult -Step $name -Status Passed -Detail 'Already assigned at directory scope.'))
                continue
            }
            if (-not $PSCmdlet.ShouldProcess("customer $tenant", "Assign $($wanted.DisplayName) to automation app service principal $appSpId")) {
                $steps.Add((New-MspStepResult -Step $name -Status WhatIf -Detail 'Would assign the role.'))
                continue
            }
            $roleTemplateId = $wanted.RoleTemplateId
            try {
                $null = Invoke-MspGraphCall -TenantId $tenant -Method POST -Path 'roleManagement/directory/roleAssignments' -Body @{
                    '@odata.type'    = '#microsoft.graph.unifiedRoleAssignment'
                    principalId      = $appSpId
                    roleDefinitionId = $roleTemplateId
                    directoryScopeId = '/'
                }
                $check = Wait-MspCondition -TimeoutSeconds $ReadbackTimeoutSeconds -Condition {
                    $path = "roleManagement/directory/roleAssignments?`$filter=principalId eq '{0}' and roleDefinitionId eq '{1}'" -f $appSpId, $roleTemplateId
                    @(Invoke-MspGraphCall -TenantId $tenant -Path $path).Count -gt 0
                }
                if ($check.Satisfied) { $steps.Add((New-MspStepResult -Step $name -Status Changed -Detail 'Assigned and confirmed.')) }
                else { $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "The assignment was sent but is not visible on readback. $($check.LastError)")) }
            }
            catch {
                $steps.Add((New-MspStepResult -Step $name -Status Failed -Detail "$($_.Exception.Message) Assigning directory roles needs Privileged Role Administrator in this customer through GDAP, and RoleManagement.ReadWrite.Directory on the partner app."))
            }
        }

        try { $info.InitialDomain = (Get-MspInitialDomain -TenantId $tenant).InitialDomain } catch { $info.InitialDomain = $null }
        & $finish
    }
}
