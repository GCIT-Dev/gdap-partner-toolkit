function Test-MspGdapAccess {
    <#
    .SYNOPSIS
        Shows which GDAP roles the signed-in technician effectively holds in a customer, and optionally checks for required roles.
    .DESCRIPTION
        Read-only. Works it out from the partner side, which is what Microsoft enforces:
          1. active GDAP relationships with the customer
          2. their active access assignments (security group to roles)
          3. the technician's transitive group memberships in the partner tenant (/me/transitiveMemberOf)
        Effective roles = roles of the assignments whose group the technician belongs to.
        It then calls the customer tenant with the technician's token (GET /organization) to prove the token
        works there and reaches the right tenant.
        It does not read roles from the customer side. A GDAP technician is a partner user, not an object in
        the customer directory, so GET /me/memberOf in the customer fails (HTTP 400, "Current authenticated
        context is not valid"). The 'Customer-side role view' step is reported as Skipped for that reason.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .PARAMETER RequiredRole
        Every one of these roles must be held (names or template IDs).
    .PARAMETER AnyRole
        At least one of these roles must be held, for example 'Cloud Application Administrator','Application Administrator'.
    .PARAMETER SkipCustomerProbe
        Do not call the customer tenant.
    .EXAMPLE
        Test-MspGdapAccess -TenantId 'fabrikam.onmicrosoft.com' -AnyRole 'Cloud Application Administrator','Application Administrator'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [string[]]$RequiredRole,
        [string[]]$AnyRole,
        [switch]$SkipCustomerProbe
    )
    begin {
        $catalog = @(Get-MspGdapRoleCatalog)
        $myGroups = $null
    }
    process {
        $steps = New-Object System.Collections.Generic.List[object]
        $info = [ordered]@{ CustomerName = $null; EffectiveRoles = @(); EffectiveRoleNames = @(); Relationships = @() }
        $tenant = $TenantId
        $finish = { New-MspOperationResult -Operation 'Test-MspGdapAccess' -TenantId $tenant -Target $tenant -Steps $steps.ToArray() -Property $info }

        try {
            $required = if ($RequiredRole) { @(Resolve-MspGdapRole -Role $RequiredRole -Catalog $catalog -AllowUnknown) } else { @() }
            $any = if ($AnyRole) { @(Resolve-MspGdapRole -Role $AnyRole -Catalog $catalog -AllowUnknown) } else { @() }
            $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Prepare' -Status Failed -Detail $_.Exception.Message))
            return (& $finish)
        }

        try { $relationships = @(Get-MspGdapRelationship -TenantId $tenant -Status active) }
        catch {
            $steps.Add((New-MspStepResult -Step 'Active GDAP relationships' -Status Unknown -Detail $_.Exception.Message))
            return (& $finish)
        }
        if ($relationships.Count -eq 0) {
            $steps.Add((New-MspStepResult -Step 'Active GDAP relationships' -Status Failed -Detail 'No active GDAP relationship with this customer.'))
            return (& $finish)
        }
        $info.CustomerName = $relationships[0].CustomerName
        $info.Relationships = @($relationships | ForEach-Object { $_.DisplayName })
        $steps.Add((New-MspStepResult -Step 'Active GDAP relationships' -Status Passed -Detail ($info.Relationships -join ', ')))

        try {
            if ($null -eq $myGroups) {
                $myGroups = @{}
                foreach ($group in @(Invoke-MspGraphCall -PartnerTenant -Path 'me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName')) {
                    if ($group) { $myGroups[([string]$group.id).ToLowerInvariant()] = [string]$group.displayName }
                }
            }
        }
        catch {
            $myGroups = $null
            $steps.Add((New-MspStepResult -Step 'Technician group membership' -Status Unknown -Detail "Could not read your group memberships in the partner tenant. $($_.Exception.Message)"))
            return (& $finish)
        }
        $steps.Add((New-MspStepResult -Step 'Technician group membership' -Status Passed -Detail "$($myGroups.Count) groups in the partner tenant."))

        $effective = @{}
        foreach ($relationship in $relationships) {
            try { $assignments = @(Invoke-MspGraphCall -PartnerTenant -Path "tenantRelationships/delegatedAdminRelationships/$($relationship.Id)/accessAssignments") }
            catch {
                $steps.Add((New-MspStepResult -Step "Access assignments: $($relationship.DisplayName)" -Status Unknown -Detail $_.Exception.Message))
                continue
            }
            foreach ($assignment in $assignments) {
                if ($assignment.status -ne 'active') { continue }
                $groupId = ([string]$assignment.accessContainer.accessContainerId).ToLowerInvariant()
                if (-not $myGroups.ContainsKey($groupId)) { continue }
                foreach ($roleRef in @($assignment.accessDetails.unifiedRoles)) {
                    $roleId = ([string]$roleRef.roleDefinitionId).ToLowerInvariant()
                    if (-not $effective.ContainsKey($roleId)) {
                        $entry = @($catalog | Where-Object { $_.RoleTemplateId -eq $roleId }) | Select-Object -First 1
                        $effective[$roleId] = [pscustomobject]@{
                            DisplayName    = if ($entry) { $entry.DisplayName } else { "Unknown role $roleId" }
                            RoleTemplateId = $roleId
                            ViaGroups      = New-Object System.Collections.Generic.List[string]
                            Relationships  = New-Object System.Collections.Generic.List[string]
                        }
                    }
                    if (-not $effective[$roleId].ViaGroups.Contains($myGroups[$groupId])) { $effective[$roleId].ViaGroups.Add($myGroups[$groupId]) }
                    if (-not $effective[$roleId].Relationships.Contains($relationship.DisplayName)) { $effective[$roleId].Relationships.Add($relationship.DisplayName) }
                }
            }
        }
        $info.EffectiveRoles = @($effective.Values | Sort-Object -Property DisplayName)
        $info.EffectiveRoleNames = @($info.EffectiveRoles | ForEach-Object { $_.DisplayName })
        if ($info.EffectiveRoles.Count -eq 0) {
            $steps.Add((New-MspStepResult -Step 'Effective roles' -Status Failed -Detail 'You are not in any group with an active access assignment for this customer.'))
        }
        else {
            $steps.Add((New-MspStepResult -Step 'Effective roles' -Status Passed -Detail ($info.EffectiveRoleNames -join ', ') -Data $info.EffectiveRoles))
        }

        if ($required.Count -gt 0) {
            $missing = @($required | Where-Object { -not $effective.ContainsKey($_.RoleTemplateId) })
            if ($missing.Count -gt 0) { $steps.Add((New-MspStepResult -Step 'Required roles' -Status Failed -Detail ("Missing: {0}" -f (($missing | ForEach-Object { $_.DisplayName }) -join ', ')))) }
            else { $steps.Add((New-MspStepResult -Step 'Required roles' -Status Passed -Detail 'All required roles held.')) }
        }
        if ($any.Count -gt 0) {
            $held = @($any | Where-Object { $effective.ContainsKey($_.RoleTemplateId) })
            if ($held.Count -eq 0) { $steps.Add((New-MspStepResult -Step 'Any of roles' -Status Failed -Detail ("None held of: {0}" -f (($any | ForEach-Object { $_.DisplayName }) -join ', ')))) }
            else { $steps.Add((New-MspStepResult -Step 'Any of roles' -Status Passed -Detail ("Held: {0}" -f (($held | ForEach-Object { $_.DisplayName }) -join ', ')))) }
        }

        if (-not $SkipCustomerProbe) {
            try {
                $org = @(Invoke-MspGraphCall -TenantId $tenant -Path 'organization?$select=id,displayName') | Select-Object -First 1
                if ($org -and ([string]$org.id).ToLowerInvariant() -eq $tenant) {
                    $steps.Add((New-MspStepResult -Step 'Customer token' -Status Passed -Detail "Token works in $($org.displayName)."))
                }
                else {
                    $steps.Add((New-MspStepResult -Step 'Customer token' -Status Failed -Detail "The customer call returned organisation $($org.id), not $tenant."))
                }
            }
            catch {
                $steps.Add((New-MspStepResult -Step 'Customer token' -Status Failed -Detail "Could not call the customer tenant: $($_.Exception.Message)"))
            }
            # No /me/memberOf probe: GDAP technicians are not objects in the customer directory, so Graph returns
            # HTTP 400 there. The roles come from the partner-side access assignments worked out above.
            $steps.Add((New-MspStepResult -Step 'Customer-side role view' -Status Skipped -Detail "Not checked. GDAP technicians are not objects in the customer directory, so the customer cannot list their roles (/me/memberOf). GDAP roles come from the partner-side group assignments shown in 'Effective roles'."))
        }
        & $finish
    }
}
