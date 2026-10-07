function New-MspGdapRelationship {
    <#
    .SYNOPSIS
        Creates a GDAP relationship request for one customer with a least-privilege role set, and locks it for approval.
    .DESCRIPTION
        Role set, in order of precedence:
          -AccessMapPath  the union of every role in your access map (the same file Set-MspGdapAccessAssignment uses)
          -Role           role names or template IDs
          default         roles marked default in Data/gdap-roles.leastprivilege.json (no Global Administrator)
        Microsoft rules this command enforces or reports:
          - duration P1D to P2Y (default P730D)
          - auto-extend is P180D or off (PT0S), and never applies when Global Administrator is included
          - roles cannot be added after creation, so check the role list in the -WhatIf output
          - highly privileged roles (Privileged Role Administrator, Privileged Authentication Administrator)
            are allowed but reported as a Warning step, and role template IDs that are not in the role
            catalogue are refused unless -AllowUnknownRole is given
          - display name unique across all your relationships, 50 characters or fewer
          - unapproved requests expire after 90 days
        After creation the relationship is locked for approval (lockForApproval) and read back. The customer
        approves it from the link you copy out of Partner Center. MspGdap does not build that link because
        Microsoft does not document its format.
    .PARAMETER TenantId
        Customer tenant ID or verified domain. Alias: CustomerTenantId. Your partner tenant is refused.
    .PARAMETER CustomerDisplayName
        Optional customer name sent with the request.
    .PARAMETER DisplayName
        Relationship name. Default: <Prefix>-<yyyyMMdd>-<first 8 characters of the tenant ID>.
    .PARAMETER Prefix
        Prefix for the default display name. Default MspGdap.
    .PARAMETER Duration
        ISO 8601 duration from P1D to P730D (two years). Default P730D.
    .PARAMETER AutoExtend
        Request auto-extend (P180D). Ignored when Global Administrator is in the role set.
    .PARAMETER AccessMapPath
        Access map JSON. The role set is the union of every role in it.
    .PARAMETER Role
        Role names or template IDs.
    .PARAMETER IncludeGlobalAdministrator
        Allow Global Administrator in the role set. Auto-extend is then forced off.
    .PARAMETER AllowUnknownRole
        Accept role template IDs that are not in Data/gdap-roles.leastprivilege.json. Their privilege
        level cannot be checked, so review them yourself first.
    .PARAMETER SkipLock
        Leave the relationship in the created state (not sent for approval).
    .PARAMETER LockTimeoutSeconds
        How long to wait for the relationship to show approvalPending after the lock request. Default 60.
    .EXAMPLE
        New-MspGdapRelationship -TenantId 'fabrikam.onmicrosoft.com' -AccessMapPath ./gdap-access.json -AutoExtend -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Default')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [string]$CustomerDisplayName,
        [ValidateLength(1, 50)][string]$DisplayName,
        [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9 ._-]{0,31}$')][string]$Prefix = 'MspGdap',
        [string]$Duration = 'P730D',
        [switch]$AutoExtend,
        [Parameter(Mandatory, ParameterSetName = 'AccessMap')][string]$AccessMapPath,
        [Parameter(Mandatory, ParameterSetName = 'Role')][string[]]$Role,
        [switch]$IncludeGlobalAdministrator,
        [switch]$AllowUnknownRole,
        [switch]$SkipLock,
        [ValidateRange(0, 600)][int]$LockTimeoutSeconds = 60
    )
    $steps = New-Object System.Collections.Generic.List[object]
    $info = [ordered]@{ RelationshipId = $null; DisplayName = $null; Status = $null; Roles = @(); AutoExtendDuration = $null; Duration = $Duration; Note = $null }
    $tenant = $TenantId
    $finish = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'New-MspGdapRelationship' -TenantId $tenant -Target $info.DisplayName -Steps $steps.ToArray() -Property $info }

    try { $tenant = Resolve-MspCustomerTenant -TenantId $TenantId }
    catch {
        $steps.Add((New-MspStepResult -Step 'Resolve tenant' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }

    # Duration
    try {
        $span = [System.Xml.XmlConvert]::ToTimeSpan($Duration)
        if ($span.TotalDays -lt 1 -or $span.TotalDays -gt 730) { throw "Duration $Duration is outside P1D to P730D." }
        $steps.Add((New-MspStepResult -Step 'Check duration' -Status Passed -Detail "$Duration ($([int]$span.TotalDays) days)."))
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Check duration' -Status Failed -Detail "Duration must be ISO 8601 between P1D and P730D, for example P730D. $($_.Exception.Message)"))
        return (& $finish)
    }

    # Roles
    try {
        $catalog = @(Get-MspGdapRoleCatalog)
        $roles = switch ($PSCmdlet.ParameterSetName) {
            'AccessMap' { @(Resolve-MspGdapRole -Role @(Read-MspGdapAccessMap -Path $AccessMapPath -AllowUnknownRole:$AllowUnknownRole | ForEach-Object { $_.Roles }) -Catalog $catalog -AllowUnknown:$AllowUnknownRole) }
            'Role' { @(Resolve-MspGdapRole -Role $Role -Catalog $catalog -AllowUnknown:$AllowUnknownRole) }
            default { @($catalog | Where-Object { $_.Default }) }
        }
        if ($roles.Count -eq 0) { throw 'The role set is empty.' }
        $globalAdministratorId = Get-MspWellKnownRoleId -Name GlobalAdministrator
        $hasGlobalAdmin = @($roles | Where-Object { $_.RoleTemplateId -eq $globalAdministratorId }).Count -gt 0
        if ($hasGlobalAdmin -and -not $IncludeGlobalAdministrator) {
            throw 'Global Administrator is in the role set. Remove it, or pass -IncludeGlobalAdministrator if you really need it (auto-extend is then off).'
        }
        $info.Roles = @($roles | ForEach-Object { $_.DisplayName })
        $steps.Add((New-MspStepResult -Step 'Resolve roles' -Status Passed -Detail ("{0} roles (final once created): {1}" -f $roles.Count, ($info.Roles -join ', ')) -Data $roles))
        $privileged = @($roles | Where-Object { $_.HighlyPrivileged -and $_.RoleTemplateId -ne $globalAdministratorId })
        if ($privileged.Count -gt 0) {
            $steps.Add((New-MspStepResult -Step 'Highly privileged roles' -Status Warning -Detail ("{0} can give themselves or others more access, close to Global Administrator. Keep them in a separate, short relationship and use PIM for groups if you can." -f (($privileged | ForEach-Object { $_.DisplayName }) -join ', '))))
        }
        $unknown = @($roles | Where-Object { $_.PSObject.Properties['Unknown'] -and $_.Unknown })
        if ($unknown.Count -gt 0) {
            $steps.Add((New-MspStepResult -Step 'Roles not in catalogue' -Status Warning -Detail ("Accepted with -AllowUnknownRole, privilege level not checked: {0}" -f (($unknown | ForEach-Object { $_.RoleTemplateId }) -join ', '))))
        }
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Resolve roles' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }

    $autoExtendDuration = if ($AutoExtend) { 'P180D' } else { 'PT0S' }
    if ($AutoExtend -and $hasGlobalAdmin) {
        $autoExtendDuration = 'PT0S'
        $steps.Add((New-MspStepResult -Step 'Auto-extend' -Status Warning -Detail 'Microsoft does not auto-extend relationships that contain Global Administrator. Auto-extend set to off (PT0S).'))
    }
    $info.AutoExtendDuration = $autoExtendDuration

    # Display name
    if (-not $DisplayName) { $DisplayName = '{0}-{1:yyyyMMdd}-{2}' -f $Prefix, [datetime]::UtcNow, $tenant.Substring(0, 8) }
    $info.DisplayName = $DisplayName
    try {
        $existing = @(Invoke-MspGraphCall -PartnerTenant -Path 'tenantRelationships/delegatedAdminRelationships')
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Check existing relationships' -Status Unknown -Detail "Could not list existing relationships. $($_.Exception.Message)"))
        return (& $finish)
    }
    if (@($existing | Where-Object { $_.displayName -ieq $DisplayName }).Count -gt 0) {
        $steps.Add((New-MspStepResult -Step 'Check display name' -Status Failed -Detail "A relationship named '$DisplayName' already exists (display names must be unique across all your relationships). Choose another -DisplayName."))
        return (& $finish)
    }
    $steps.Add((New-MspStepResult -Step 'Check display name' -Status Passed -Detail "'$DisplayName' is unused."))
    $sameCustomer = @($existing | Where-Object { ([string]$_.customer.tenantId) -ieq $tenant -and $_.status -in @('active', 'approvalPending', 'created', 'activating', 'approved') })
    if ($sameCustomer.Count -gt 0) {
        $steps.Add((New-MspStepResult -Step 'Check existing relationships' -Status Warning -Detail ("This customer already has: {0}" -f (($sameCustomer | ForEach-Object { "$($_.displayName) ($($_.status))" }) -join ', '))))
    }

    $body = [ordered]@{
        displayName        = $DisplayName
        duration           = $Duration
        customer           = [ordered]@{ tenantId = $tenant }
        accessDetails      = @{ unifiedRoles = @($roles | ForEach-Object { @{ roleDefinitionId = $_.RoleTemplateId } }) }
        autoExtendDuration = $autoExtendDuration
    }
    if ($CustomerDisplayName) { $body.customer['displayName'] = $CustomerDisplayName }

    if (-not $PSCmdlet.ShouldProcess("customer $tenant", "Create GDAP relationship '$DisplayName' ($Duration, auto-extend $autoExtendDuration) with $($roles.Count) roles: $($info.Roles -join ', ')")) {
        $steps.Add((New-MspStepResult -Step 'Create relationship' -Status WhatIf -Detail "Would create '$DisplayName' with: $($info.Roles -join ', ')"))
        if (-not $SkipLock) { $steps.Add((New-MspStepResult -Step 'Lock for approval' -Status WhatIf -Detail 'Runs after creation.')) }
        return (& $finish)
    }

    try {
        $created = Invoke-MspGraphCall -PartnerTenant -Method POST -Path 'tenantRelationships/delegatedAdminRelationships' -Body $body
        if (-not $created -or -not $created.id) { throw 'Graph did not return the new relationship.' }
        $info.RelationshipId = [string]$created.id
    }
    catch {
        $steps.Add((New-MspStepResult -Step 'Create relationship' -Status Failed -Detail $_.Exception.Message))
        return (& $finish)
    }

    $relationshipPath = "tenantRelationships/delegatedAdminRelationships/$($info.RelationshipId)"
    $wantedIds = @($roles | ForEach-Object { $_.RoleTemplateId })
    $check = Wait-MspCondition -TimeoutSeconds 30 -IntervalSeconds 5 -Condition {
        $read = Invoke-MspGraphCall -PartnerTenant -Path $relationshipPath
        $readIds = @($read.accessDetails.unifiedRoles | ForEach-Object { ([string]$_.roleDefinitionId).ToLowerInvariant() })
        if (@($wantedIds | Where-Object { $readIds -notcontains $_ }).Count -eq 0) { $read }
    }
    if (-not $check.Satisfied) {
        $steps.Add((New-MspStepResult -Step 'Create relationship' -Status Failed -Detail "Created $($info.RelationshipId) but the readback does not show every role. $($check.LastError)"))
        return (& $finish)
    }
    $info.Status = [string]$check.Value.status
    $steps.Add((New-MspStepResult -Step 'Create relationship' -Status Changed -Detail "Created $($info.RelationshipId) (status $($info.Status)) and confirmed its roles."))

    if ($SkipLock) {
        $steps.Add((New-MspStepResult -Step 'Lock for approval' -Status Skipped -Detail 'Left in the created state (-SkipLock). Lock it before the customer can approve.'))
        $info.Note = 'Not yet sent for approval.'
        return (& $finish)
    }

    $request = $null
    $lockError = $null
    for ($try = 1; $try -le 3 -and -not $request; $try++) {
        try { $request = Invoke-MspGraphCall -PartnerTenant -Method POST -Path "$relationshipPath/requests" -Body @{ action = 'lockForApproval' } }
        catch { $lockError = $_.Exception.Message; Start-Sleep -Seconds 5 }
    }
    if (-not $request) {
        $steps.Add((New-MspStepResult -Step 'Lock for approval' -Status Failed -Detail "The relationship exists (status created) but could not be locked. Lock it in Partner Center, or terminate it and run this command again. $lockError"))
        return (& $finish)
    }
    $seen = @{ Status = 'unknown' }
    $lockCheck = Wait-MspCondition -TimeoutSeconds $LockTimeoutSeconds -IntervalSeconds 5 -Condition {
        $read = Invoke-MspGraphCall -PartnerTenant -Path $relationshipPath
        if ($read) { $seen.Status = [string]$read.status }
        if ($read.status -eq 'approvalPending') { $read }
    }
    if ($lockCheck.Satisfied) {
        $info.Status = 'approvalPending'
        $steps.Add((New-MspStepResult -Step 'Lock for approval' -Status Changed -Detail 'Locked and confirmed: status approvalPending.'))
        $info.Note = 'Copy the approval link for this relationship from Partner Center (Customers, select the customer, Admin relationships) and send it to the customer. Requests expire after 90 days.'
    }
    else {
        $info.Status = $seen.Status
        $steps.Add((New-MspStepResult -Step 'Lock for approval' -Status Unknown -Detail "Lock request sent (request status $($request.status)) but the relationship is not approvalPending yet (last status $($seen.Status)). Check it with Get-MspGdapRelationship -RelationshipId $($info.RelationshipId)."))
    }
    & $finish
}
