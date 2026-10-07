function Set-MspGdapAccessAssignment {
    <#
    .SYNOPSIS
        Makes a GDAP relationship's security group assignments match your access map, one group at a time.
    .DESCRIPTION
        For each group in the access map:
          - no assignment yet: POST accessAssignments (starts as pending)
          - assignment with different roles: PATCH with the ETag (If-Match), as Microsoft requires
          - assignment with the same roles: nothing to do
        Each change is read back until it is active or the wait runs out. A group's roles must all be in the
        relationship, because roles cannot be added to a relationship after it is created.
        Microsoft allows 100 security groups per customer. The command refuses to go over that.
        With -RemoveUnlisted, assignments for groups that are not in the map are deleted.
        Returns one result per group.
    .PARAMETER RelationshipId
        The relationship to configure. It must be active (approved by the customer).
    .PARAMETER TenantId
        Instead of -RelationshipId: the customer, which must have exactly one active relationship.
    .PARAMETER AccessMapPath
        Your access map JSON (see Data/gdap-access.example.json).
    .PARAMETER WaitSeconds
        How long to wait for each assignment to become active. Default 120.
    .PARAMETER RemoveUnlisted
        Delete assignments for security groups that are not in the access map.
    .PARAMETER PollIntervalSeconds
        Seconds between readbacks while waiting. Default 10.
    .EXAMPLE
        Set-MspGdapAccessAssignment -RelationshipId $relationship.Id -AccessMapPath ./gdap-access.json -WhatIf
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Relationship')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Relationship', ValueFromPipelineByPropertyName)][Alias('Id')][ValidateNotNullOrEmpty()][string]$RelationshipId,
        [Parameter(Mandatory, ParameterSetName = 'Tenant')][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory)][string]$AccessMapPath,
        [switch]$RemoveUnlisted,
        [ValidateRange(0, 1800)][int]$WaitSeconds = 120,
        [ValidateRange(1, 120)][int]$PollIntervalSeconds = 10
    )
    process {
        $fail = {
            param($Step, $Detail, $Tenant)
            New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Set-MspGdapAccessAssignment' -TenantId $Tenant -Target $RelationshipId -Steps @(New-MspStepResult -Step $Step -Status Failed -Detail $Detail)
        }
        # Unknown role IDs are accepted here: every role must already be in the relationship, which
        # New-MspGdapRelationship only creates with unknown roles when -AllowUnknownRole was given.
        try { $map = @(Read-MspGdapAccessMap -Path $AccessMapPath -AllowUnknownRole) }
        catch { return (& $fail 'Load access map' $_.Exception.Message $TenantId) }

        if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
            try {
                $active = @(Get-MspGdapRelationship -TenantId $TenantId -Status active)
            }
            catch { return (& $fail 'Find relationship' $_.Exception.Message $TenantId) }
            if ($active.Count -ne 1) {
                return (& $fail 'Find relationship' ("Found {0} active relationships for {1}. Pass -RelationshipId." -f $active.Count, $TenantId) $TenantId)
            }
            $RelationshipId = $active[0].Id
        }

        $relationshipPath = "tenantRelationships/delegatedAdminRelationships/$RelationshipId"
        try { $relationship = Invoke-MspGraphCall -PartnerTenant -Path $relationshipPath }
        catch { return (& $fail 'Read relationship' $_.Exception.Message $TenantId) }
        $customerTenant = ([string]$relationship.customer.tenantId).ToLowerInvariant()
        if ($relationship.status -ne 'active') {
            return (& $fail 'Read relationship' "Relationship '$($relationship.displayName)' is $($relationship.status). Access can only be assigned once the customer has approved it (active)." $customerTenant)
        }
        $relationshipRoleIds = @($relationship.accessDetails.unifiedRoles | ForEach-Object { ([string]$_.roleDefinitionId).ToLowerInvariant() })

        try {
            $existing = @(Invoke-MspGraphCall -PartnerTenant -Path "$relationshipPath/accessAssignments" | Where-Object { $_.status -notin @('deleted', 'deleting') })
        }
        catch { return (& $fail 'Read access assignments' $_.Exception.Message $customerTenant) }

        $existingByGroup = @{}
        foreach ($assignment in $existing) { $existingByGroup[([string]$assignment.accessContainer.accessContainerId).ToLowerInvariant()] = $assignment }
        $groupCount = $existingByGroup.Count
        $mapGroupIds = @($map | ForEach-Object { $_.GroupId })

        foreach ($entry in $map) {
            $steps = New-Object System.Collections.Generic.List[object]
            $groupLabel = "$($entry.GroupDisplayName) ($($entry.GroupId))"
            $wantedIds = @($entry.Roles | ForEach-Object { $_.RoleTemplateId } | Sort-Object -Unique)
            $roleNames = ($entry.Roles | ForEach-Object { $_.DisplayName }) -join ', '
            $emit = { New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Set-MspGdapAccessAssignment' -TenantId $customerTenant -Target $groupLabel -Steps $steps.ToArray() -Property ([ordered]@{ RelationshipId = $RelationshipId; GroupId = $entry.GroupId; Roles = @($entry.Roles | ForEach-Object { $_.DisplayName }) }) }

            $notInRelationship = @($entry.Roles | Where-Object { $relationshipRoleIds -notcontains $_.RoleTemplateId })
            if ($notInRelationship.Count -gt 0) {
                $steps.Add((New-MspStepResult -Step 'Check roles against relationship' -Status Failed -Detail ("Not in relationship '{0}': {1}. Roles cannot be added to an existing relationship. Create a new one with New-MspGdapRelationship." -f $relationship.displayName, (($notInRelationship | ForEach-Object { $_.DisplayName }) -join ', '))))
                & $emit; continue
            }

            $current = if ($existingByGroup.ContainsKey($entry.GroupId)) { $existingByGroup[$entry.GroupId] } else { $null }
            $assignmentId = $null
            if ($current) {
                $currentIds = @($current.accessDetails.unifiedRoles | ForEach-Object { ([string]$_.roleDefinitionId).ToLowerInvariant() } | Sort-Object -Unique)
                if (@(Compare-Object -ReferenceObject $currentIds -DifferenceObject $wantedIds).Count -eq 0) {
                    $state = if ($current.status -eq 'active') { 'Passed' } else { 'Warning' }
                    $steps.Add((New-MspStepResult -Step 'Assign roles' -Status $state -Detail "Already assigned ($($current.status)): $roleNames"))
                    if ($current.status -eq 'error') { $steps.Add((New-MspStepResult -Step 'Assignment status' -Status Failed -Detail 'The existing assignment is in error. Remove it in Partner Center and run again.')) }
                    & $emit; continue
                }
                if (-not $PSCmdlet.ShouldProcess("relationship $RelationshipId, group $groupLabel", "Change roles to: $roleNames")) {
                    $steps.Add((New-MspStepResult -Step 'Assign roles' -Status WhatIf -Detail "Would change roles to: $roleNames")); & $emit; continue
                }
                $etag = [string]$current.'@odata.etag'
                if (-not $etag) {
                    try { $etag = [string](Invoke-MspGraphCall -PartnerTenant -Path "$relationshipPath/accessAssignments/$($current.id)").'@odata.etag' } catch { $etag = $null }
                }
                if (-not $etag) { $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail 'Could not read the assignment ETag needed for the update.')); & $emit; continue }
                $response = Invoke-MspGraphDirect -PartnerTenant -Method PATCH -Path "$relationshipPath/accessAssignments/$($current.id)" -Headers @{ 'If-Match' = $etag } -Body @{ accessDetails = @{ unifiedRoles = @($wantedIds | ForEach-Object { @{ roleDefinitionId = $_ } }) } }
                if ($response.StatusCode -notin 200, 202) {
                    $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail "Update returned $($response.StatusCode). $($response.Error)")); & $emit; continue
                }
                $assignmentId = [string]$current.id
            }
            else {
                if ($groupCount -ge 100) {
                    $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail 'This customer already has 100 security group assignments, the Microsoft limit.')); & $emit; continue
                }
                if (-not $PSCmdlet.ShouldProcess("relationship $RelationshipId, group $groupLabel", "Assign roles: $roleNames")) {
                    $steps.Add((New-MspStepResult -Step 'Assign roles' -Status WhatIf -Detail "Would assign: $roleNames")); & $emit; continue
                }
                $body = @{
                    accessContainer = [ordered]@{ accessContainerId = $entry.GroupId; accessContainerType = 'securityGroup' }
                    accessDetails   = @{ unifiedRoles = @($wantedIds | ForEach-Object { @{ roleDefinitionId = $_ } }) }
                }
                try {
                    $created = Invoke-MspGraphCall -PartnerTenant -Method POST -Path "$relationshipPath/accessAssignments" -Body $body
                    $assignmentId = [string]$created.id
                    $groupCount++
                }
                catch {
                    $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail $_.Exception.Message)); & $emit; continue
                }
                if (-not $assignmentId) { $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail 'Graph did not return the new assignment.')); & $emit; continue }
            }

            $seen = @{ Status = 'unknown'; RolesMatch = $false }
            $check = Wait-MspCondition -TimeoutSeconds $WaitSeconds -IntervalSeconds $PollIntervalSeconds -Condition {
                $read = Invoke-MspGraphCall -PartnerTenant -Path "$relationshipPath/accessAssignments/$assignmentId"
                $readIds = @($read.accessDetails.unifiedRoles | ForEach-Object { ([string]$_.roleDefinitionId).ToLowerInvariant() } | Sort-Object -Unique)
                $seen.Status = [string]$read.status
                $seen.RolesMatch = (@(Compare-Object -ReferenceObject $readIds -DifferenceObject $wantedIds).Count -eq 0)
                if ($seen.RolesMatch -and ($read.status -in @('active', 'error'))) { $read }
            }
            if ($seen.Status -eq 'error') {
                $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail 'Microsoft reports the assignment in error. Check the group exists in the partner tenant and is a security group.'))
            }
            elseif ($check.Satisfied) {
                $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Changed -Detail "Active and confirmed: $roleNames"))
            }
            elseif ($seen.RolesMatch) {
                $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Unknown -Detail "Saved with the right roles but still $($seen.Status) after $WaitSeconds s. Run Test-MspGdapAccess later to confirm it is active."))
            }
            else {
                $steps.Add((New-MspStepResult -Step 'Assign roles' -Status Failed -Detail "The readback does not show the requested roles (status $($seen.Status)). $($check.LastError)"))
            }
            & $emit
        }

        if ($RemoveUnlisted) {
            foreach ($assignment in $existing) {
                $groupId = ([string]$assignment.accessContainer.accessContainerId).ToLowerInvariant()
                if ($mapGroupIds -contains $groupId) { continue }
                $steps = New-Object System.Collections.Generic.List[object]
                if ($PSCmdlet.ShouldProcess("relationship $RelationshipId, group $groupId", 'Remove access assignment (group not in access map)')) {
                    $etag = [string]$assignment.'@odata.etag'
                    $response = Invoke-MspGraphDirect -PartnerTenant -Method DELETE -Path "$relationshipPath/accessAssignments/$($assignment.id)" -Headers @{ 'If-Match' = $etag }
                    if ($response.StatusCode -in 200, 202, 204) {
                        $gone = Wait-MspCondition -TimeoutSeconds $WaitSeconds -IntervalSeconds $PollIntervalSeconds -Condition {
                            $read = Invoke-MspGraphDirect -PartnerTenant -Method GET -Path "$relationshipPath/accessAssignments/$($assignment.id)"
                            ($read.StatusCode -eq 404) -or ($read.Body -and $read.Body.status -in @('deleted', 'deleting'))
                        }
                        if ($gone.Satisfied) { $steps.Add((New-MspStepResult -Step 'Remove assignment' -Status Changed -Detail 'Removed and confirmed.')) }
                        else { $steps.Add((New-MspStepResult -Step 'Remove assignment' -Status Unknown -Detail 'Delete accepted but not yet confirmed.')) }
                    }
                    else {
                        $steps.Add((New-MspStepResult -Step 'Remove assignment' -Status Failed -Detail "Delete returned $($response.StatusCode). $($response.Error)"))
                    }
                }
                else {
                    $steps.Add((New-MspStepResult -Step 'Remove assignment' -Status WhatIf -Detail 'Would remove this group (not in the access map).'))
                }
                New-MspOperationResult -Cmdlet $PSCmdlet -Operation 'Set-MspGdapAccessAssignment' -TenantId $customerTenant -Target $groupId -Steps $steps.ToArray() -Property ([ordered]@{ RelationshipId = $RelationshipId; GroupId = $groupId; Roles = @() })
            }
        }
    }
}
