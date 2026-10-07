function ConvertTo-MspGdapRelationship {
    <#
    .SYNOPSIS
        Adds readable, review-friendly properties to a Graph delegatedAdminRelationship.
    .PARAMETER InputObject
        A delegatedAdminRelationship from Microsoft Graph.
    .PARAMETER Catalog
        Role catalogue entries. Loaded when omitted.
    .PARAMETER Now
        Reference time for day counts (tests).
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipeline)][object]$InputObject,
        [object[]]$Catalog,
        [datetime]$Now = [datetime]::UtcNow
    )
    begin { if (-not $Catalog) { $Catalog = @(Get-MspGdapRoleCatalog) } }
    process {
        $roleIds = @($InputObject.accessDetails.unifiedRoles | ForEach-Object { ([string]$_.roleDefinitionId).ToLowerInvariant() })
        $roles = foreach ($id in $roleIds) {
            $entry = @($Catalog | Where-Object { $_.RoleTemplateId -eq $id }) | Select-Object -First 1
            [pscustomobject]@{
                DisplayName      = if ($entry) { $entry.DisplayName } else { "Unknown role $id" }
                RoleTemplateId   = $id
                HighlyPrivileged = if ($entry) { $entry.HighlyPrivileged } else { $false }
            }
        }
        $endDate = $null
        if ($InputObject.PSObject.Properties['endDateTime'] -and $InputObject.endDateTime) { $endDate = ([datetime]$InputObject.endDateTime).ToUniversalTime() }
        $created = $null
        if ($InputObject.PSObject.Properties['createdDateTime'] -and $InputObject.createdDateTime) { $created = ([datetime]$InputObject.createdDateTime).ToUniversalTime() }
        $status = [string]$InputObject.status
        $requestAge = if ($status -eq 'approvalPending' -and $created) { [int][Math]::Floor(($Now - $created).TotalDays) } else { $null }

        [pscustomobject]@{
            PSTypeName                  = 'MspGdap.GdapRelationship'
            Id                          = [string]$InputObject.id
            DisplayName                 = [string]$InputObject.displayName
            Status                      = $status
            CustomerTenantId            = ([string]$InputObject.customer.tenantId).ToLowerInvariant()
            CustomerName                = [string]$InputObject.customer.displayName
            Duration                    = [string]$InputObject.duration
            AutoExtendDuration          = [string]$InputObject.autoExtendDuration
            CreatedDateTime             = $created
            ActivatedDateTime           = if ($InputObject.PSObject.Properties['activatedDateTime'] -and $InputObject.activatedDateTime) { ([datetime]$InputObject.activatedDateTime).ToUniversalTime() } else { $null }
            EndDateTime                 = $endDate
            DaysRemaining               = if ($endDate) { [int][Math]::Floor(($endDate - $Now).TotalDays) } else { $null }
            Roles                       = @($roles)
            RoleNames                   = @($roles | ForEach-Object { $_.DisplayName })
            ContainsGlobalAdministrator = ($roleIds -contains (Get-MspWellKnownRoleId -Name GlobalAdministrator))
            ContainsPrivilegedRoles     = (@($roles | Where-Object { $_.HighlyPrivileged }).Count -gt 0)
            ApprovalRequestAgeDays      = $requestAge
            ApprovalExpiresSoon         = ($null -ne $requestAge -and $requestAge -ge 75)
            Raw                         = $InputObject
        }
    }
}
