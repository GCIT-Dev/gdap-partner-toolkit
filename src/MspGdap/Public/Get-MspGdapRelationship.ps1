function Get-MspGdapRelationship {
    <#
    .SYNOPSIS
        Lists GDAP relationships from the partner tenant, with review-friendly properties.
    .DESCRIPTION
        Reads tenantRelationships/delegatedAdminRelationships in the partner tenant. Terminated relationships are
        left out unless -IncludeTerminated or -Status asks for them.
        Each result shows role names, ContainsGlobalAdministrator, ContainsPrivilegedRoles (Global Administrator,
        Privileged Role Administrator or Privileged Authentication Administrator, as in Microsoft's Default GDAP),
        DaysRemaining, and ApprovalExpiresSoon for requests waiting 75 days or more (requests expire at 90 days).
    .PARAMETER TenantId
        Only relationships with this customer (tenant ID or verified domain). Alias: CustomerTenantId.
    .PARAMETER Status
        Only these statuses, for example active or approvalPending.
    .PARAMETER IncludeTerminated
        Include terminated relationships when no -Status is given.
    .PARAMETER RelationshipId
        One relationship by ID.
    .PARAMETER ApiVersion
        Microsoft Graph version, v1.0 (default) or beta.
    .EXAMPLE
        Get-MspGdapRelationship -Status active | Where-Object ContainsPrivilegedRoles
    .EXAMPLE
        Get-MspGdapRelationship -TenantId 'fabrikam.onmicrosoft.com'
    #>
    [CmdletBinding(DefaultParameterSetName = 'List')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(ParameterSetName = 'List', ValueFromPipelineByPropertyName)][Alias('CustomerTenantId')][string]$TenantId,
        [Parameter(ParameterSetName = 'List')][ValidateSet('activating', 'active', 'approvalPending', 'approved', 'created', 'expired', 'expiring', 'terminated', 'terminating', 'terminationRequested')][string[]]$Status,
        [Parameter(ParameterSetName = 'List')][switch]$IncludeTerminated,
        [Parameter(Mandatory, ParameterSetName = 'Id')][ValidateNotNullOrEmpty()][string]$RelationshipId,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0'
    )
    begin { $catalog = @(Get-MspGdapRoleCatalog) }
    process {
        if ($PSCmdlet.ParameterSetName -eq 'Id') {
            $item = Invoke-MspGraphCall -PartnerTenant -Path "tenantRelationships/delegatedAdminRelationships/$RelationshipId" -ApiVersion $ApiVersion
            if ($item) { ConvertTo-MspGdapRelationship -InputObject $item -Catalog $catalog }
            return
        }
        $customerId = $null
        if ($TenantId) { $customerId = ([string](Resolve-MspTenantId -TenantId $TenantId)).ToLowerInvariant() }
        $all = @(Invoke-MspGraphCall -PartnerTenant -Path 'tenantRelationships/delegatedAdminRelationships' -ApiVersion $ApiVersion)
        foreach ($item in $all) {
            if ($null -eq $item) { continue }
            $itemStatus = [string]$item.status
            if ($Status) { if ($Status -notcontains $itemStatus) { continue } }
            elseif (-not $IncludeTerminated -and $itemStatus -eq 'terminated') { continue }
            if ($customerId -and ([string]$item.customer.tenantId).ToLowerInvariant() -ne $customerId) { continue }
            ConvertTo-MspGdapRelationship -InputObject $item -Catalog $catalog
        }
    }
}
