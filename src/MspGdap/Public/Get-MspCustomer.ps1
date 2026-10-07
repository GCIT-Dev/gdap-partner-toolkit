function Get-MspCustomer {
    <#
    .SYNOPSIS
        Lists your customers from the partner tenant, optionally with GDAP status.
    .DESCRIPTION
        Reads Microsoft Graph /contracts in your partner tenant (customers with
        a reseller or legacy contract). With -IncludeGdapStatus it also reads
        /tenantRelationships/delegatedAdminRelationships and adds the GDAP
        picture for each customer: active relationship count, latest end date,
        auto-extend, pending approvals and whether any relationship includes
        Global Administrator. Customers that only have GDAP (no reseller
        relationship) are included in that mode too, because GDAP does not
        need a reseller relationship.

        Reads only. Runs against the partner tenant, using the technician's
        delegated token.
    .PARAMETER Name
        Filter by display name or default domain (wildcards allowed).
    .PARAMETER TenantId
        Return only this customer (tenant GUID or verified domain).
    .PARAMETER IncludeGdapStatus
        Add GDAP relationship details (needs DelegatedAdminRelationship.Read.All).
    .PARAMETER UserPrincipalName
        Technician whose refresh token is used.
    .EXAMPLE
        Get-MspCustomer
    .EXAMPLE
        Get-MspCustomer -Name 'contoso*' -IncludeGdapStatus
    .EXAMPLE
        Get-MspCustomer -IncludeGdapStatus | Where-Object GdapStatus -ne 'active'
    #>
    [CmdletBinding()]
    [OutputType('MspGdap.Customer')]
    param(
        [SupportsWildcards()]
        [string]$Name,

        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [switch]$IncludeGdapStatus,

        [string]$UserPrincipalName
    )

    $common = @{ PartnerTenant = $true }
    if ($UserPrincipalName) { $common.UserPrincipalName = $UserPrincipalName }

    try {
        $globalAdministratorTemplateId = Get-MspWellKnownRoleId -Name GlobalAdministrator
        $tenantFilter = if ($TenantId) { Resolve-MspTenantId -Tenant $TenantId } else { $null }
        $customers = [ordered]@{}
        $contracts = @(Invoke-MspGraphRequest @common -Uri 'contracts?$select=customerId,defaultDomainName,displayName')
        foreach ($contract in $contracts) {
            if (-not $contract.customerId) { continue }
            $id = ([string]$contract.customerId).ToLowerInvariant()
            $customers[$id] = [ordered]@{
                TenantId          = $id
                DisplayName       = $contract.displayName
                DefaultDomainName = $contract.defaultDomainName
                HasContract       = $true
                Relationships     = [System.Collections.Generic.List[object]]::new()
            }
        }

        if ($IncludeGdapStatus) {
            $relationships = @(Invoke-MspGraphRequest @common -Uri 'tenantRelationships/delegatedAdminRelationships')
            foreach ($relationship in $relationships) {
                $customerTenant = $null
                if ($relationship.customer -and $relationship.customer.tenantId) { $customerTenant = ([string]$relationship.customer.tenantId).ToLowerInvariant() }
                if (-not $customerTenant) { continue }
                if (-not $customers.Contains($customerTenant)) {
                    $customers[$customerTenant] = [ordered]@{
                        TenantId          = $customerTenant
                        DisplayName       = $relationship.customer.displayName
                        DefaultDomainName = $null
                        HasContract       = $false
                        Relationships     = [System.Collections.Generic.List[object]]::new()
                    }
                }
                $roleIds = @()
                if ($relationship.accessDetails -and $relationship.accessDetails.unifiedRoles) {
                    $roleIds = @($relationship.accessDetails.unifiedRoles | ForEach-Object { [string]$_.roleDefinitionId })
                }
                $customers[$customerTenant].Relationships.Add([pscustomobject]@{
                        Id                          = $relationship.id
                        DisplayName                 = $relationship.displayName
                        Status                      = $relationship.status
                        CreatedDateTime             = $relationship.createdDateTime
                        EndDateTime                 = $relationship.endDateTime
                        AutoExtendDuration          = $relationship.autoExtendDuration
                        RoleCount                   = $roleIds.Count
                        IncludesGlobalAdministrator = $roleIds -contains $globalAdministratorTemplateId
                    })
            }
        }

        foreach ($customer in $customers.Values) {
            if ($tenantFilter -and $customer.TenantId -ne $tenantFilter) { continue }
            if ($Name -and -not ($customer.DisplayName -like $Name -or ($customer.DefaultDomainName -and $customer.DefaultDomainName -like $Name))) { continue }

            $output = [ordered]@{
                PSTypeName        = 'MspGdap.Customer'
                TenantId          = $customer.TenantId
                DisplayName       = $customer.DisplayName
                DefaultDomainName = $customer.DefaultDomainName
                HasContract       = $customer.HasContract
            }
            if ($IncludeGdapStatus) {
                $rels = @($customer.Relationships)
                $active = @($rels | Where-Object { $_.Status -eq 'active' })
                $output.GdapStatus = if ($active.Count -gt 0) { 'active' } elseif ($rels.Count -gt 0) { [string]($rels | Select-Object -First 1).Status } else { 'none' }
                $output.ActiveRelationshipCount = $active.Count
                $output.GdapEndDateTime = ($active | Sort-Object -Property EndDateTime -Descending | Select-Object -First 1).EndDateTime
                $output.PendingApprovalCount = @($rels | Where-Object { $_.Status -eq 'approvalPending' }).Count
                $output.IncludesGlobalAdministrator = [bool](@($active | Where-Object IncludesGlobalAdministrator).Count)
                $output.Relationships = $rels
            }
            [pscustomobject]$output
        }
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}
