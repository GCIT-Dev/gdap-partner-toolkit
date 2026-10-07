#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports every customer's subscriptions with enabled, consumed and unused licence counts.

.DESCRIPTION
    For each customer the script reads the subscribed SKUs through Microsoft
    Graph and returns one row per SKU with:
    - EnabledUnits, the licences that are active and can be assigned,
    - WarningUnits and SuspendedUnits, licences in a grace or suspended state,
    - ConsumedUnits, the licences assigned to users,
    - UnusedUnits, EnabledUnits minus ConsumedUnits.

    Use -OnlyUnused to keep only the SKUs with at least one unused licence. Free
    and trial SKUs with very large unit counts (for example Microsoft Power
    Automate Free) can be excluded with -ExcludeSkuPartNumber.

    Billing reports in Partner Center or your distributor's portal show what you
    are charged for. This report shows what each tenant has assigned, which is
    what you need to find licences to reclaim.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER OnlyUnused
    Returns only SKUs with at least one unused licence.

.PARAMETER ExcludeSkuPartNumber
    SKU part numbers to leave out, such as FLOW_FREE or POWER_BI_STANDARD.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-unused-licences.ps1 -AllCustomers -OnlyUnused -OutputPath ./unused-licences.csv

    Exports the SKUs with unused licences in every active GDAP customer.

.EXAMPLE
    ./export-unused-licences.ps1 -TenantId 'contoso.onmicrosoft.com' -ExcludeSkuPartNumber 'FLOW_FREE', 'POWER_BI_STANDARD'

    Shows one customer's subscriptions without two free SKUs.

.NOTES
    Replaces the original 2017 method: Connect-MsolService -Credential (no MFA), Get-MsolPartnerContract -All (DAP) and Get-MsolAccountSku -TenantId with ActiveUnits minus ConsumedUnits (MSOnline module, retired 30 May 2025).
    Required GDAP roles: Global Reader or Directory Readers.
    Required partner app permissions: Microsoft Graph delegated LicenseAssignment.Read.All (covered in manifests/partner-app.full.json by Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/export-list-unused-office-365-licenses-delegated-administration-tenants/

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [switch]$OnlyUnused,

    [string[]]$ExcludeSkuPartNumber,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'DefaultDomain', 'SkuPartNumber', 'SkuId', 'CapabilityStatus', 'AppliesTo', 'EnabledUnits', 'WarningUnits', 'SuspendedUnits', 'ConsumedUnits', 'UnusedUnits', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $knownDomains = @{}
    $results = [System.Collections.Generic.List[object]]::new()

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
            $knownDomains[$customer.TenantId] = $customer.DefaultDomainName
        }
    }

    foreach ($target in $targets) {
        $tenant = $target
        $customerName = $null
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $customerName = $knownNames[$tenant]
            $defaultDomain = $knownDomains[$tenant]
            if (-not $customerName -or -not $defaultDomain) {
                $organisation = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName,verifiedDomains')[0]
                if (-not $customerName) { $customerName = $organisation.displayName }
                if (-not $defaultDomain) { $defaultDomain = (@($organisation.verifiedDomains) | Where-Object { $_.isDefault } | Select-Object -First 1).name }
            }

            $skus = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber,capabilityStatus,appliesTo,prepaidUnits,consumedUnits')
            foreach ($sku in $skus) {
                if ($ExcludeSkuPartNumber -and $ExcludeSkuPartNumber -contains $sku.skuPartNumber) { continue }
                $enabled = [int]$sku.prepaidUnits.enabled
                $consumed = [int]$sku.consumedUnits
                $unused = $enabled - $consumed
                if ($OnlyUnused -and $unused -le 0) { continue }

                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId = $tenant
                    CustomerName     = $customerName
                    DefaultDomain    = $defaultDomain
                    SkuPartNumber    = $sku.skuPartNumber
                    SkuId            = $sku.skuId
                    CapabilityStatus = $sku.capabilityStatus
                    AppliesTo        = $sku.appliesTo
                    EnabledUnits     = $enabled
                    WarningUnits     = [int]$sku.prepaidUnits.warning
                    SuspendedUnits   = [int]$sku.prepaidUnits.suspended
                    ConsumedUnits    = $consumed
                    UnusedUnits      = $unused
                    Status           = 'OK'
                }
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $row = Get-ResultRow -Column $columns -Value @{
                CustomerTenantId = $tenant
                CustomerName     = $customerName
                Status           = 'Failed'
                Error            = $_.Exception.Message
            }
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
