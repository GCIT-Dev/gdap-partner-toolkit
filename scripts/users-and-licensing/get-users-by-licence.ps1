#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Lists the users who hold a specific licence (SKU) in one or more customer tenants.

.DESCRIPTION
    Looks up the SKU by its part number (for example ENTERPRISEPREMIUM for
    Office 365 E5 or SPB for Microsoft 365 Business Premium) in each customer's
    subscribed SKUs, then asks Microsoft Graph for the users with that SKU
    assigned (directly or through a group).

    Use -ListSkus first to see the SKU part numbers a customer has, with the
    enabled and consumed unit counts. This replaces Get-MsolAccountSku.

    Results are objects, so use -OutputPath (or Export-Csv) for a real CSV file.
    The original article piped the output to Out-File, which writes a text table
    with a .csv name, not a CSV.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER SkuPartNumber
    The SKU part number to look for, such as ENTERPRISEPREMIUM, SPE_E3 or SPB.
    Not case sensitive. Wildcards are allowed (for example *E5*), as the original
    -match filter allowed part of the name. Every matching SKU is listed.

.PARAMETER ListSkus
    Lists the SKUs in each tenant instead of users.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./get-users-by-licence.ps1 -TenantId 'contoso.onmicrosoft.com' -ListSkus

    Shows the SKU part numbers and unit counts in one customer tenant.

.EXAMPLE
    ./get-users-by-licence.ps1 -TenantId 'contoso.onmicrosoft.com' -SkuPartNumber 'ENTERPRISEPREMIUM' -OutputPath ./e5-users.csv

    Exports the users with Office 365 E5 to CSV.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Get-MsolAccountSku and Get-MsolUser filtered on Licenses.AccountSkuId (MSOnline module, retired 30 May 2025).
    Required GDAP roles: Directory Readers or Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/get-office-365-users-specific-license-type-via-powershell/

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(DefaultParameterSetName = 'TenantUsers')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'TenantUsers', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Parameter(Mandatory, ParameterSetName = 'TenantSkus', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllUsers')]
    [Parameter(Mandatory, ParameterSetName = 'AllSkus')]
    [switch]$AllCustomers,

    [Parameter(Mandatory, ParameterSetName = 'TenantUsers')]
    [Parameter(Mandatory, ParameterSetName = 'AllUsers')]
    [ValidatePattern('^[A-Za-z0-9_\-\.\*\?]+$')]
    [SupportsWildcards()]
    [string]$SkuPartNumber,

    [Parameter(Mandatory, ParameterSetName = 'TenantSkus')]
    [Parameter(Mandatory, ParameterSetName = 'AllSkus')]
    [switch]$ListSkus,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'SkuPartNumber', 'SkuId', 'EnabledUnits', 'ConsumedUnits', 'DisplayName', 'UserPrincipalName', 'AccountEnabled', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
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
    if ($PSCmdlet.ParameterSetName -like 'Tenant*') {
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
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $customerName = $knownNames[$tenant]
            if (-not $customerName) {
                $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $skus = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber,prepaidUnits,consumedUnits')

            if ($ListSkus) {
                foreach ($sku in $skus) {
                    $row = Get-ResultRow -Column $columns -Value @{
                        CustomerTenantId = $tenant
                        CustomerName     = $customerName
                        SkuPartNumber    = $sku.skuPartNumber
                        SkuId            = $sku.skuId
                        EnabledUnits     = $sku.prepaidUnits.enabled
                        ConsumedUnits    = $sku.consumedUnits
                        Status           = 'OK'
                    }
                    $results.Add($row)
                    $row
                }
                continue
            }

            $matched = @($skus | Where-Object { $_.skuPartNumber -like $SkuPartNumber })
            if ($matched.Count -eq 0) {
                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId = $tenant
                    CustomerName     = $customerName
                    SkuPartNumber    = $SkuPartNumber
                    Status           = 'SkuNotFound'
                }
                $results.Add($row)
                $row
                continue
            }

            foreach ($sku in $matched) {
                $uri = "v1.0/users?`$filter=assignedLicenses/any(x:x/skuId eq $($sku.skuId))&`$select=id,displayName,userPrincipalName,accountEnabled&`$top=999"
                foreach ($user in @(Invoke-MspGraphRequest -TenantId $tenant -Uri $uri)) {
                    $row = Get-ResultRow -Column $columns -Value @{
                        CustomerTenantId  = $tenant
                        CustomerName      = $customerName
                        SkuPartNumber     = $sku.skuPartNumber
                        SkuId             = $sku.skuId
                        EnabledUnits      = $sku.prepaidUnits.enabled
                        ConsumedUnits     = $sku.consumedUnits
                        DisplayName       = $user.displayName
                        UserPrincipalName = $user.userPrincipalName
                        AccountEnabled    = $user.accountEnabled
                        Status            = 'OK'
                    }
                    $results.Add($row)
                    $row
                }
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
