#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Lists the customers with at least a given number of licensed users, with their licensed user count and subscriptions.

.DESCRIPTION
    For each customer the script counts the users with at least one licence and
    lists the tenant's SKUs (with consumed units) through Microsoft Graph. By
    default it returns only the customers with -MinimumLicensedUsers (5) or more
    licensed users. Use -IncludeAll to return every customer with a
    MeetsThreshold column instead.

    Customers are processed one at a time with their own token, so a customer
    where you have no suitable GDAP role is recorded as a row with Status 'Failed'
    and the script moves on.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER MinimumLicensedUsers
    The licensed user count a customer needs to be included. Default 5.

.PARAMETER IncludeAll
    Returns every customer, with MeetsThreshold set to true or false.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-customers-by-licensed-user-count.ps1 -AllCustomers -OutputPath ./customers-5-or-more.csv

    Exports the customers with five or more licensed users.

.EXAMPLE
    ./export-customers-by-licensed-user-count.ps1 -AllCustomers -MinimumLicensedUsers 20 -IncludeAll | Sort-Object LicensedUserCount -Descending

    Lists every customer by licensed user count and marks the ones with 20 or more.

.NOTES
    Replaces the original 2018 method: AzureAD module (retired), Connect-AzureAD -Credential (no MFA), Get-AzureADContract (DAP), then Connect-AzureAD -TenantId, Get-AzureADUser and Get-AzureADSubscribedSku per customer.
    Required GDAP roles: Directory Readers or Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/export-list-customers-certain-number-licensed-office-365-users/

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

    [ValidateRange(1, 1000000)]
    [int]$MinimumLicensedUsers = 5,

    [switch]$IncludeAll,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'LicensedUserCount', 'MeetsThreshold', 'Licences', 'Status', 'Error')
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
    if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
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

            $users = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/users?$select=id,assignedLicenses&$top=999')
            $licensedCount = @($users | Where-Object { @($_.assignedLicenses | Where-Object { $_ }).Count -gt 0 }).Count
            $meets = $licensedCount -ge $MinimumLicensedUsers
            if (-not $meets -and -not $IncludeAll) { continue }

            $skus = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuPartNumber,consumedUnits')
            $row = Get-ResultRow -Column $columns -Value @{
                CustomerTenantId  = $tenant
                CustomerName      = $customerName
                LicensedUserCount = $licensedCount
                MeetsThreshold    = $meets
                Licences          = ($skus | Sort-Object -Property skuPartNumber | ForEach-Object { '{0} ({1})' -f $_.skuPartNumber, $_.consumedUnits }) -join ', '
                Status            = 'OK'
            }
            $results.Add($row)
            $row
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
