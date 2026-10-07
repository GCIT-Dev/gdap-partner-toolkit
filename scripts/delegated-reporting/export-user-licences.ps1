#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports every licensed user and their licences across all your customers (or the customers you name).

.DESCRIPTION
    For each customer the script reads the users and the subscribed SKUs through
    Microsoft Graph and returns one row per licensed user, with the licence names
    (SKU part numbers), whether each licence comes from a group, the account's
    sign-in state and its usage location.

    Customers are processed one at a time with their own token, so a customer
    where you have no suitable GDAP role is recorded as a row with Status 'Failed'
    and the script moves on to the next one.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-user-licences.ps1 -AllCustomers -OutputPath ./user-licence-report.csv

    Exports every licensed user in every active GDAP customer.

.EXAMPLE
    Get-MspCustomer -Name 'contoso*' | ./export-user-licences.ps1

    Exports the licensed users of the customers whose name starts with contoso.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Get-MsolPartnerContract -All (DAP) and Get-MsolUser -TenantId -All with Licenses.AccountSkuId (MSOnline module, retired 30 May 2025).
    Required GDAP roles: Directory Readers or Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/export-list-office-365-users-licenses-customer-tenants-delegated-administration/

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

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'DisplayName', 'UserPrincipalName', 'AccountEnabled', 'UsageLocation', 'Licences', 'GroupAssignedLicences', 'Status', 'Error')
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
            Write-Verbose -Message "Reading licensed users for $customerName ($tenant)"

            $skuNames = @{}
            foreach ($sku in @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber')) {
                $skuNames[[string]$sku.skuId] = $sku.skuPartNumber
            }

            $uri = 'v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,usageLocation,assignedLicenses,licenseAssignmentStates&$top=999'
            foreach ($user in @(Invoke-MspGraphRequest -TenantId $tenant -Uri $uri)) {
                $assigned = @($user.assignedLicenses | Where-Object { $_ })
                if ($assigned.Count -eq 0) { continue }

                $names = @($assigned | ForEach-Object {
                        $id = [string]$_.skuId
                        if ($skuNames.ContainsKey($id)) { $skuNames[$id] } else { $id }
                    })
                $groupNames = @($user.licenseAssignmentStates | Where-Object { $_ -and $_.assignedByGroup } | ForEach-Object {
                        $id = [string]$_.skuId
                        if ($skuNames.ContainsKey($id)) { $skuNames[$id] } else { $id }
                    } | Sort-Object -Unique)

                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId      = $tenant
                    CustomerName          = $customerName
                    DisplayName           = $user.displayName
                    UserPrincipalName     = $user.userPrincipalName
                    AccountEnabled        = $user.accountEnabled
                    UsageLocation         = $user.usageLocation
                    Licences              = ($names | Sort-Object) -join ', '
                    GroupAssignedLicences = $groupNames -join ', '
                    Status                = 'OK'
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
