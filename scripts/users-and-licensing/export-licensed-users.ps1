#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports the licensed users in one or more customer tenants to CSV.

.DESCRIPTION
    Lists every user in each customer tenant through Microsoft Graph, keeps the
    users that have at least one licence assigned, and returns their display name,
    user principal name, sign-in state and licence names (SKU part numbers).

    Use -IncludeUnlicensed to return every user, with IsLicensed set to false for
    users without a licence. Each customer is processed on its own, so a failure in
    one tenant is recorded as a row with Status 'Failed' and the script moves on.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER IncludeUnlicensed
    Also returns users without a licence.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./export-licensed-users.ps1 -TenantId 'contoso.onmicrosoft.com' -OutputPath ./contoso-licensed-users.csv

    Exports the licensed users of one customer to CSV.

.EXAMPLE
    ./export-licensed-users.ps1 -AllCustomers -IncludeUnlicensed | Where-Object Status -eq 'Failed'

    Runs against every active GDAP customer and shows only the customers that failed.

.NOTES
    Replaces the original 2016 method: Connect-MsolService and Get-MsolUser filtered on IsLicensed (MSOnline module, retired 30 May 2025), plus the Microsoft Online Services Sign-In Assistant.
    Required GDAP roles: Directory Readers or Global Reader.
    Required partner app permissions: Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/how-to-export-a-list-of-office-365-users-to-csv/

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

    [switch]$IncludeUnlicensed,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'DisplayName', 'UserPrincipalName', 'AccountEnabled', 'IsLicensed', 'Licences', 'Status', 'Error')
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

            $skuNames = @{}
            foreach ($sku in @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber')) {
                $skuNames[[string]$sku.skuId] = $sku.skuPartNumber
            }

            $users = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/users?$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses&$top=999')
            foreach ($user in $users) {
                $licences = @($user.assignedLicenses | Where-Object { $_ } | ForEach-Object {
                        $id = [string]$_.skuId
                        if ($skuNames.ContainsKey($id)) { $skuNames[$id] } else { $id }
                    })
                $isLicensed = $licences.Count -gt 0
                if (-not $isLicensed -and -not $IncludeUnlicensed) { continue }

                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId  = $tenant
                    CustomerName      = $customerName
                    DisplayName       = $user.displayName
                    UserPrincipalName = $user.userPrincipalName
                    AccountEnabled    = $user.accountEnabled
                    IsLicensed        = $isLicensed
                    Licences          = ($licences | Sort-Object) -join ', '
                    Status            = 'OK'
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
