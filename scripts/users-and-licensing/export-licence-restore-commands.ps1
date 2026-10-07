#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports every licensed user's licences with a ready-made command to reassign each one, before a CSP partner or licensing change.

.DESCRIPTION
    Moving a customer between CSP partners or distributors, or swapping one
    subscription for another, can briefly leave users without licences. Run this
    script first to keep a record of who had which licence and which service plans
    were turned off, so you can put everything back quickly.

    For each licensed user and each licence, the script returns one row with:
    - the SKU part number and SKU ID,
    - the disabled service plans,
    - whether the licence is assigned directly or inherited from a group,
    - for direct assignments, a RestoreCommand that reassigns the licence with
      Invoke-MspGraphRequest (POST /users/{id}/assignLicense) and supports -WhatIf.

    Licences inherited from a group have no RestoreCommand, because they come back
    when the group's licence is restored. Group-based licensing is the better
    safeguard for this situation, because a group assignment is restored as one
    change instead of one per user.

    The script only reads. It never changes a tenant. Review a RestoreCommand
    before you run it, and run it with -WhatIf first.

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
    ./export-licence-restore-commands.ps1 -TenantId 'contoso.onmicrosoft.com' -OutputPath ./contoso-licences-before-transfer.csv

    Saves the licence assignments of one customer before a CSP transfer.

.EXAMPLE
    Import-Csv ./contoso-licences-before-transfer.csv | Where-Object AssignedBy -match 'Direct' | ForEach-Object {
        $body = @{ addLicenses = @(@{ skuId = $_.SkuId; disabledPlans = @($_.DisabledPlans -split ', ' | Where-Object { $_ }) }); removeLicenses = @() }
        Invoke-MspGraphRequest -TenantId $_.CustomerTenantId -Method POST -Uri "v1.0/users/$($_.UserId)/assignLicense" -Body $body -WhatIf
    }

    Previews reassigning every direct licence from the saved file, built from the
    SkuId, UserId and DisabledPlans columns rather than by running text from the
    file. Remove -WhatIf only after you have checked the preview.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Get-MsolPartnerContract (DAP), Get-MsolUser and Set-MsolUserLicense -AddLicenses with AccountSkuId values (MSOnline module, retired 30 May 2025).
    Required GDAP roles: Directory Readers or Global Reader to export, License Administrator or User Administrator to run the restore commands.
    Required partner app permissions: Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All to export, LicenseAssignment.ReadWrite.All to restore (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/transition-smoothly-office-365-csp-partners/

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
    $columns = @('CustomerTenantId', 'CustomerName', 'DisplayName', 'UserPrincipalName', 'UserId', 'SkuPartNumber', 'SkuId', 'DisabledPlans', 'AssignedBy', 'RestoreCommand', 'Status', 'Error')
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

            $uri = 'v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses,licenseAssignmentStates&$top=999'
            foreach ($user in @(Invoke-MspGraphRequest -TenantId $tenant -Uri $uri)) {
                foreach ($licence in @($user.assignedLicenses | Where-Object { $_ })) {
                    $skuId = [string]$licence.skuId
                    $states = @($user.licenseAssignmentStates | Where-Object { $_ -and [string]$_.skuId -eq $skuId })
                    $direct = @($states | Where-Object { -not $_.assignedByGroup }).Count -gt 0 -or $states.Count -eq 0
                    $groups = @($states | Where-Object { $_.assignedByGroup } | ForEach-Object { [string]$_.assignedByGroup })
                    $disabled = @($licence.disabledPlans | Where-Object { $_ } | ForEach-Object { [string]$_ })

                    $assignedBy = @()
                    if ($direct) { $assignedBy += 'Direct' }
                    if ($groups.Count -gt 0) { $assignedBy += ($groups | ForEach-Object { "Group:$_" }) }

                    $restore = $null
                    if ($direct) {
                        $planList = if ($disabled.Count -gt 0) { "'" + ($disabled -join "', '") + "'" } else { '' }
                        $restore = "Invoke-MspGraphRequest -TenantId '$tenant' -Method POST -Uri 'v1.0/users/$($user.id)/assignLicense' -Body @{ addLicenses = @(@{ skuId = '$skuId'; disabledPlans = @($planList) }); removeLicenses = @() }"
                    }

                    $row = Get-ResultRow -Column $columns -Value @{
                        CustomerTenantId  = $tenant
                        CustomerName      = $customerName
                        DisplayName       = $user.displayName
                        UserPrincipalName = $user.userPrincipalName
                        UserId            = $user.id
                        SkuPartNumber     = if ($skuNames.ContainsKey($skuId)) { $skuNames[$skuId] } else { $null }
                        SkuId             = $skuId
                        DisabledPlans     = $disabled -join ', '
                        AssignedBy        = $assignedBy -join ', '
                        RestoreCommand    = $restore
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
