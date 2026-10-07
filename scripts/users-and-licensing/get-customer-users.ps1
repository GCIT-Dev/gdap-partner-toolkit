#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds a customer by name and lists the users in its tenant through GDAP.

.DESCRIPTION
    The original article looked up a customer's tenant ID with
    Get-MsolPartnerContract, kept it in a global variable and passed it to Msol
    cmdlets with -TenantId. This script does the same job with MspGdap:

    - -CustomerName searches your customers by display name or default domain
      (wildcards allowed) with Get-MspCustomer, so you don't need to look up the
      tenant ID first.
    - -TenantId and -AllCustomers work as in the other scripts in this folder.
    - Users are read through Microsoft Graph in each customer tenant.
    - -UserPrincipalName returns one user instead of the whole directory.

    The original article also unblocked a shared, unlicensed admin account in the
    customer tenant and reset its password to a fixed value so it could be used
    for Exchange Online. That workaround is not needed and must not be used. With
    GDAP, connect to Exchange Online as yourself with
    Connect-MspExchangeOnline -TenantId, using your Exchange Administrator or
    Exchange Recipient Administrator role. Shared standing admin accounts in
    customer tenants should be removed, not reactivated.

    The script only reads. It never changes a tenant.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER CustomerName
    Part of a customer's display name or default domain. Wildcards are allowed,
    and the value is wrapped in * when it has none. Every matching customer is
    listed, with a warning when more than one matches.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER UserPrincipalName
    Returns only this user (user principal name or object ID).

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./get-customer-users.ps1 -CustomerName 'contoso' | Format-Table CustomerName, DisplayName, UserPrincipalName, AccountEnabled

    Finds customers whose name or domain contains "contoso" and lists their users.

.EXAMPLE
    ./get-customer-users.ps1 -TenantId 'contoso.onmicrosoft.com' -UserPrincipalName 'adele.vance@contoso.com'

    Shows one user in one customer tenant.

.NOTES
    Replaces the original 2017 method: Connect-MsolService, Get-MsolPartnerContract (DAP) and Get-MsolUser -TenantId, plus Set-MsolUser -BlockCredential and Set-MsolUserPassword on a shared customer admin account (removed as unsafe).
    Required GDAP roles: Directory Readers or Global Reader in each customer.
    Required partner app permissions: Microsoft Graph delegated User.Read.All in the customer, and Directory.Read.All in the partner tenant for Get-MspCustomer (covered in manifests/partner-app.full.json by User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/managing-users-in-office-365-delegated-tenants-via-powershell/

.LINK
    docs/05-exchange-access.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'Name')]
    [ValidateNotNullOrEmpty()]
    [string]$CustomerName,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'DisplayName', 'UserPrincipalName', 'UserId', 'AccountEnabled', 'UserType', 'IsLicensed', 'CreatedDateTime', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()
    $select = 'id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses,createdDateTime'

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
    if ($AllCustomers -or $PSCmdlet.ParameterSetName -eq 'Name') {
        $customers = @(Get-MspCustomer -IncludeGdapStatus | Where-Object { $_.GdapStatus -eq 'active' })
        if ($PSCmdlet.ParameterSetName -eq 'Name') {
            $pattern = if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($CustomerName)) { $CustomerName } else { "*$CustomerName*" }
            $customers = @($customers | Where-Object { $_.DisplayName -like $pattern -or ($_.DefaultDomainName -and $_.DefaultDomainName -like $pattern) })
            if ($customers.Count -eq 0) {
                Write-Warning -Message "No customer with an active GDAP relationship matches '$CustomerName'."
            }
            elseif ($customers.Count -gt 1) {
                Write-Warning -Message "$($customers.Count) customers match '$CustomerName': $(($customers.DisplayName) -join ', '). Users from all of them are listed. Use a longer name or -TenantId to pick one."
            }
        }
        foreach ($customer in $customers) {
            $targets.Add($customer.TenantId)
            $knownNames[$customer.TenantId] = $customer.DisplayName
        }
    }

    foreach ($target in $targets) {
        $tenant = $target
        $name = $null
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $name = $knownNames[$tenant]
            if (-not $name) {
                $name = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $users = if ($UserPrincipalName) {
                @(Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/users/{0}?$select={1}' -f [uri]::EscapeDataString($UserPrincipalName), $select))
            }
            else {
                @(Invoke-MspGraphRequest -TenantId $tenant -Uri "v1.0/users?`$select=$select&`$top=999")
            }

            foreach ($user in $users) {
                $row = Get-ResultRow -Column $columns -Value @{
                    CustomerTenantId  = $tenant
                    CustomerName      = $name
                    DisplayName       = $user.displayName
                    UserPrincipalName = $user.userPrincipalName
                    UserId            = $user.id
                    AccountEnabled    = $user.accountEnabled
                    UserType          = $user.userType
                    IsLicensed        = @($user.assignedLicenses | Where-Object { $_ }).Count -gt 0
                    CreatedDateTime   = $user.createdDateTime
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
                CustomerName     = $name
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
