#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Gives your unattended automation app Exchange Online access in customer tenants, and finds the standing
    Exchange admin accounts that older scripts created for the same purpose.

.DESCRIPTION
    This is the safe replacement for creating an Exchange administrator user account, with a shared
    password, in every customer tenant. No user account or password is created. Instead:

    1. Test-MspExchangeAppAccess checks whether your separate automation app (not the MspGdap partner app)
       already has the Exchange.ManageAsApp app role and a directory role in the customer.
    2. With -Apply, Enable-MspExchangeAppAccess adds what is missing and reads it back. Scheduled jobs then
       connect with Connect-MspExchangeOnline -AppOnly and a certificate (docs/05 and docs/08).
    3. With -LegacyAccountUpnPrefix, the script lists user accounts whose user principal name starts with
       that prefix (for example the reporting accounts an older script created) with their enabled state
       and directory roles. With -DisableLegacyAccounts and -Apply it blocks sign-in for them. Review and
       then delete those accounts, and rotate anything that used their shared password.

    By default the script only reports. -WhatIf shows every change -Apply would make.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER AppId
    Application (client) ID of your separate multi-tenant automation app, registered in your partner tenant.

.PARAMETER Role
    Directory role for the automation app's service principal. Exchange Administrator (default), Exchange
    Recipient Administrator, Global Reader or Security Reader. Use the narrowest role your jobs need.

.PARAMETER LegacyAccountUpnPrefix
    User principal name prefix of standing admin accounts created by older scripts, for example 'msp-reports'.

.PARAMETER DisableLegacyAccounts
    With -Apply, block sign-in for the accounts found with -LegacyAccountUpnPrefix.

.PARAMETER Apply
    Make the changes. Without it the script only reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./enable-exchange-app-access.ps1 -AllCustomers -AppId '00000000-0000-0000-0000-000000000000' -LegacyAccountUpnPrefix 'msp-reports' -OutputPath ./exchange-app-access.csv

    Reports, for every customer, whether the automation app is ready for Exchange and whether any legacy
    reporting accounts still exist.

.EXAMPLE
    ./enable-exchange-app-access.ps1 -TenantId 'contoso.onmicrosoft.com' -AppId '00000000-0000-0000-0000-000000000000' -Role 'Exchange Recipient Administrator' -LegacyAccountUpnPrefix 'msp-reports' -DisableLegacyAccounts -Apply -WhatIf

    Shows the app role, directory role and account changes that would be made in one customer.

.NOTES
    Replaces the original 2017 method: three Azure Functions v1 functions that used MSOnline and DAP
    (New-MsolUser, Add-MsolRoleMember -RoleName 'Exchange Service Administrator', Set-MsolUser
    -BlockCredential) to create a standing Exchange administrator with the same password in every customer
    tenant, and unblocked it on demand from an HTTP function.
    Required GDAP roles: Cloud Application Administrator and Privileged Role Administrator (just in time) for
    -Apply, Global Reader for the report, Privileged Authentication Administrator to block sign-in for
    accounts that hold admin roles.
    Required partner app permissions: Microsoft Graph delegated Application.ReadWrite.All,
    AppRoleAssignment.ReadWrite.All and RoleManagement.ReadWrite.Directory (Enable-MspExchangeAppAccess and
    Test-MspExchangeAppAccess), User.ReadWrite.All (find and block legacy accounts). All are in the full
    manifest.

.LINK
    https://gcit.com.au/knowledge-base/create-exchange-administrators-customer-office-365-tenants-using-azure-functions-delegated-administration/

.LINK
    docs/05-exchange-access.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string]$AppId,

    [ValidateSet('Exchange Administrator', 'Exchange Recipient Administrator', 'Global Reader', 'Security Reader')]
    [string]$Role = 'Exchange Administrator',

    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string]$LegacyAccountUpnPrefix,

    [switch]$DisableLegacyAccounts,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-ResultRow {
        param($Customer, $Ready, $Action, $Outcome, [string[]]$Legacy, $LegacyAction, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId     = $Customer.TenantId
            CustomerName         = $Customer.Name
            AppId                = $AppId
            Role                 = $Role
            AppAccessReady       = $Ready
            Action               = $Action
            Outcome              = $Outcome
            LegacyAccounts       = ($Legacy -join '; ')
            LegacyAccountsAction = $LegacyAction
            Error                = $ErrorMessage
        }
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $test = Test-MspExchangeAppAccess -TenantId $customer.TenantId -AppId $AppId -Role $Role
            $ready = [bool]$test.Success
            $outcome = $test.Outcome
            if ($ready) {
                $action = 'AlreadyEnabled'
            }
            elseif (-not $Apply) {
                $action = 'WouldEnable'
            }
            elseif ($PSCmdlet.ShouldProcess($customer.Name, "Enable Exchange app-only access for $AppId with role $Role")) {
                $enable = Enable-MspExchangeAppAccess -TenantId $customer.TenantId -AppId $AppId -Role $Role -Confirm:$false -ErrorAction SilentlyContinue |
                    Select-Object -Last 1
                $ready = [bool]$enable.Success
                $outcome = $enable.Outcome
                $action = if ($ready) { 'Enabled' } else { 'EnableFailed' }
            }
            else {
                $action = 'WhatIf'
            }

            $legacy = @()
            $legacyAction = $null
            if ($LegacyAccountUpnPrefix) {
                $accounts = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "users?`$filter=startswith(userPrincipalName,'$LegacyAccountUpnPrefix')&`$select=id,userPrincipalName,accountEnabled")
                $legacyAction = if ($accounts.Count -eq 0) { 'NoneFound' } elseif ($DisableLegacyAccounts) { 'Pending' } else { 'ReportOnly' }
                foreach ($account in $accounts) {
                    $roles = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "roleManagement/directory/roleAssignments?`$filter=principalId eq '$($account.id)'&`$expand=roleDefinition" |
                            ForEach-Object { $_.roleDefinition.displayName })
                    $enabled = [bool]$account.accountEnabled
                    if ($DisableLegacyAccounts -and $enabled) {
                        if (-not $Apply) {
                            $legacyAction = 'WouldDisable'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($account.userPrincipalName) in $($customer.Name)", 'Block sign-in')) {
                            $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method PATCH -Uri "users/$($account.id)" -Body @{ accountEnabled = $false } -Confirm:$false
                            $enabled = [bool](Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "users/$($account.id)?`$select=accountEnabled").accountEnabled
                            $legacyAction = if ($enabled) { 'DisableNotConfirmed' } else { 'Disabled' }
                        }
                        else {
                            $legacyAction = 'WhatIf'
                        }
                    }
                    $state = if ($enabled) { 'enabled' } else { 'blocked' }
                    $legacy += "$($account.userPrincipalName) ($state, roles: $(if ($roles) { $roles -join ', ' } else { 'none' }))"
                }
            }

            $row = ConvertTo-ResultRow -Customer $customer -Ready $ready -Action $action -Outcome $outcome -Legacy $legacy -LegacyAction $legacyAction
            $results.Add($row)
            $row
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
