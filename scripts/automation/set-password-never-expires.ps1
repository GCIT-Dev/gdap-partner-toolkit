#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally sets, the password expiry policy on every verified domain in customer tenants.

.DESCRIPTION
    Reads each verified domain in each customer tenant through Microsoft Graph and reports its
    passwordValidityPeriodInDays value. Microsoft recommends that passwords do not expire unless there is
    evidence of compromise, and the value 2147483647 means "never expire".

    By default the script only reports. Add -Apply to set passwordValidityPeriodInDays to 2147483647 on
    domains that still expire passwords. -WhatIf shows what -Apply would change without changing anything.

    Access comes from the technician's MspGdap token and the GDAP roles in each customer. No password,
    key file or per-customer account is used.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER Apply
    Set the domains that still expire passwords to never expire. Without it the script only reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./set-password-never-expires.ps1 -AllCustomers -OutputPath ./password-expiry.csv

    Reports the password expiry policy for every verified domain of every customer, and saves a CSV.

.EXAMPLE
    ./set-password-never-expires.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows which domains in one customer would be changed to never expire, without changing them.

.NOTES
    Replaces the original 2017 method: an Azure Functions v1 timer function that decrypted a stored delegated
    admin password with an AES key file, ran Connect-MsolService -Credential, Get-MsolPartnerContract and
    Set-MsolPasswordPolicy -ValidityPeriod 2147483647 under DAP.
    Required GDAP roles: Domain Name Administrator (to change the policy), Global Reader (report only).
    Required partner app permissions: Microsoft Graph delegated Domain.ReadWrite.All (update domain, and it
    also covers listing domains), or Domain.Read.All for the report only. Domain.ReadWrite.All is in the full manifest.
    Microsoft Learn: https://learn.microsoft.com/en-us/graph/api/domain-update

.LINK
    https://gcit.com.au/knowledge-base/connect-azure-function-office-365/

.LINK
    https://gcit.com.au/connect-azure-function-office-365/

.LINK
    docs/07-migrating-from-dap-msonline.md

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

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $neverExpire = 2147483647
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-ResultRow {
        param($Customer, $Domain, $CurrentDays, $NotificationDays, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId         = $Customer.TenantId
            CustomerName             = $Customer.Name
            Domain                   = $Domain
            PasswordValidityDays     = $CurrentDays
            PasswordNotificationDays = $NotificationDays
            PasswordsExpire          = if ($null -eq $CurrentDays) { $null } else { [int64]$CurrentDays -ne $neverExpire }
            Action                   = $Action
            Error                    = $ErrorMessage
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

            $domains = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'domains' |
                    Where-Object { $_.isVerified })

            foreach ($domain in $domains) {
                # One domain failing (for example a federated domain the API won't update) must not stop the
                # rest, as in the original.
                $current = $domain.passwordValidityPeriodInDays
                $notice = $domain.passwordNotificationWindowInDays
                try {
                    if ($null -ne $current -and [int64]$current -eq $neverExpire) {
                        $row = ConvertTo-ResultRow -Customer $customer -Domain $domain.id -CurrentDays $current -NotificationDays $notice -Action 'AlreadyNeverExpires'
                    }
                    elseif (-not $Apply) {
                        $row = ConvertTo-ResultRow -Customer $customer -Domain $domain.id -CurrentDays $current -NotificationDays $notice -Action 'WouldSetNeverExpire'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($domain.id) in $($customer.Name)", 'Set password expiry to never expire')) {
                        $body = @{ passwordValidityPeriodInDays = $neverExpire }
                        $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method PATCH -Uri "domains/$($domain.id)" -Body $body -Confirm:$false
                        $check = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri "domains/$($domain.id)"
                        $action = if ([int64]$check.passwordValidityPeriodInDays -eq $neverExpire) { 'SetNeverExpire' } else { 'SetNotConfirmed' }
                        $row = ConvertTo-ResultRow -Customer $customer -Domain $domain.id -CurrentDays $check.passwordValidityPeriodInDays -NotificationDays $check.passwordNotificationWindowInDays -Action $action
                    }
                    else {
                        $row = ConvertTo-ResultRow -Customer $customer -Domain $domain.id -CurrentDays $current -NotificationDays $notice -Action 'WhatIf'
                    }
                }
                catch {
                    Write-Warning "Customer $($customer.TenantId), domain $($domain.id): $($_.Exception.Message)"
                    $row = ConvertTo-ResultRow -Customer $customer -Domain $domain.id -CurrentDays $current -NotificationDays $notice -Action 'Failed' -ErrorMessage $_.Exception.Message
                }
                $results.Add($row)
                $row
            }
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
