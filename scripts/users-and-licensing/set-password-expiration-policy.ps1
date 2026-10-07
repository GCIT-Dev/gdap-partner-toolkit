#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally sets, "passwords never expire" on every verified domain in one or more customer tenants.

.DESCRIPTION
    Microsoft and current NIST guidance recommend that passwords don't expire on
    a schedule, as long as MFA is in place. In Microsoft Entra ID the setting is
    per domain: passwordValidityPeriodInDays set to 2147483647 means passwords for
    users on that domain never expire.

    For each customer the script reads every domain through Microsoft Graph and
    returns one row per verified, managed domain with its current validity period
    and notification window. With -Apply it sets passwordValidityPeriodInDays to
    2147483647 on the domains that still expire, one ShouldProcess confirmation
    per domain (supports -WhatIf), then reads the domain back to confirm.

    Federated domains are listed but never changed, because their password policy
    is set by the identity provider. The domain setting also applies only to
    cloud-only users. Users synchronised from Active Directory follow the
    on-premises password policy.

    The original article also set PasswordNeverExpires on every single user in
    every customer. That is not needed once the domain policy is set, and this
    script does not do it. Its per-user report is still available: -IncludeUsers
    adds one row per user with whether the user's own password policy disables
    expiry (passwordPolicies DisablePasswordExpiration), when the password was last
    changed and whether the user is synchronised from Active Directory. User rows
    are report only.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER IncludeUsers
    Also returns one row per user (RowType User) with PasswordNeverExpires,
    LastPasswordChangeDateTime and OnPremisesSyncEnabled. Report only.

.PARAMETER Apply
    Sets passwords to never expire on the domains that still expire. Without it
    the script only reports.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./set-password-expiration-policy.ps1 -AllCustomers -OutputPath ./password-expiry.csv

    Reports the password expiry policy of every verified domain in every active GDAP customer.

.EXAMPLE
    ./set-password-expiration-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows which domains would be changed in one customer. Remove -WhatIf to change them.

.NOTES
    Replaces the original 2017 method: Connect-MsolService -Credential with a named admin account (no MFA), Get-MsolPartnerContract (DAP), Get-MsolPasswordPolicy and Set-MsolPasswordPolicy -ValidityPeriod 2147483647, and Set-MsolUser -PasswordNeverExpires on every user (MSOnline module, retired 30 May 2025).
    Required GDAP roles: Global Reader or Directory Readers to report, Domain Name Administrator for -Apply.
    Required partner app permissions: Microsoft Graph delegated Domain.Read.All to report, User.Read.All for -IncludeUsers (covered in manifests/partner-app.full.json by Directory.ReadWrite.All and User.ReadWrite.All) and Domain.ReadWrite.All for -Apply (in manifests/partner-app.full.json).

.LINK
    https://gcit.com.au/knowledge-base/set-office-365-password-expiration-policy-on-all-delegated-customer-tenants/

.LINK
    https://learn.microsoft.com/en-us/microsoft-365/admin/manage/set-password-expiration-policy

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [switch]$IncludeUsers,

    [switch]$Apply,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $neverExpires = 2147483647
    $columns = @('CustomerTenantId', 'CustomerName', 'RowType', 'Domain', 'AuthenticationType', 'PasswordValidityPeriodInDays', 'PasswordNotificationWindowInDays', 'PasswordsExpire', 'UserPrincipalName', 'DisplayName', 'PasswordNeverExpires', 'LastPasswordChangeDateTime', 'OnPremisesSyncEnabled', 'Action', 'Status', 'Error')
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

            $domains = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/domains?$select=id,isVerified,authenticationType,passwordValidityPeriodInDays,passwordNotificationWindowInDays' | Where-Object { $_.isVerified })
            foreach ($domain in $domains) {
                $validity = $domain.passwordValidityPeriodInDays
                $expires = $validity -ne $neverExpires
                $values = [ordered]@{
                    CustomerTenantId                 = $tenant
                    CustomerName                     = $customerName
                    RowType                          = 'Domain'
                    Domain                           = $domain.id
                    AuthenticationType               = $domain.authenticationType
                    PasswordValidityPeriodInDays     = $validity
                    PasswordNotificationWindowInDays = $domain.passwordNotificationWindowInDays
                    PasswordsExpire                  = $expires
                    Action                           = 'None'
                    Status                           = 'OK'
                }

                if ($domain.authenticationType -eq 'Federated') {
                    $values.Action = 'SkippedFederated'
                }
                elseif (-not $expires) {
                    $values.Action = 'AlreadyNeverExpires'
                }
                elseif (-not $Apply) {
                    $values.Action = 'WouldSetNeverExpires'
                }
                elseif ($PSCmdlet.ShouldProcess("$($domain.id) in $customerName ($tenant)", 'Set passwords to never expire')) {
                    try {
                        $encoded = [uri]::EscapeDataString($domain.id)
                        Invoke-MspGraphRequest -TenantId $tenant -Method PATCH -Uri "v1.0/domains/$encoded" -Body @{ passwordValidityPeriodInDays = $neverExpires } -Confirm:$false | Out-Null
                        $check = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/domains/{0}?$select=id,passwordValidityPeriodInDays' -f $encoded)
                        $values.PasswordValidityPeriodInDays = $check.passwordValidityPeriodInDays
                        $values.PasswordsExpire = $check.passwordValidityPeriodInDays -ne $neverExpires
                        $values.Action = if ($values.PasswordsExpire) { 'Unconfirmed' } else { 'Changed' }
                        if ($values.PasswordsExpire) { $values.Status = 'Warning' }
                    }
                    catch {
                        $values.Action = 'ChangeFailed'
                        $values.Status = 'Failed'
                        $values.Error = $_.Exception.Message
                    }
                }
                else {
                    $values.Action = 'WhatIf'
                }

                $row = Get-ResultRow -Column $columns -Value $values
                $results.Add($row)
                $row
            }

            if ($IncludeUsers) {
                $userUri = 'v1.0/users?$select=id,displayName,userPrincipalName,passwordPolicies,lastPasswordChangeDateTime,onPremisesSyncEnabled&$top=999'
                foreach ($user in @(Invoke-MspGraphRequest -TenantId $tenant -Uri $userUri)) {
                    $row = Get-ResultRow -Column $columns -Value @{
                        CustomerTenantId           = $tenant
                        CustomerName               = $customerName
                        RowType                    = 'User'
                        Domain                     = ([string]$user.userPrincipalName -split '@', 2)[-1]
                        UserPrincipalName          = $user.userPrincipalName
                        DisplayName                = $user.displayName
                        PasswordNeverExpires       = [string]$user.passwordPolicies -match 'DisablePasswordExpiration'
                        LastPasswordChangeDateTime = $user.lastPasswordChangeDateTime
                        OnPremisesSyncEnabled      = [bool]$user.onPremisesSyncEnabled
                        Action                     = 'None'
                        Status                     = 'OK'
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
