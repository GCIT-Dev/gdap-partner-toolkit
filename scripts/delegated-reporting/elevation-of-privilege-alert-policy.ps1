#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports admin privilege alert policies in GDAP customers and optionally adds an alert policy that emails you.
.DESCRIPTION
    Activity alerts (New-ActivityAlert) that the original script created have been replaced by alert
    policies. Microsoft 365 now includes a default alert policy named "Elevation of Exchange admin
    privilege", which raises an alert when someone is given Exchange admin permissions.

    For each customer the script connects to Security and Compliance PowerShell with the
    technician's delegated GDAP token (Connect-MspSecurityCompliance, confirmed in a live test) and
    reports:
      - whether the default "Elevation of Exchange admin privilege" policy exists, is turned on and
        whom it notifies, and
      - whether a custom alert policy named -PolicyName exists.

    Without -Apply the script only reports. With -Apply it creates the custom policy with
    New-ProtectionAlert where it is missing. The policy watches the Add-RoleGroupMember activity
    (Exchange admin audit records use the cmdlet name as the activity) and emails -NotifyUser for
    every occurrence. Single-event policies like this one don't need an E5 licence.

    Security and Compliance PowerShell ignores -WhatIf, so the script decides itself and never
    calls New-ProtectionAlert unless -Apply is set and ShouldProcess agrees.

    If Connect-MspSecurityCompliance fails in a customer, connect interactively with
    Connect-IPPSSession -DelegatedOrganization, as described in docs/05.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER NotifyUser
    Email addresses that receive alerts from the custom policy, for example your service desk.
.PARAMETER PolicyName
    Name of the custom alert policy.
.PARAMETER Operation
    Audited activity the custom policy watches. Defaults to Add-RoleGroupMember.
.PARAMETER Apply
    Create the custom alert policy where it is missing. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./elevation-of-privilege-alert-policy.ps1 -AllCustomers -NotifyUser 'alerts@contoso.com' -OutputPath ./alert-policies.csv

    Reports the admin privilege alert policies of every GDAP customer.
.EXAMPLE
    ./elevation-of-privilege-alert-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -NotifyUser 'alerts@contoso.com' -Apply -WhatIf

    Shows whether the custom policy would be created in one customer, without changing anything.
.NOTES
    Replaces the original 2018 method: basic authentication remote PowerShell to Security and
    Compliance (New-PSSession to ps.compliance.protection.outlook.com and ?DelegatedOrg=), MSOnline
    and DAP (Get-MsolPartnerContract, Get-MsolDomain), and activity alerts (New-ActivityAlert -Type
    ElevationOfPrivilege), which alert policies have replaced.
    Required GDAP roles: Security Administrator or Compliance Administrator (alert policies need the
    Manage Alerts role, which those roles include). Security Reader or Global Reader for report-only runs.
    Required partner app permissions: a delegated token for the Security and Compliance PowerShell
    resource (https://ps.compliance.protection.outlook.com by default in Connect-MspSecurityCompliance).
    The partner app manifests do not list that resource, and none is needed: a live test on
    7 October 2026 connected with the existing consent. Try one customer first.
.LINK
    https://gcit.com.au/knowledge-base/get-alerts-elevation-privilege-operations-office-365-customer-tenants/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    https://learn.microsoft.com/en-us/purview/alert-policies
.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-protectionalert
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant', SupportsShouldProcess, ConfirmImpact = 'Medium')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [Parameter(Mandatory)]
    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string[]]$NotifyUser,

    [ValidateNotNullOrEmpty()]
    [string]$PolicyName = 'Exchange admin role group change',

    [ValidateNotNullOrEmpty()]
    [string]$Operation = 'Add-RoleGroupMember',

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $defaultPolicyName = 'Elevation of Exchange admin privilege'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @(
        'Status', 'DefaultPolicyPresent', 'DefaultPolicyEnabled', 'DefaultPolicyNotifyUser',
        'CustomPolicyName', 'CustomPolicyPresent', 'Action', 'Detail'
    )

    function ConvertTo-ResultRow {
        param(
            [Parameter(Mandatory)][object]$Customer,
            [System.Collections.IDictionary]$Values = @{}
        )
        $row = [ordered]@{
            CustomerTenantId = [string]$Customer.TenantId
            CustomerName     = [string]$Customer.DisplayName
        }
        foreach ($column in $resultColumns) {
            $row[$column] = if ($Values.Contains($column)) { $Values[$column] } else { $null }
        }
        [pscustomobject]$row
    }

    function Get-TargetCustomer {
        param([switch]$All, [string[]]$Requested)
        if ($All) {
            return @(Get-MspCustomer -IncludeGdapStatus)
        }
        foreach ($id in $Requested) {
            $match = $null
            try {
                $match = Get-MspCustomer -TenantId $id -IncludeGdapStatus | Select-Object -First 1
            }
            catch {
                Write-Warning "Could not look up '$id' in your customer list: $($_.Exception.Message)"
            }
            if ($match) { $match } else { [pscustomobject]@{ TenantId = $id; DisplayName = $id } }
        }
    }
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        $values = @{ CustomPolicyName = $PolicyName }
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $values.Status = 'Skipped'
            $values.Detail = "No active GDAP relationship (status: $($customer.GdapStatus))."
            $row = ConvertTo-ResultRow -Customer $customer -Values $values
            $results.Add($row)
            $row
            continue
        }

        try {
            Write-Verbose "Checking alert policies for $($customer.DisplayName)"
            $null = Connect-MspSecurityCompliance -TenantId $customer.TenantId -WarningAction SilentlyContinue

            $policies = @(Get-ProtectionAlert)
            $default = $policies | Where-Object { $_.Name -eq $defaultPolicyName } | Select-Object -First 1
            $custom = $policies | Where-Object { $_.Name -eq $PolicyName } | Select-Object -First 1

            $values.DefaultPolicyPresent = [bool]$default
            $values.DefaultPolicyEnabled = if ($default) { -not [bool]$default.Disabled } else { $null }
            $values.DefaultPolicyNotifyUser = if ($default) { @($default.NotifyUser) -join ', ' } else { $null }
            $values.CustomPolicyPresent = [bool]$custom
            $values.Action = if ($custom) { 'None' } else { 'ReportOnly' }

            if (-not $custom -and $Apply) {
                if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) ($($customer.TenantId))", "Create alert policy '$PolicyName' for $Operation")) {
                    $alertParams = @{
                        Name            = $PolicyName
                        Category        = 'AccessGovernance'
                        ThreatType      = 'Activity'
                        Operation       = $Operation
                        NotifyUser      = $NotifyUser
                        AggregationType = 'None'
                        Severity        = 'High'
                        Description     = 'Alerts when a user is added to an Exchange Online admin role group.'
                    }
                    $null = New-ProtectionAlert @alertParams
                    $check = Get-ProtectionAlert -Identity $PolicyName
                    $values.CustomPolicyPresent = [bool]$check
                    $values.Action = if ($check) { 'Created' } else { 'ChangeNotConfirmed' }
                }
                else {
                    $values.Action = 'WhatIf'
                }
            }
            $values.Status = if ($values.Action -eq 'ChangeNotConfirmed') { 'Failed' } else { 'Succeeded' }
        }
        catch {
            $values.Status = 'Failed'
            $values.Detail = $_.Exception.Message
        }
        finally {
            Disconnect-ExchangeOnline -Confirm:$false -WhatIf:$false -ErrorAction SilentlyContinue
        }

        $row = ConvertTo-ResultRow -Customer $customer -Values $values
        $results.Add($row)
        $row
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
        Write-Verbose "Saved $($results.Count) rows to $OutputPath"
    }
}
