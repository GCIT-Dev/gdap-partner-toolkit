#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally turns on, unified audit log ingestion in GDAP customer tenants.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline), reads Get-AdminAuditLogConfig and reports whether unified audit
    log ingestion is on. With -Apply it turns ingestion on with Set-AdminAuditLogConfig, running
    Enable-OrganizationCustomization first when the organisation is still dehydrated, and then
    reads the setting back.

    No accounts are created in any customer. The technician's own GDAP role does the work, and
    every change is logged against that named technician.

    With -LegacyAdminPrefix the script also lists user accounts whose user principal name starts
    with that prefix. Use it to find temporary admin accounts left behind by the retired version of
    this script, which created a Global Administrator with a shared password in each customer. The
    script only reports those accounts. Review each one and delete it in the Microsoft Entra admin
    center, because a prefix match can also catch a real user.

    Microsoft says it can take up to 60 minutes for the change to take effect.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER Apply
    Turn on unified audit log ingestion where it is off. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER LegacyAdminPrefix
    Optional user principal name prefix of temporary admin accounts created by the retired script,
    for example 'auditadmin'. Matching accounts are reported, never changed.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./enable-unified-audit-log.ps1 -AllCustomers -OutputPath ./ual-status.csv

    Reports the unified audit log status of every GDAP customer and saves it to CSV.
.EXAMPLE
    ./enable-unified-audit-log.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -LegacyAdminPrefix 'auditadmin' -WhatIf

    Shows what would change in one customer and lists any leftover temporary admin accounts.
.NOTES
    Replaces the original 2017 method: MSOnline and DAP (Get-MsolPartnerContract), temporary Global
    Administrator accounts with a shared password created with New-MsolUser, and basic authentication
    remote PowerShell (New-PSSession and Invoke-Command to ps.outlook.com).
    Required GDAP roles: Exchange Administrator (or Global Reader for report-only runs). User
    Administrator or Global Reader to read accounts with -LegacyAdminPrefix.
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated). With
    -LegacyAdminPrefix, Microsoft Graph User.Read.All (delegated), which User.ReadWrite.All in the
    partner app manifests satisfies.
.LINK
    https://gcit.com.au/knowledge-base/enabling-unified-audit-log-delegated-office-365-tenants-via-powershell/
.LINK
    https://gcit.com.au/enabling-unified-audit-log-delegated-office-365-tenants-via-powershell/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    https://learn.microsoft.com/en-us/purview/audit-log-enable-disable
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

    [switch]$Apply,

    [ValidatePattern('^[A-Za-z0-9._-]{3,64}$')]
    [string]$LegacyAdminPrefix,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'UnifiedAuditLogEnabled', 'WasDehydrated', 'Action', 'LegacyAdminAccounts', 'Detail')

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
        $values = @{}
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        try {
            Write-Verbose "Checking unified audit log for $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            $auditConfig = Get-AdminAuditLogConfig
            $enabled = [bool]$auditConfig.UnifiedAuditLogIngestionEnabled
            $values.UnifiedAuditLogEnabled = $enabled
            $values.Action = 'None'

            if (-not $enabled) {
                $values.Action = 'ReportOnly'
                if ($Apply) {
                    $organisation = Get-OrganizationConfig
                    $values.WasDehydrated = [bool]$organisation.IsDehydrated
                    if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) ($($customer.TenantId))", 'Turn on unified audit log ingestion')) {
                        if ($organisation.IsDehydrated) {
                            Enable-OrganizationCustomization
                        }
                        Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
                        $readBack = [bool](Get-AdminAuditLogConfig).UnifiedAuditLogIngestionEnabled
                        $values.UnifiedAuditLogEnabled = $readBack
                        $values.Action = if ($readBack) { 'Enabled' } else { 'EnabledPendingReadback' }
                    }
                    else {
                        $values.Action = 'WhatIf'
                    }
                }
            }

            if ($LegacyAdminPrefix) {
                $filter = "startswith(userPrincipalName,'$LegacyAdminPrefix')"
                $uri = "v1.0/users?`$filter=$filter&`$select=id,userPrincipalName,accountEnabled,createdDateTime"
                $legacy = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri $uri)
                $values.LegacyAdminAccounts = ($legacy | ForEach-Object { $_.userPrincipalName }) -join ', '
            }

            $values.Status = 'Succeeded'
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
