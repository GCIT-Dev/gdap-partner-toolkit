#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft Purview privileged access management for Exchange Online: whether it is on, its approval policies and recent requests.

.DESCRIPTION
    Privileged access management (PAM) in Microsoft Purview gives zero standing access to chosen
    Exchange Online tasks: an admin raises a request (New-ElevatedAccessRequest), an approver
    approves it, and the access lasts only for the requested time. It still applies to Exchange
    Online tasks only, and is configured in the Microsoft 365 admin center or in Exchange Online
    PowerShell.

    This script connects to Exchange Online in each customer through GDAP with
    Connect-MspExchangeOnline and reports, without changing anything:
    - whether PAM is turned on (ElevatedAccessControl in Get-OrganizationConfig),
    - each approval policy (Get-ElevatedAccessApprovalPolicy), and
    - each access request (Get-ElevatedAccessRequest).

    Raising a request stays a deliberate, manual step, for example:
    New-ElevatedAccessRequest -Task 'Exchange\New-JournalRule' -Reason 'Journal rule for ticket 123' -DurationHours 4

    The cmdlets are documented in the Microsoft Purview article on configuring privileged access
    management, not in the Exchange cmdlet reference, so their output properties can change. The
    script reads them defensively and puts anything it does not recognise in the Detail column.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER IncludeRequests
    Also list access requests (Get-ElevatedAccessRequest).

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./get-privileged-access-management.ps1 -AllCustomers -OutputPath 'C:\Reports\PrivilegedAccess.csv'

    Shows which customers have privileged access management on, and their approval policies.

.EXAMPLE
    ./get-privileged-access-management.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeRequests

    Lists one customer's approval policies and access requests.

.NOTES
    Replaces the original 2018 method: a single New-ElevatedAccessRequest example, run in an
    Exchange Online PowerShell session (at the time, the Basic authentication New-PSSession
    connection), with licensing named as the Advanced Compliance SKU.
    Required GDAP roles: Exchange Administrator.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Licensing: privileged access management needs Microsoft 365 E5, Office 365 E5 or a Microsoft
    Purview compliance add-on that includes it. Check the current Microsoft Purview service
    description for the customer's plan.

    Microsoft Learn: Get started with privileged access management
    https://learn.microsoft.com/en-us/purview/privileged-access-management-configuration

.LINK
    https://gcit.com.au/privileged-access-management-in-office-365/

.LINK
    ../../docs/05-exchange-access.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [switch]$IncludeRequests,

    [string]$OutputPath
)

begin {
    function Get-TargetCustomer {
        param([string[]]$Tenant, [switch]$All)
        if ($All) {
            foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus -ErrorAction Stop)) {
                $skip = $null
                if ($customer.GdapStatus -ne 'active') { $skip = "No active GDAP relationship (status: $($customer.GdapStatus))." }
                [pscustomobject]@{ TenantId = $customer.TenantId; Name = $customer.DisplayName; Skip = $skip }
            }
            return
        }
        foreach ($entry in $Tenant) {
            $id = $entry
            $name = $entry
            try {
                $customer = @(Get-MspCustomer -TenantId $entry -IncludeGdapStatus -ErrorAction Stop) | Select-Object -First 1
                if ($customer) {
                    $id = $customer.TenantId
                    $name = $customer.DisplayName
                }
            }
            catch {
                Write-Verbose "Could not look up the customer name for $entry. $($_.Exception.Message)"
            }
            [pscustomobject]@{ TenantId = $id; Name = $name; Skip = $null }
        }
    }

    function Close-CustomerSession {
        param([object]$Connection)
        if (-not (Get-Command -Name 'Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue)) { return }
        $disconnect = @{ Confirm = $false; WhatIf = $false; ErrorAction = 'SilentlyContinue' }
        if ($Connection -and $Connection.ConnectionId) { $disconnect['ConnectionId'] = $Connection.ConnectionId }
        Disconnect-ExchangeOnline @disconnect
    }

    function ConvertTo-ResultRow {
        param([object]$Customer, [hashtable]$Data)
        $row = [ordered]@{ CustomerTenantId = $Customer.TenantId; CustomerName = $Customer.Name }
        foreach ($column in 'RowType', 'Task', 'Value', 'Detail', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    function Get-FirstPropertyValue {
        # Returns the first property that exists on the object, so a renamed property does not break the report.
        param([object]$InputObject, [string[]]$Name)
        foreach ($candidate in $Name) {
            $property = $InputObject.PSObject.Properties[$candidate]
            if ($property -and $null -ne $property.Value -and [string]$property.Value -ne '') { return $property.Value }
        }
        $null
    }

    function Format-Detail {
        param([object]$InputObject, [string[]]$Skip)
        $pairs = foreach ($property in $InputObject.PSObject.Properties) {
            if ($property.Name -in $Skip -or $property.Name -like 'PS*' -or $null -eq $property.Value -or [string]$property.Value -eq '') { continue }
            '{0}={1}' -f $property.Name, (@($property.Value) -join ',')
        }
        @($pairs) -join '; '
    }

    $tenantInput = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
}

process {
    foreach ($entry in $TenantId) {
        if ($entry) { $tenantInput.Add($entry) }
    }
}

end {
    $customers = @(Get-TargetCustomer -Tenant $tenantInput -All:$AllCustomers)
    foreach ($customer in $customers) {
        if ($customer.Skip) {
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'Skipped'; Error = $customer.Skip }
            $results.Add($row)
            $row
            continue
        }

        $connection = $null
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop

            $organization = Get-OrganizationConfig -ErrorAction Stop
            $enabled = Get-FirstPropertyValue -InputObject $organization -Name 'ElevatedAccessControl'
            $approvers = Get-FirstPropertyValue -InputObject $organization -Name 'ElevatedAccessApprovers'
            $row = ConvertTo-ResultRow -Customer $customer -Data @{
                RowType = 'Organisation'
                Value   = if ($null -eq $enabled) { 'Unknown' } else { [string]$enabled }
                Detail  = if ($approvers) { "Approvers: $(@($approvers) -join ',')" } else { $null }
                Status  = 'Reported'
            }
            $results.Add($row)
            $row

            foreach ($policy in @(Get-ElevatedAccessApprovalPolicy -ErrorAction Stop)) {
                $row = ConvertTo-ResultRow -Customer $customer -Data @{
                    RowType = 'ApprovalPolicy'
                    Task    = [string](Get-FirstPropertyValue -InputObject $policy -Name 'Task', 'Name', 'Identity')
                    Value   = [string](Get-FirstPropertyValue -InputObject $policy -Name 'ApprovalType')
                    Detail  = Format-Detail -InputObject $policy -Skip 'Task', 'ApprovalType'
                    Status  = 'Reported'
                }
                $results.Add($row)
                $row
            }

            if ($IncludeRequests) {
                foreach ($request in @(Get-ElevatedAccessRequest -ErrorAction Stop)) {
                    $row = ConvertTo-ResultRow -Customer $customer -Data @{
                        RowType = 'Request'
                        Task    = [string](Get-FirstPropertyValue -InputObject $request -Name 'Task', 'Identity')
                        Value   = [string](Get-FirstPropertyValue -InputObject $request -Name 'RequestStatus', 'ApprovalStatus', 'Status', 'State')
                        Detail  = Format-Detail -InputObject $request -Skip 'Task', 'RequestStatus'
                        Status  = 'Reported'
                    }
                    $results.Add($row)
                    $row
                }
            }
        }
        catch {
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Status = 'Failed'; Error = $_.Exception.Message }
            $results.Add($row)
            $row
        }
        finally {
            Close-CustomerSession -Connection $connection
        }
    }

    if ($OutputPath) {
        $folder = Split-Path -Path $OutputPath -Parent
        if ($folder -and -not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false }
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false
        Write-Verbose "Saved $($results.Count) row(s) to $OutputPath."
    }
}
