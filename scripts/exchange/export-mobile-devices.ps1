#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Exports the mobile devices that sync with Exchange Online mailboxes.

.DESCRIPTION
    Connects to Exchange Online in each customer through GDAP with Connect-MspExchangeOnline and
    lists every mobile device partnership (Exchange ActiveSync, Outlook for iOS and Android, and
    other clients that Get-MobileDevice reports), with the mailbox it belongs to. It reads only
    and changes nothing.

    The original script called Get-MobileDevice once per mailbox. This version reads every device
    in one call and matches it to its mailbox by distinguished name, which is much faster in
    larger tenants.

    The columns of the original export are kept (FriendlyName, ClientType, ClientVersion, DeviceId,
    DeviceMobileOperator, DeviceModel, DeviceOS, DeviceTelephoneNumber, DeviceType, FirstSyncTime and
    UserDisplayName), plus the device access state and the Intune managed and compliant flags.

    Devices enrolled in Intune are better reported from Intune (Microsoft Graph managedDevices).
    This report covers what Exchange Online knows about, including unmanaged phones.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./export-mobile-devices.ps1 -TenantId 'contoso.onmicrosoft.com' -OutputPath 'C:\Reports\MobileDevices.csv'

    Exports the mobile devices of one customer to a CSV file.

.EXAMPLE
    ./export-mobile-devices.ps1 -AllCustomers | Where-Object DeviceAccessState -ne 'Allowed'

    Lists devices in every GDAP customer that are not in the Allowed state.

.NOTES
    Replaces the original 2016 method: a single-tenant script that opened a Basic authentication
    remote PowerShell session (New-PSSession to outlook.office365.com/powershell-liveid) with
    Get-Credential and ran Get-MobileDevice -Mailbox for every mailbox.
    Required GDAP roles: Exchange Administrator, or Global Reader for this read-only report.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

.LINK
    https://gcit.com.au/knowledge-base/export-a-list-of-mobile-devices-connected-to-office-365/

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

    $deviceColumns = 'FriendlyName', 'DeviceModel', 'DeviceOS', 'DeviceType', 'ClientType', 'ClientVersion', 'DeviceId',
    'DeviceMobileOperator', 'DeviceTelephoneNumber', 'DeviceAccessState', 'DeviceAccessStateReason', 'IsManaged',
    'IsCompliant', 'FirstSyncTime', 'UserDisplayName'

    function ConvertTo-ResultRow {
        param([object]$Customer, [hashtable]$Data, [string[]]$Column)
        $row = [ordered]@{ CustomerTenantId = $Customer.TenantId; CustomerName = $Customer.Name; UserPrincipalName = $Data['UserPrincipalName']; DisplayName = $Data['DisplayName'] }
        foreach ($name in $Column) { $row[$name] = $Data[$name] }
        $row['Status'] = $Data['Status']
        $row['Error'] = $Data['Error']
        [pscustomobject]$row
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
            $row = ConvertTo-ResultRow -Customer $customer -Column $deviceColumns -Data @{ Status = 'Skipped'; Error = $customer.Skip }
            $results.Add($row)
            $row
            continue
        }

        $connection = $null
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop

            $byDistinguishedName = @{}
            foreach ($mailbox in @(Get-EXOMailbox -ResultSize Unlimited -Properties DistinguishedName -ErrorAction Stop)) {
                if ($mailbox.DistinguishedName) { $byDistinguishedName[[string]$mailbox.DistinguishedName] = $mailbox }
            }

            $devices = @(Get-MobileDevice -ResultSize Unlimited -ErrorAction Stop)
            foreach ($device in $devices) {
                # A device's DN is CN=<device>,CN=ExchangeActiveSyncDevices,<mailbox DN>.
                $owner = $null
                if ([string]$device.DistinguishedName -match ',CN=ExchangeActiveSyncDevices,(?<parent>.+)$') {
                    $owner = $byDistinguishedName[$Matches['parent']]
                }
                $data = @{
                    UserPrincipalName = if ($owner) { [string]$owner.UserPrincipalName } else { $null }
                    DisplayName       = if ($owner) { $owner.DisplayName } else { [string]$device.UserDisplayName }
                    Status            = 'Found'
                }
                foreach ($name in $deviceColumns) { $data[$name] = $device.$name }
                $row = ConvertTo-ResultRow -Customer $customer -Column $deviceColumns -Data $data
                $results.Add($row)
                $row
            }

            if ($devices.Count -eq 0) {
                $row = ConvertTo-ResultRow -Customer $customer -Column $deviceColumns -Data @{ Status = 'NoneFound' }
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Column $deviceColumns -Data @{ Status = 'Failed'; Error = $_.Exception.Message }
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
