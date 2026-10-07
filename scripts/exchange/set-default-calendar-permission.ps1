#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally sets, the Default calendar permission on user mailboxes.

.DESCRIPTION
    The Default entry on a mailbox's calendar decides what everyone else in the organisation
    can see. This script connects to Exchange Online in each customer through GDAP with
    Connect-MspExchangeOnline and reads the Default permission on every user mailbox calendar
    (or only the mailboxes you name).

    It is report-only unless you add -Apply. With -Apply it changes every calendar whose Default
    permission differs from -AccessRights, using Set-MailboxFolderPermission. Use -WhatIf with
    -Apply to preview the changes.

    The calendar is addressed by primary SMTP address, not alias, because an alias is not always
    unique. If the calendar folder is not called Calendar (mailboxes created in another language),
    the script looks up the real folder name with Get-EXOMailboxFolderStatistics.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER AccessRights
    The permission level the Default user should have. LimitedDetails (free/busy, subject and
    location) is the default. AvailabilityOnly shows free/busy only. Reviewer shows full details.

.PARAMETER Mailbox
    Limit the run to these mailboxes (UPN or primary SMTP address). By default every user mailbox
    is included. Shared and resource mailboxes are only included when named here.

.PARAMETER Apply
    Make the change. Without it the script only reports what it would change.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./set-default-calendar-permission.ps1 -TenantId 'contoso.onmicrosoft.com'

    Reports the current Default calendar permission for every user mailbox and what would change.

.EXAMPLE
    ./set-default-calendar-permission.ps1 -TenantId 'contoso.onmicrosoft.com' -AccessRights LimitedDetails -Apply -WhatIf

    Shows each Set-MailboxFolderPermission change without making it. Remove -WhatIf to apply.

.EXAMPLE
    ./set-default-calendar-permission.ps1 -TenantId 'contoso.onmicrosoft.com' -Mailbox 'ceo@contoso.com' -AccessRights AvailabilityOnly -Apply

    Keeps one calendar at free/busy only, as an exception to the organisation default.

.NOTES
    Replaces the original 2015 method: a bulk script that opened a Basic authentication remote
    PowerShell session (New-PSSession to outlook.office365.com/powershell-liveid) with
    Get-Credential and set LimitedDetails on alias:Calendar for every user mailbox.
    Required GDAP roles: Exchange Administrator (Global Reader is enough for report-only runs).
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    The organisation sharing policy (Get-SharingPolicy) controls external calendar sharing. This
    script only changes the internal Default entry.

.LINK
    https://gcit.com.au/knowledge-base/set-default-sharing-policy-for-office-365-users-calendars/

.LINK
    ../../docs/05-exchange-access.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateSet('None', 'AvailabilityOnly', 'LimitedDetails', 'Reviewer')]
    [string]$AccessRights = 'LimitedDetails',

    [string[]]$Mailbox,

    [switch]$Apply,

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
        foreach ($column in 'Mailbox', 'DisplayName', 'CalendarFolder', 'CurrentAccessRights', 'DesiredAccessRights', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    function Get-DefaultCalendarPermission {
        # Returns the folder identity and the Default entry, trying \Calendar first and the real
        # (possibly localised) calendar folder name second.
        param([string]$Address)
        $folder = "$($Address):\Calendar"
        try {
            $entry = Get-EXOMailboxFolderPermission -Identity $folder -User Default -ErrorAction Stop
            return [pscustomobject]@{ Folder = $folder; Entry = $entry }
        }
        catch {
            $calendar = @(Get-EXOMailboxFolderStatistics -Identity $Address -Folderscope Calendar -ErrorAction Stop |
                    Where-Object { $_.FolderType -eq 'Calendar' }) | Select-Object -First 1
            if (-not $calendar) { throw "No default calendar folder found for $Address." }
            $folder = '{0}:{1}' -f $Address, ([string]$calendar.FolderPath).Replace('/', '\')
            $entry = Get-EXOMailboxFolderPermission -Identity $folder -User Default -ErrorAction Stop
            return [pscustomobject]@{ Folder = $folder; Entry = $entry }
        }
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

            $mailboxes = [System.Collections.Generic.List[object]]::new()
            if ($Mailbox) {
                foreach ($identity in $Mailbox) {
                    try { $mailboxes.Add((Get-EXOMailbox -Identity $identity -ErrorAction Stop)) }
                    catch {
                        # One mistyped mailbox should not stop the rest.
                        $row = ConvertTo-ResultRow -Customer $customer -Data @{ Mailbox = $identity; DesiredAccessRights = $AccessRights; Status = 'Failed'; Error = $_.Exception.Message }
                        $results.Add($row)
                        $row
                    }
                }
            }
            else {
                foreach ($found in @(Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop)) { $mailboxes.Add($found) }
            }

            foreach ($item in @($mailboxes)) {
                $address = [string]$item.PrimarySmtpAddress
                $data = @{ Mailbox = $address; DisplayName = $item.DisplayName; DesiredAccessRights = $AccessRights }
                try {
                    $permission = Get-DefaultCalendarPermission -Address $address
                    $current = @($permission.Entry.AccessRights | ForEach-Object { [string]$_ }) -join ','
                    $data['CalendarFolder'] = $permission.Folder
                    $data['CurrentAccessRights'] = $current
                    if ($current -eq $AccessRights) {
                        $data['Status'] = 'AlreadySet'
                    }
                    elseif (-not $Apply) {
                        $data['Status'] = 'WouldChange'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $($permission.Folder)", "Set Default calendar permission to $AccessRights (was $current)")) {
                        Set-MailboxFolderPermission -Identity $permission.Folder -User Default -AccessRights $AccessRights -Confirm:$false -ErrorAction Stop
                        $data['Status'] = 'Changed'
                    }
                    else {
                        $data['Status'] = 'WhatIf'
                    }
                }
                catch {
                    $data['Status'] = 'Failed'
                    $data['Error'] = $_.Exception.Message
                }
                $row = ConvertTo-ResultRow -Customer $customer -Data $data
                $results.Add($row)
                $row
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
