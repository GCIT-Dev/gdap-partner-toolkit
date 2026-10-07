#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally grants, one user's access to every mailbox calendar in a GDAP customer.
.DESCRIPTION
    Connects to each customer's Exchange Online with the technician's delegated GDAP token
    (Connect-MspExchangeOnline), finds every mailbox's calendar folder (by folder type, so it also
    works when the folder name is localised) and reads the user's current permission on it.

    Without -Apply the script only reports which calendars would change. With -Apply it adds the
    permission with Add-MailboxFolderPermission where the user has none, or changes it with
    Set-MailboxFolderPermission where the user has a different one, then reads it back.

    The user must have a mailbox in the same customer. A customer where the user isn't found is
    reported as Failed and skipped, so -AllCustomers is only useful with an address that exists in
    several customers.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER User
    Email address or user principal name of the user who needs access to the calendars.
.PARAMETER AccessRight
    Calendar permission level to grant. Choose the lowest level the user needs. Reviewer gives
    read-only access to calendar items. The original script used PublishingAuthor.
.PARAMETER Apply
    Add or change permissions. Without this switch the script only reports.
    Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./grant-calendar-access.ps1 -TenantId 'contoso.onmicrosoft.com' -User 'reception@contoso.com' -AccessRight Reviewer

    Lists which calendars the receptionist can't yet read.
.EXAMPLE
    ./grant-calendar-access.ps1 -TenantId 'contoso.onmicrosoft.com' -User 'reception@contoso.com' -AccessRight Editor -Apply -WhatIf

    Shows which calendar permissions would be added or changed, without changing anything.
.NOTES
    Replaces the original 2017 method: basic authentication remote PowerShell (New-PSSession to
    outlook.office365.com/powershell-liveid/ and ?DelegatedOrg=), MSOnline and DAP
    (Get-MsolPartnerContract, Get-MsolDomain), and advice to allow-list IP addresses to skip MFA.
    Required GDAP roles: Exchange Administrator or Exchange Recipient Administrator (Global Reader
    for report-only runs).
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
.LINK
    https://gcit.com.au/knowledge-base/give-office-365-user-access-calendars-via-powershell/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/add-mailboxfolderpermission
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
    [string]$User,

    [Parameter(Mandatory)]
    [ValidateSet('AvailabilityOnly', 'LimitedDetails', 'Reviewer', 'Author', 'NonEditingAuthor', 'PublishingAuthor', 'Editor', 'PublishingEditor')]
    [string]$AccessRight,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @('Status', 'Mailbox', 'CalendarFolder', 'User', 'CurrentAccessRights', 'RequestedAccessRight', 'Action', 'Detail')

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

    function Get-CurrentRight {
        param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Grantee)
        try {
            $permission = Get-EXOMailboxFolderPermission -Identity $Folder -User $Grantee -ErrorAction Stop
            @($permission.AccessRights | ForEach-Object { [string]$_ })
        }
        catch {
            # No entry for this user on the folder.
            @()
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
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; User = $User; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        $customerRows = [System.Collections.Generic.List[object]]::new()
        try {
            Write-Verbose "Checking calendar permissions for $User in $($customer.DisplayName)"
            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId

            try {
                $grantee = Get-EXOMailbox -Identity $User
            }
            catch {
                throw "User '$User' has no mailbox in this customer. $($_.Exception.Message)"
            }
            $granteeAddress = [string]$grantee.PrimarySmtpAddress

            $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox, RoomMailbox, EquipmentMailbox)
            foreach ($mailbox in $mailboxes) {
                $address = [string]$mailbox.PrimarySmtpAddress
                if ($address -eq $granteeAddress) { continue }

                $values = @{ Mailbox = $address; User = $granteeAddress; RequestedAccessRight = $AccessRight; Status = 'Succeeded' }
                try {
                    $calendar = Get-EXOMailboxFolderStatistics -Identity $address -Folderscope Calendar |
                        Where-Object { $_.FolderType -eq 'Calendar' } |
                        Select-Object -First 1
                    if (-not $calendar) { throw 'No default calendar folder found.' }
                    $folder = '{0}:\{1}' -f $address, $calendar.Name
                    $values.CalendarFolder = $folder

                    $current = @(Get-CurrentRight -Folder $folder -Grantee $granteeAddress)
                    $values.CurrentAccessRights = $current -join ', '

                    if ($current -contains $AccessRight) {
                        $values.Action = 'AlreadyPresent'
                    }
                    elseif (-not $Apply) {
                        $values.Action = 'ReportOnly'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.DisplayName) calendar $folder", "Give $granteeAddress $AccessRight")) {
                        if ($current.Count -eq 0) {
                            $null = Add-MailboxFolderPermission -Identity $folder -User $granteeAddress -AccessRights $AccessRight
                        }
                        else {
                            $null = Set-MailboxFolderPermission -Identity $folder -User $granteeAddress -AccessRights $AccessRight
                        }
                        $after = @(Get-CurrentRight -Folder $folder -Grantee $granteeAddress)
                        $values.CurrentAccessRights = $after -join ', '
                        if ($after -contains $AccessRight) {
                            $values.Action = if ($current.Count -eq 0) { 'Added' } else { 'Updated' }
                        }
                        else {
                            $values.Action = 'ChangeNotConfirmed'
                            $values.Status = 'Failed'
                        }
                    }
                    else {
                        $values.Action = 'WhatIf'
                    }
                }
                catch {
                    $values.Status = 'Failed'
                    $values.Detail = $_.Exception.Message
                }
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values $values))
            }
        }
        catch {
            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; User = $User; Detail = $_.Exception.Message }))
        }
        finally {
            Disconnect-ExchangeOnline -Confirm:$false -WhatIf:$false -ErrorAction SilentlyContinue
        }

        foreach ($row in $customerRows) {
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8
        Write-Verbose "Saved $($results.Count) rows to $OutputPath"
    }
}
