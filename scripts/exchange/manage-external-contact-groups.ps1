#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Manages distribution groups of external contacts: report members, add a contact, remove a contact from every group, or create a group.

.DESCRIPTION
    Some businesses keep company-wide distribution groups of external people (suppliers, board
    members, partners). In Exchange Online that means a mail contact for each external address,
    added to a distribution group. This script replaces a menu-driven tool with one parameter per
    task. It connects to Exchange Online in each customer through GDAP with
    Connect-MspExchangeOnline.

    -Action Report (the default) lists every distribution group and its members, with the
    external address of mail contacts.

    -Action AddContact finds or creates the mail contact for -ContactEmail (New-MailContact) and
    adds it to -GroupIdentity (Add-DistributionGroupMember).

    -Action RemoveContact removes the contact for -ContactEmail from every distribution group it
    is a member of (Remove-DistributionGroupMember). The contact itself is kept unless you add
    -DeleteContact, which then deletes the mail contact (Remove-MailContact), as the original
    menu's "remove a contact" option did.

    -Action NewGroup creates a distribution group called -GroupName (New-DistributionGroup).

    AddContact, RemoveContact and NewGroup change nothing unless you add -Apply. Use -WhatIf with
    -Apply to preview.

    Permissions: the original article told admins to add staff to the Organization Management role
    group, which is close to full Exchange administration. Creating mail contacts and groups only
    needs the Recipient Management role group in the customer, and for a partner technician the
    Exchange Recipient Administrator GDAP role. Changing the members of a group you are not an
    owner of (ManagedBy) needs -BypassSecurityGroupManagerCheck, and that needs Exchange
    Administrator. Use that role only for those runs, or make the right person an owner of the group.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped. Only allowed with -Action Report.

.PARAMETER Action
    Report (default), AddContact, RemoveContact or NewGroup.

.PARAMETER ContactEmail
    External email address of the contact, for AddContact and RemoveContact.

.PARAMETER ContactName
    Display name for a new mail contact. Defaults to the email address.

.PARAMETER GroupIdentity
    The distribution group (name, alias or email address) for AddContact.

.PARAMETER GroupName
    Name of the new distribution group for NewGroup.

.PARAMETER GroupAddress
    Optional email address for the new distribution group.

.PARAMETER DeleteContact
    With -Action RemoveContact, also delete the mail contact after removing it from every group.

.PARAMETER BypassSecurityGroupManagerCheck
    Pass -BypassSecurityGroupManagerCheck to the membership cmdlets, for groups the technician
    does not own. Needs Exchange Administrator.

.PARAMETER Apply
    Make the change. Without it the script only reports what it would do.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./manage-external-contact-groups.ps1 -TenantId 'contoso.onmicrosoft.com' -OutputPath 'C:\Reports\DistributionGroups.csv'

    Lists every distribution group and its members.

.EXAMPLE
    ./manage-external-contact-groups.ps1 -TenantId 'contoso.onmicrosoft.com' -Action AddContact -ContactEmail 'alex@fabrikam.com' -ContactName 'Alex (Fabrikam)' -GroupIdentity 'Suppliers' -Apply -WhatIf

    Shows the contact creation and group membership change without making them. Remove -WhatIf to apply.

.EXAMPLE
    ./manage-external-contact-groups.ps1 -TenantId 'contoso.onmicrosoft.com' -Action RemoveContact -ContactEmail 'alex@fabrikam.com' -Apply

    Removes the contact from every distribution group it belongs to. Add -DeleteContact to delete
    the mail contact too.

.NOTES
    Replaces the original 2016 method: a downloadable menu script whose connect option opened a
    Basic authentication remote PowerShell session (New-PSSession to
    outlook.office365.com/powershell-liveid), with advice to add users to the Organization
    Management role group so they could run it.
    Required GDAP roles: Exchange Recipient Administrator (Exchange Administrator for
    -BypassSecurityGroupManagerCheck, which membership changes need on groups the technician does
    not own). Global Reader is enough for -Action Report.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: New-MailContact
    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/new-mailcontact

.LINK
    https://gcit.com.au/simplify-external-contact-group-management-powershell/

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

    [ValidateSet('Report', 'AddContact', 'RemoveContact', 'NewGroup')]
    [string]$Action = 'Report',

    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$ContactEmail,

    [string]$ContactName,

    [string]$GroupIdentity,

    [string]$GroupName,

    [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
    [string]$GroupAddress,

    [switch]$DeleteContact,

    [switch]$BypassSecurityGroupManagerCheck,

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
        foreach ($column in 'Group', 'GroupAddress', 'Member', 'MemberType', 'ExternalAddress', 'Action', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    function Get-ContactAddress {
        param([object]$Recipient)
        $external = [string]$Recipient.ExternalEmailAddress
        if ($external) { return ($external -replace '^(?i)smtp:', '') }
        [string]$Recipient.PrimarySmtpAddress
    }

    switch ($Action) {
        'AddContact' { if (-not $ContactEmail -or -not $GroupIdentity) { throw '-Action AddContact needs -ContactEmail and -GroupIdentity.' } }
        'RemoveContact' { if (-not $ContactEmail) { throw '-Action RemoveContact needs -ContactEmail.' } }
        'NewGroup' { if (-not $GroupName) { throw '-Action NewGroup needs -GroupName.' } }
    }
    if ($DeleteContact -and $Action -ne 'RemoveContact') {
        throw '-DeleteContact only works with -Action RemoveContact.'
    }
    if ($Action -ne 'Report' -and $AllCustomers) {
        throw "-Action $Action changes one customer at a time. Use -TenantId."
    }
    $bypass = @{}
    if ($BypassSecurityGroupManagerCheck) { $bypass['BypassSecurityGroupManagerCheck'] = $true }

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
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Action = $Action; Status = 'Skipped'; Error = $customer.Skip }
            $results.Add($row)
            $row
            continue
        }

        $connection = $null
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop

            switch ($Action) {
                'Report' {
                    foreach ($group in @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)) {
                        $groupAddress = [string]$group.PrimarySmtpAddress
                        $members = @(Get-DistributionGroupMember -Identity $groupAddress -ResultSize Unlimited -ErrorAction Stop)
                        if ($members.Count -eq 0) {
                            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Group = $group.DisplayName; GroupAddress = $groupAddress; Action = 'Report'; Status = 'Empty' }
                            $results.Add($row)
                            $row
                        }
                        foreach ($item in $members) {
                            $data = @{
                                Group           = $group.DisplayName
                                GroupAddress    = $groupAddress
                                Member          = $item.DisplayName
                                MemberType      = [string]$item.RecipientTypeDetails
                                ExternalAddress = if ([string]$item.RecipientType -in 'MailContact', 'MailUser') { Get-ContactAddress -Recipient $item } else { $null }
                                Action          = 'Report'
                                Status          = 'Reported'
                            }
                            $row = ConvertTo-ResultRow -Customer $customer -Data $data
                            $results.Add($row)
                            $row
                        }
                    }
                }
                'AddContact' {
                    $group = Get-DistributionGroup -Identity $GroupIdentity -ErrorAction Stop
                    $groupAddress = [string]$group.PrimarySmtpAddress
                    $contact = $null
                    try { $contact = Get-MailContact -Identity $ContactEmail -ErrorAction Stop } catch { $contact = $null }
                    $contactData = @{ Group = $group.DisplayName; GroupAddress = $groupAddress; Member = if ($ContactName) { $ContactName } else { $ContactEmail }; MemberType = 'MailContact'; ExternalAddress = $ContactEmail; Action = 'CreateContact' }
                    $contactReady = [bool]$contact
                    if ($contact) {
                        $contactData['Status'] = 'AlreadyExists'
                    }
                    elseif (-not $Apply) {
                        $contactData['Status'] = 'WouldCreate'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $ContactEmail", 'Create mail contact')) {
                        $name = if ($ContactName) { $ContactName } else { $ContactEmail }
                        $contact = New-MailContact -Name $name -ExternalEmailAddress $ContactEmail -Confirm:$false -ErrorAction Stop
                        $contactData['Status'] = 'Created'
                        $contactReady = $true
                    }
                    else {
                        $contactData['Status'] = 'WhatIf'
                    }
                    $row = ConvertTo-ResultRow -Customer $customer -Data $contactData
                    $results.Add($row)
                    $row

                    $memberData = @{ Group = $group.DisplayName; GroupAddress = $groupAddress; Member = $contactData['Member']; MemberType = 'MailContact'; ExternalAddress = $ContactEmail; Action = 'AddMember' }
                    $members = @(Get-DistributionGroupMember -Identity $groupAddress -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { (Get-ContactAddress -Recipient $_).ToLowerInvariant() })
                    if ($members -contains $ContactEmail.ToLowerInvariant()) {
                        $memberData['Status'] = 'AlreadyMember'
                    }
                    elseif (-not $Apply) {
                        $memberData['Status'] = 'WouldAdd'
                    }
                    elseif (-not $contactReady) {
                        $memberData['Status'] = 'WhatIf'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $groupAddress", "Add $ContactEmail")) {
                        Add-DistributionGroupMember -Identity $groupAddress -Member ([string]$contact.Identity) @bypass -Confirm:$false -ErrorAction Stop
                        $memberData['Status'] = 'Added'
                    }
                    else {
                        $memberData['Status'] = 'WhatIf'
                    }
                    $row = ConvertTo-ResultRow -Customer $customer -Data $memberData
                    $results.Add($row)
                    $row
                }
                'RemoveContact' {
                    $contact = Get-MailContact -Identity $ContactEmail -ErrorAction Stop
                    $target = $ContactEmail.ToLowerInvariant()
                    $found = 0
                    $membershipFailed = $false
                    foreach ($group in @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop)) {
                        $groupAddress = [string]$group.PrimarySmtpAddress
                        $members = @(Get-DistributionGroupMember -Identity $groupAddress -ResultSize Unlimited -ErrorAction Stop | ForEach-Object { (Get-ContactAddress -Recipient $_).ToLowerInvariant() })
                        if ($members -notcontains $target) { continue }
                        $found++
                        $data = @{ Group = $group.DisplayName; GroupAddress = $groupAddress; Member = $contact.DisplayName; MemberType = 'MailContact'; ExternalAddress = $ContactEmail; Action = 'RemoveMember' }
                        if (-not $Apply) {
                            $data['Status'] = 'WouldRemove'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $groupAddress", "Remove $ContactEmail")) {
                            try {
                                Remove-DistributionGroupMember -Identity $groupAddress -Member ([string]$contact.Identity) @bypass -Confirm:$false -ErrorAction Stop
                                $data['Status'] = 'Removed'
                            }
                            catch {
                                $data['Status'] = 'Failed'
                                $data['Error'] = $_.Exception.Message
                                $membershipFailed = $true
                            }
                        }
                        else {
                            $data['Status'] = 'WhatIf'
                        }
                        $row = ConvertTo-ResultRow -Customer $customer -Data $data
                        $results.Add($row)
                        $row
                    }
                    if ($found -eq 0) {
                        $row = ConvertTo-ResultRow -Customer $customer -Data @{ Member = $contact.DisplayName; ExternalAddress = $ContactEmail; Action = 'RemoveMember'; Status = 'NotAMember' }
                        $results.Add($row)
                        $row
                    }
                    if ($DeleteContact) {
                        $deleteData = @{ Member = $contact.DisplayName; MemberType = 'MailContact'; ExternalAddress = $ContactEmail; Action = 'DeleteContact' }
                        if (-not $Apply) {
                            $deleteData['Status'] = 'WouldDelete'
                        }
                        elseif ($membershipFailed) {
                            $deleteData['Status'] = 'Skipped'
                            $deleteData['Error'] = 'Not deleted, because removing it from a group failed.'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $ContactEmail", 'Delete mail contact')) {
                            try {
                                Remove-MailContact -Identity ([string]$contact.Identity) -Confirm:$false -ErrorAction Stop
                                $deleteData['Status'] = 'Deleted'
                            }
                            catch {
                                $deleteData['Status'] = 'Failed'
                                $deleteData['Error'] = $_.Exception.Message
                            }
                        }
                        else {
                            $deleteData['Status'] = 'WhatIf'
                        }
                        $row = ConvertTo-ResultRow -Customer $customer -Data $deleteData
                        $results.Add($row)
                        $row
                    }
                }
                'NewGroup' {
                    $data = @{ Group = $GroupName; GroupAddress = $GroupAddress; Action = 'CreateGroup' }
                    $existing = $null
                    try { $existing = Get-DistributionGroup -Identity $GroupName -ErrorAction Stop } catch { $existing = $null }
                    if ($existing) {
                        $data['GroupAddress'] = [string]$existing.PrimarySmtpAddress
                        $data['Status'] = 'AlreadyExists'
                    }
                    elseif (-not $Apply) {
                        $data['Status'] = 'WouldCreate'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $GroupName", 'Create distribution group')) {
                        $newGroup = @{ Name = $GroupName; Type = 'Distribution'; Confirm = $false; ErrorAction = 'Stop' }
                        if ($GroupAddress) { $newGroup['PrimarySmtpAddress'] = $GroupAddress }
                        $created = New-DistributionGroup @newGroup
                        $data['GroupAddress'] = [string]$created.PrimarySmtpAddress
                        $data['Status'] = 'Created'
                    }
                    else {
                        $data['Status'] = 'WhatIf'
                    }
                    $row = ConvertTo-ResultRow -Customer $customer -Data $data
                    $results.Add($row)
                    $row
                }
            }
        }
        catch {
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ Action = $Action; Status = 'Failed'; Error = $_.Exception.Message }
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
