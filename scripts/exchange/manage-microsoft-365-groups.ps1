#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Lists Microsoft 365 Groups (formerly Outlook or Office 365 Groups) and their members, and optionally creates a group or adds members.

.DESCRIPTION
    Connects to Exchange Online in each customer through GDAP with Connect-MspExchangeOnline and
    uses the UnifiedGroup cmdlets, which still work in Exchange Online PowerShell.

    -Action Report (the default) lists every Microsoft 365 Group with its address, privacy and
    member count. Add -IncludeMembers for one row per member.

    -Action AddMember adds -Member to the group named in -Identity (Add-UnifiedGroupLinks
    -LinkType Members). Members who are already in the group are reported and skipped.

    -Action Create creates a group with -DisplayName and -Alias (New-UnifiedGroup). It is private
    unless you pass -AccessType Public. New-UnifiedGroup on its own creates public groups, which is
    what the original article's command did.

    AddMember and Create change nothing unless you add -Apply. Use -WhatIf with -Apply to preview.
    Deleting groups is deliberately not scripted here. Remove-UnifiedGroup is easy to run by hand,
    and a deleted group stays restorable for 30 days.

    These cmdlets need a delegated (technician) session. Microsoft notes that New-UnifiedGroup and
    Add-UnifiedGroupLinks don't work with app-only Exchange connections.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped. Only allowed with -Action Report.

.PARAMETER Action
    Report (default), AddMember or Create.

.PARAMETER IncludeMembers
    With -Action Report, list the members of each group.

.PARAMETER Identity
    The group (alias, email address or name) for -Action AddMember.

.PARAMETER Member
    Users (UPN or email address) to add with -Action AddMember.

.PARAMETER DisplayName
    Display name of the new group for -Action Create.

.PARAMETER Alias
    Alias (mail nickname) of the new group for -Action Create. The address becomes alias@ the
    tenant's default domain.

.PARAMETER AccessType
    Privacy of the new group for -Action Create: Private (default) or Public.

.PARAMETER Apply
    Make the change. Without it the script only reports what it would do.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./manage-microsoft-365-groups.ps1 -TenantId 'contoso.onmicrosoft.com' -IncludeMembers

    Lists every Microsoft 365 Group and its members in one customer.

.EXAMPLE
    ./manage-microsoft-365-groups.ps1 -TenantId 'contoso.onmicrosoft.com' -Action AddMember -Identity 'team' -Member 'jane@contoso.com', 'sam@contoso.com' -Apply -WhatIf

    Shows the members that would be added to the team group. Remove -WhatIf to add them.

.EXAMPLE
    ./manage-microsoft-365-groups.ps1 -TenantId 'contoso.onmicrosoft.com' -Action Create -DisplayName 'Contoso Team' -Alias 'team' -Apply

    Creates a private Microsoft 365 Group called Contoso Team. To remove a group, run
    Remove-UnifiedGroup -Identity team by hand in a Connect-MspExchangeOnline session.

.NOTES
    Replaces the original 2016 method: a list of UnifiedGroup cmdlets (New-UnifiedGroup,
    Get-UnifiedGroup, Add-UnifiedGroupLinks, Get-UnifiedGroupLinks, Remove-UnifiedGroup) run after
    connecting with the retired Basic authentication remote PowerShell guide
    (New-PSSession to outlook.office365.com/powershell-liveid).
    Required GDAP roles: Exchange Recipient Administrator or Exchange Administrator. Global Reader
    is enough for -Action Report.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Manage Microsoft 365 Groups with PowerShell
    https://learn.microsoft.com/en-us/microsoft-365/enterprise/manage-microsoft-365-groups-with-powershell

.LINK
    https://gcit.com.au/knowledge-base/managing-outlook-groups-via-powershell/

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

    [ValidateSet('Report', 'AddMember', 'Create')]
    [string]$Action = 'Report',

    [switch]$IncludeMembers,

    [string]$Identity,

    [string[]]$Member,

    [string]$DisplayName,

    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string]$Alias,

    [ValidateSet('Private', 'Public')]
    [string]$AccessType = 'Private',

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
        foreach ($column in 'GroupName', 'GroupAddress', 'AccessType', 'MemberCount', 'Member', 'Action', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    switch ($Action) {
        'AddMember' {
            if (-not $Identity -or -not $Member) { throw '-Action AddMember needs -Identity and -Member.' }
        }
        'Create' {
            if (-not $DisplayName -or -not $Alias) { throw '-Action Create needs -DisplayName and -Alias.' }
        }
    }
    if ($Action -ne 'Report' -and $AllCustomers) {
        throw "-Action $Action changes one customer at a time. Use -TenantId."
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
                    foreach ($group in @(Get-UnifiedGroup -ResultSize Unlimited -ErrorAction Stop)) {
                        $base = @{
                            GroupName    = $group.DisplayName
                            GroupAddress = [string]$group.PrimarySmtpAddress
                            AccessType   = [string]$group.AccessType
                            MemberCount  = $group.GroupMemberCount
                            Action       = 'Report'
                            Status       = 'Reported'
                        }
                        if (-not $IncludeMembers) {
                            $row = ConvertTo-ResultRow -Customer $customer -Data $base
                            $results.Add($row)
                            $row
                            continue
                        }
                        $links = @(Get-UnifiedGroupLinks -Identity ([string]$group.PrimarySmtpAddress) -LinkType Members -ResultSize Unlimited -ErrorAction Stop)
                        if ($links.Count -eq 0) {
                            $row = ConvertTo-ResultRow -Customer $customer -Data $base
                            $results.Add($row)
                            $row
                        }
                        foreach ($link in $links) {
                            $data = $base.Clone()
                            $data['Member'] = [string]$link.PrimarySmtpAddress
                            $row = ConvertTo-ResultRow -Customer $customer -Data $data
                            $results.Add($row)
                            $row
                        }
                    }
                }
                'AddMember' {
                    $group = Get-UnifiedGroup -Identity $Identity -ErrorAction Stop
                    $groupAddress = [string]$group.PrimarySmtpAddress
                    $current = @(Get-UnifiedGroupLinks -Identity $groupAddress -LinkType Members -ResultSize Unlimited -ErrorAction Stop |
                            ForEach-Object { ([string]$_.PrimarySmtpAddress).ToLowerInvariant() })
                    foreach ($user in $Member) {
                        $data = @{ GroupName = $group.DisplayName; GroupAddress = $groupAddress; AccessType = [string]$group.AccessType; Member = $user; Action = 'AddMember' }
                        if ($current -contains $user.ToLowerInvariant()) {
                            $data['Status'] = 'AlreadyMember'
                        }
                        elseif (-not $Apply) {
                            $data['Status'] = 'WouldAdd'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $groupAddress", "Add member $user")) {
                            try {
                                Add-UnifiedGroupLinks -Identity $groupAddress -LinkType Members -Links $user -Confirm:$false -ErrorAction Stop
                                $data['Status'] = 'Added'
                            }
                            catch {
                                $data['Status'] = 'Failed'
                                $data['Error'] = $_.Exception.Message
                            }
                        }
                        else {
                            $data['Status'] = 'WhatIf'
                        }
                        $row = ConvertTo-ResultRow -Customer $customer -Data $data
                        $results.Add($row)
                        $row
                    }
                }
                'Create' {
                    $data = @{ GroupName = $DisplayName; AccessType = $AccessType; Action = 'Create' }
                    $existing = $null
                    try { $existing = Get-UnifiedGroup -Identity $Alias -ErrorAction Stop } catch { $existing = $null }
                    if ($existing) {
                        $data['GroupAddress'] = [string]$existing.PrimarySmtpAddress
                        $data['Status'] = 'AlreadyExists'
                    }
                    elseif (-not $Apply) {
                        $data['Status'] = 'WouldCreate'
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $DisplayName ($Alias)", 'Create Microsoft 365 Group')) {
                        $created = New-UnifiedGroup -DisplayName $DisplayName -Alias $Alias -AccessType $AccessType -Confirm:$false -ErrorAction Stop
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
