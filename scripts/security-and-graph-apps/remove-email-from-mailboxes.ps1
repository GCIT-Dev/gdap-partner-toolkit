#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds a specific email (for example a phishing message) in every mailbox of a customer tenant with a
    Microsoft Purview content search, and optionally purges it.

.DESCRIPTION
    For each customer named in -TenantId the script connects to Security and Compliance PowerShell with
    Connect-MspSecurityCompliance (confirmed in a live test, see docs/05), then:

    1. builds a keyword query from -Subject, -From, -ReceivedAfter and -ReceivedBefore, or uses
       -ContentMatchQuery as given,
    2. creates and starts a content search across all Exchange mailboxes (New-ComplianceSearch,
       Start-ComplianceSearch) and waits for it to finish, and
    3. reports the number of items, their size and the mailboxes that hold them.

    With -Apply it then purges the results (New-ComplianceSearchAction -Purge) and waits for the purge.
    SoftDelete (the default) moves items to Recoverable Items, where users and admins can still recover them
    until the deleted item retention period ends. A purge removes at most 10 items per mailbox per run, so
    run it again if a mailbox holds more copies. Check the reported count before you purge.

    -WhatIf creates nothing. The search object stays in the customer's Microsoft Purview portal as a record.
    In tenants with Microsoft Defender for Office 365 Plan 2, Threat Explorer remediation is an alternative.
    There is deliberately no -AllCustomers switch.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.

.PARAMETER Subject
    Subject of the message to find (an exact phrase).

.PARAMETER From
    Sender address of the message to find.

.PARAMETER ReceivedAfter
    Only find messages received on or after this date.

.PARAMETER ReceivedBefore
    Only find messages received on or before this date.

.PARAMETER ContentMatchQuery
    A full KQL query. Overrides -Subject, -From, -ReceivedAfter and -ReceivedBefore.

.PARAMETER PurgeType
    SoftDelete (default, recoverable) or HardDelete.

.PARAMETER TimeoutMinutes
    How long to wait for the search and for the purge. Default 15.

.PARAMETER Apply
    Purge the items found. Without it the script only searches and reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./remove-email-from-mailboxes.ps1 -TenantId 'contoso.onmicrosoft.com' -Subject 'Invoice overdue' -From 'billing@fabrikam.example' -ReceivedAfter (Get-Date).AddDays(-2)

    Searches every mailbox in one customer and reports how many copies of the message were found.

.EXAMPLE
    ./remove-email-from-mailboxes.ps1 -TenantId 'contoso.onmicrosoft.com' -Subject 'Invoice overdue' -From 'billing@fabrikam.example' -Apply -WhatIf

    Shows the search and purge that would run, without creating anything.

.NOTES
    Replaces the original 2019 method: an AdminAgents (DAP) multi-tenant app with tenant-wide Mail.ReadWrite
    application permission and a client secret, created with the AzureAD and AzureRM modules, which searched
    and deleted messages in every mailbox through Microsoft Graph.
    Required GDAP roles: Compliance Administrator for the search (or an eDiscovery role), and for -Apply the
    Search And Purge role, which is part of the Microsoft Purview Organization Management role group that
    Global Administrator maps to. Assign that just in time, for the purge only.
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage (used by
    Connect-MspSecurityCompliance), Microsoft Graph delegated User.Read (organisation name lookup).
    Microsoft Learn: https://learn.microsoft.com/en-us/purview/ediscovery-search-for-and-delete-email-messages

.LINK
    https://gcit.com.au/knowledge-base/delete-specific-emails-from-office-365-inboxes-in-customer-tenants-via-powershell/

.LINK
    docs/05-exchange-access.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess)]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [string]$Subject,

    [ValidatePattern('^[^@\s"]+@[^@\s"]+$')]
    [string]$From,

    [datetime]$ReceivedAfter,

    [datetime]$ReceivedBefore,

    [string]$ContentMatchQuery,

    [ValidateSet('SoftDelete', 'HardDelete')]
    [string]$PurgeType = 'SoftDelete',

    [ValidateRange(1, 120)]
    [int]$TimeoutMinutes = 15,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    $query = if ($ContentMatchQuery) {
        $ContentMatchQuery
    }
    else {
        $parts = @()
        if ($Subject) { $parts += 'subject:"{0}"' -f ($Subject -replace '"', '') }
        if ($From) { $parts += 'from:{0}' -f $From }
        if ($PSBoundParameters.ContainsKey('ReceivedAfter')) { $parts += 'received>={0}' -f $ReceivedAfter.ToString('yyyy-MM-dd') }
        if ($PSBoundParameters.ContainsKey('ReceivedBefore')) { $parts += 'received<={0}' -f $ReceivedBefore.ToString('yyyy-MM-dd') }
        $parts -join ' AND '
    }
    if (-not $query -or ($query -notmatch 'subject|from|participants|messageid|internetmessageid' -and -not $ContentMatchQuery)) {
        throw 'Give -Subject or -From (optionally with dates), or a -ContentMatchQuery. A date range alone would match every message.'
    }

    function ConvertTo-ResultRow {
        param($Customer, $SearchName, $Search, $Action, $ErrorMessage)
        $locations = @(([string]$Search.SuccessResults) -split '[\r\n]+' | Where-Object { $_ -match 'Item count: [1-9]' } |
                ForEach-Object { if ($_ -match 'Location: ([^,]+)') { $Matches[1] } })
        [pscustomobject][ordered]@{
            CustomerTenantId  = $Customer.TenantId
            CustomerName      = $Customer.Name
            SearchName        = $SearchName
            Query             = $query
            SearchStatus      = $Search.Status
            Items             = $Search.Items
            SizeBytes         = $Search.Size
            MailboxesWithHits = $locations -join '; '
            PurgeType         = $PurgeType
            Action            = $Action
            Error             = $ErrorMessage
        }
    }

    function Wait-Completion {
        param([scriptblock]$Read, [int]$Minutes)
        $deadline = [datetime]::UtcNow.AddMinutes($Minutes)
        do {
            $state = & $Read
            if ([string]$state.Status -eq 'Completed') { return $state }
            Start-Sleep -Seconds 15
        } while ([datetime]::UtcNow -lt $deadline)
        $state
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })

    foreach ($customer in $customers) {
        $searchName = 'MspGdap-search-{0}' -f [datetime]::UtcNow.ToString('yyyyMMdd-HHmmss')
        try {
            $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
            $customer.Name = $organisation.displayName
            if ($organisation.id) { $customer.TenantId = $organisation.id }

            if (-not $PSCmdlet.ShouldProcess($customer.Name, "Create and run content search '$searchName' for: $query")) {
                $row = ConvertTo-ResultRow -Customer $customer -SearchName $searchName -Action 'WhatIf'
                $results.Add($row)
                $row
                continue
            }

            $null = Connect-MspSecurityCompliance -TenantId $customer.TenantId
            try {
                $null = New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $query -ErrorAction Stop
                Start-ComplianceSearch -Identity $searchName -ErrorAction Stop
                $search = Wait-Completion -Minutes $TimeoutMinutes -Read { Get-ComplianceSearch -Identity $searchName -ErrorAction Stop }

                if ([string]$search.Status -ne 'Completed') {
                    $action = 'SearchTimedOut'
                }
                elseif ([int64]$search.Items -eq 0) {
                    $action = 'NothingFound'
                }
                elseif (-not $Apply) {
                    $action = 'Found'
                }
                elseif ($PSCmdlet.ShouldProcess($customer.Name, "Purge ($PurgeType) $($search.Items) item(s) found by '$searchName'")) {
                    $null = New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType $PurgeType -Confirm:$false -ErrorAction Stop
                    $purge = Wait-Completion -Minutes $TimeoutMinutes -Read { Get-ComplianceSearchAction -Identity "$($searchName)_Purge" -ErrorAction Stop }
                    $action = if ([string]$purge.Status -eq 'Completed') { 'Purged' } else { 'PurgeNotConfirmed' }
                }
                else {
                    $action = 'Found'
                }
                $row = ConvertTo-ResultRow -Customer $customer -SearchName $searchName -Search $search -Action $action
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            $results.Add($row)
            $row
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -SearchName $searchName -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
