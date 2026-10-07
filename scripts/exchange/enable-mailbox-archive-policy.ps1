#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Moves older mail to the online archive: creates a move-to-archive retention policy, assigns it and enables the archive.

.DESCRIPTION
    Use it when a mailbox is near its quota, for example during a migration from Google Workspace,
    where one message with several labels lands in several folders and the mailbox grows quickly.

    For each customer the script connects to Exchange Online through GDAP with
    Connect-MspExchangeOnline and then, for the mailboxes you name:
    1. makes sure a retention tag exists that moves items older than -AgeLimitDays to the archive
       (New-RetentionPolicyTag -Type All -RetentionAction MoveToArchive),
    2. makes sure a retention policy exists that links that tag (New-RetentionPolicy, or
       Set-RetentionPolicy to add the tag to an existing policy of that name),
    3. assigns the policy to each mailbox (Set-Mailbox -RetentionPolicy),
    4. enables the online archive (Enable-Mailbox -Archive), and
    5. starts the Managed Folder Assistant (Start-ManagedFolderAssistant) so items begin to move.

    It is report-only unless you add -Apply, and reports each mailbox's current retention policy,
    archive status and archive size. Use -WhatIf with -Apply to preview.

    If the tenant is dehydrated, Exchange refuses new retention tags until
    Enable-OrganizationCustomization has run. With -Apply the script runs it first, as part of the
    organisation-level change.

    The online archive needs a licence that includes it, such as Exchange Online Plan 2, Microsoft
    365 Business Premium, Microsoft 365 E3 or E5, or the Exchange Online Archiving add-on. Retention
    tags and policies (messaging records management) still work in Exchange Online, and Microsoft
    Purview retention policies are the newer way to retain or delete content.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped. Mailboxes named in -Mailbox that are
    not found in a customer are reported as failed for that customer.

.PARAMETER Mailbox
    Mailboxes (UPN or primary SMTP address) to archive.

.PARAMETER AgeLimitDays
    Age in days after which items move to the archive. Defaults to 365.

.PARAMETER TagName
    Name of the retention tag. Defaults to "Move to archive after <days> days".

.PARAMETER PolicyName
    Name of the retention policy. Defaults to "Archive after <days> days".

.PARAMETER Apply
    Make the changes. Without it the script only reports what it would do.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./enable-mailbox-archive-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -Mailbox 'jane@contoso.com'

    Reports Jane's retention policy, archive status and archive size, and what would change.

.EXAMPLE
    ./enable-mailbox-archive-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -Mailbox 'jane@contoso.com', 'sam@contoso.com' -AgeLimitDays 365 -Apply -WhatIf

    Shows every change without making it. Remove -WhatIf to apply them.

.NOTES
    Replaces the original 2016 method: New-RetentionPolicyTag, New-RetentionPolicy,
    Set-Mailbox -RetentionPolicy, Enable-Mailbox -Archive and Start-ManagedFolderAssistant, run one
    by one after connecting with the retired Basic authentication remote PowerShell guide
    (New-PSSession to outlook.office365.com/powershell-liveid).
    Required GDAP roles: Exchange Administrator (Global Reader is enough for report-only runs).
    Changing an existing policy affects every mailbox that already uses it, so the script only ever
    adds the tag to a policy with the name you give, and never edits Default MRM Policy unless you
    name it.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Enable archive mailboxes
    https://learn.microsoft.com/en-us/purview/enable-archive-mailboxes
    Microsoft Learn: Create and apply retention policies in Exchange Online (MRM)
    https://learn.microsoft.com/en-us/exchange/security-and-compliance/messaging-records-management/create-a-retention-policy

.LINK
    https://gcit.com.au/knowledge-base/working-archive-policies-office-365-skykick/

.LINK
    https://gcit.com.au/working-archive-policies-office-365/

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

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Mailbox,

    [ValidateRange(1, 24855)]
    [int]$AgeLimitDays = 365,

    [ValidateNotNullOrEmpty()]
    [string]$TagName,

    [ValidateNotNullOrEmpty()]
    [string]$PolicyName,

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
        foreach ($column in 'Target', 'CurrentRetentionPolicy', 'ArchiveStatus', 'ArchiveItemCount', 'ArchiveSize', 'Actions', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    if (-not $TagName) { $TagName = "Move to archive after $AgeLimitDays days" }
    if (-not $PolicyName) { $PolicyName = "Archive after $AgeLimitDays days" }
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

            # Organisation level: customisation, retention tag and retention policy.
            $orgActions = [System.Collections.Generic.List[string]]::new()
            $orgData = @{ Target = "Retention policy '$PolicyName' with tag '$TagName'" }
            $orgReady = $true
            $organization = Get-OrganizationConfig -ErrorAction Stop
            $tag = $null
            try { $tag = Get-RetentionPolicyTag -Identity $TagName -ErrorAction Stop } catch { $tag = $null }
            $policy = $null
            try { $policy = Get-RetentionPolicy -Identity $PolicyName -ErrorAction Stop } catch { $policy = $null }

            # An existing policy of that name must also link the tag, or nothing moves to the archive.
            $policyLinksTag = $false
            if ($policy) {
                $linkedTags = @($policy.RetentionPolicyTagLinks | ForEach-Object { [string]$_ })
                $policyLinksTag = $linkedTags -contains $TagName
            }
            if (-not $tag -and $organization.IsDehydrated) { $orgActions.Add('Enable-OrganizationCustomization') }
            if (-not $tag) { $orgActions.Add("New-RetentionPolicyTag ($AgeLimitDays days, MoveToArchive)") }
            if (-not $policy) { $orgActions.Add('New-RetentionPolicy') }
            elseif (-not $policyLinksTag) { $orgActions.Add("Set-RetentionPolicy (add tag '$TagName')") }
            $orgData['Actions'] = $orgActions -join '; '
            if ($tag -and [string]$tag.RetentionAction -and [string]$tag.RetentionAction -ne 'MoveToArchive') {
                $orgData['Error'] = "The existing tag '$TagName' uses RetentionAction $($tag.RetentionAction), not MoveToArchive. Check it or pass a different -TagName."
            }

            if ($orgActions.Count -eq 0) {
                $orgData['Status'] = 'AlreadyInPlace'
            }
            elseif (-not $Apply) {
                $orgData['Status'] = 'WouldChange'
                $orgReady = $false
            }
            elseif ($PSCmdlet.ShouldProcess("$($customer.Name): organisation", $orgData['Actions'])) {
                try {
                    if (-not $tag -and $organization.IsDehydrated) { Enable-OrganizationCustomization -Confirm:$false -ErrorAction Stop }
                    if (-not $tag) {
                        $null = New-RetentionPolicyTag -Name $TagName -Type All -RetentionEnabled $true -AgeLimitForRetention $AgeLimitDays -RetentionAction MoveToArchive -Confirm:$false -ErrorAction Stop
                    }
                    if (-not $policy) {
                        $null = New-RetentionPolicy -Name $PolicyName -RetentionPolicyTagLinks $TagName -Confirm:$false -ErrorAction Stop
                    }
                    elseif (-not $policyLinksTag) {
                        Set-RetentionPolicy -Identity $PolicyName -RetentionPolicyTagLinks @{ Add = $TagName } -Confirm:$false -ErrorAction Stop
                    }
                    $orgData['Status'] = 'Changed'
                }
                catch {
                    $orgData['Status'] = 'Failed'
                    $orgData['Error'] = $_.Exception.Message
                    $orgReady = $false
                }
            }
            else {
                $orgData['Status'] = 'WhatIf'
                $orgReady = $false
            }
            $row = ConvertTo-ResultRow -Customer $customer -Data $orgData
            $results.Add($row)
            $row

            # Mailbox level.
            foreach ($identity in $Mailbox) {
                $data = @{ Target = $identity }
                try {
                    $item = Get-EXOMailbox -Identity $identity -Properties RetentionPolicy, ArchiveStatus, ArchiveGuid -ErrorAction Stop
                    $address = [string]$item.PrimarySmtpAddress
                    $data['Target'] = $address
                    $data['CurrentRetentionPolicy'] = [string]$item.RetentionPolicy
                    $data['ArchiveStatus'] = [string]$item.ArchiveStatus
                    $archiveActive = ([string]$item.ArchiveStatus -eq 'Active')
                    if ($archiveActive) {
                        try {
                            $statistics = Get-EXOMailboxStatistics -Identity $address -Archive -ErrorAction Stop
                            $data['ArchiveItemCount'] = $statistics.ItemCount
                            $data['ArchiveSize'] = [string]$statistics.TotalItemSize
                        }
                        catch {
                            Write-Verbose "Could not read archive statistics for $address. $($_.Exception.Message)"
                        }
                    }

                    $actions = [System.Collections.Generic.List[string]]::new()
                    if ([string]$item.RetentionPolicy -ne $PolicyName) { $actions.Add("Set-Mailbox -RetentionPolicy '$PolicyName'") }
                    if (-not $archiveActive) { $actions.Add('Enable-Mailbox -Archive') }
                    $actions.Add('Start-ManagedFolderAssistant')
                    $data['Actions'] = $actions -join '; '

                    if (-not $Apply) {
                        $data['Status'] = 'WouldChange'
                    }
                    elseif (-not $orgReady) {
                        $data['Status'] = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' }
                        if (-not $WhatIfPreference) { $data['Error'] = 'The retention tag or policy is not in place.' }
                    }
                    elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $address", $data['Actions'])) {
                        if ([string]$item.RetentionPolicy -ne $PolicyName) {
                            Set-Mailbox -Identity $address -RetentionPolicy $PolicyName -Confirm:$false -ErrorAction Stop
                        }
                        if (-not $archiveActive) {
                            $null = Enable-Mailbox -Identity $address -Archive -Confirm:$false -ErrorAction Stop
                        }
                        Start-ManagedFolderAssistant -Identity $address -ErrorAction Stop
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
