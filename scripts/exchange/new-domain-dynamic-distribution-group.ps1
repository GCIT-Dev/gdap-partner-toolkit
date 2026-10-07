#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Previews, and optionally creates, a dynamic distribution group of every user mailbox on one email domain.

.DESCRIPTION
    Connects to Exchange Online in each customer through GDAP with Connect-MspExchangeOnline and
    builds this recipient filter:

        (RecipientTypeDetails -eq 'UserMailbox') -and (WindowsEmailAddress -like '*@<domain>')

    WindowsEmailAddress is the user's primary email address. The original article filtered on
    WindowsLiveID with -eq, which does not handle wildcards. OPATH wildcards need -like.

    The script always previews the members first with Get-Recipient -RecipientPreviewFilter. It is
    report-only unless you add -Apply. With -Apply it creates the group with
    New-DynamicDistributionGroup, unless a group with that name already exists. Use -WhatIf with
    -Apply to preview.

    Exchange Online recalculates dynamic distribution group membership on a schedule, so a new
    group can take a while to show its members in Get-DynamicDistributionGroupMember.

    Microsoft's filter reference says a leading wildcard ('*@...') isn't supported in some
    Exchange Online filters and isn't recommended for performance. The domain match needs one, so
    always check the preview count. If the preview fails or is empty when it shouldn't be, set a
    custom attribute (for example CustomAttribute1 to the domain) and filter on that instead.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER Domain
    The email domain to group by, for example contoso.com.

.PARAMETER Name
    Name of the dynamic distribution group. Defaults to "All Users - <domain>".

.PARAMETER PrimarySmtpAddress
    Optional email address for the new group. Exchange generates one from the name when omitted.

.PARAMETER Apply
    Create the group. Without it the script only previews the members.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./new-domain-dynamic-distribution-group.ps1 -TenantId 'contoso.onmicrosoft.com' -Domain 'contoso.com'

    Shows how many user mailboxes, and which ones, the group would contain.

.EXAMPLE
    ./new-domain-dynamic-distribution-group.ps1 -TenantId 'contoso.onmicrosoft.com' -Domain 'contoso.com' -PrimarySmtpAddress 'allstaff@contoso.com' -Apply -WhatIf

    Shows the New-DynamicDistributionGroup call without making it. Remove -WhatIf to create the group.

.NOTES
    Replaces the original 2020 method: New-DynamicDistributionGroup with a WindowsLiveID -eq filter
    (the snippet also had an unclosed quote), run after connecting with the retired Basic
    authentication remote PowerShell guide.
    Required GDAP roles: Exchange Recipient Administrator or Exchange Administrator.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Filterable properties for the RecipientFilter parameter
    https://learn.microsoft.com/en-us/powershell/exchange/recipientfilter-properties

.LINK
    https://gcit.com.au/knowledge-base/create-dynamic-distribution-list-based-on-email-domain/

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
    [ValidatePattern('^[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+$')]
    [string]$Domain,

    [ValidateNotNullOrEmpty()]
    [string]$Name,

    [ValidatePattern('^[^@\s]+@[^@\s]+$')]
    [string]$PrimarySmtpAddress,

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
        foreach ($column in 'GroupName', 'RecipientFilter', 'PreviewMemberCount', 'PreviewMembers', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    $domainLower = $Domain.ToLowerInvariant()
    if (-not $Name) { $Name = "All Users - $domainLower" }
    $recipientFilter = "(RecipientTypeDetails -eq 'UserMailbox') -and (WindowsEmailAddress -like '*@$domainLower')"
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
            $row = ConvertTo-ResultRow -Customer $customer -Data @{ GroupName = $Name; RecipientFilter = $recipientFilter; Status = 'Skipped'; Error = $customer.Skip }
            $results.Add($row)
            $row
            continue
        }

        $connection = $null
        $data = @{ GroupName = $Name; RecipientFilter = $recipientFilter }
        try {
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop

            $preview = @(Get-Recipient -RecipientPreviewFilter $recipientFilter -ResultSize Unlimited -ErrorAction Stop)
            $data['PreviewMemberCount'] = $preview.Count
            $data['PreviewMembers'] = ($preview | ForEach-Object { [string]$_.PrimarySmtpAddress }) -join '; '

            $existing = $null
            try { $existing = Get-DynamicDistributionGroup -Identity $Name -ErrorAction Stop }
            catch { $existing = $null }

            if ($existing) {
                $data['Status'] = 'AlreadyExists'
                # Exchange adds its own system exclusions to a stored filter, so only check the domain clause.
                if ([string]$existing.RecipientFilter -notlike "*@$domainLower*") {
                    $data['Error'] = "A group with this name exists but its filter does not mention @$($domainLower): $($existing.RecipientFilter)"
                }
            }
            elseif (-not $Apply) {
                $data['Status'] = 'WouldCreate'
            }
            elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $Name", "Create dynamic distribution group with filter $recipientFilter")) {
                $newGroup = @{ Name = $Name; RecipientFilter = $recipientFilter; ErrorAction = 'Stop' }
                if ($PrimarySmtpAddress) { $newGroup['PrimarySmtpAddress'] = $PrimarySmtpAddress }
                $null = New-DynamicDistributionGroup @newGroup
                $data['Status'] = 'Created'
            }
            else {
                $data['Status'] = 'WhatIf'
            }
        }
        catch {
            Write-Warning "$($customer.Name) ($($customer.TenantId)) failed: $($_.Exception.Message)"
            $data['Status'] = 'Failed'
            $data['Error'] = $_.Exception.Message
        }
        finally {
            Close-CustomerSession -Connection $connection
        }
        $row = ConvertTo-ResultRow -Customer $customer -Data $data
        $results.Add($row)
        $row
    }

    if ($OutputPath) {
        $folder = Split-Path -Path $OutputPath -Parent
        if ($folder -and -not (Test-Path -LiteralPath $folder)) { $null = New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false }
        $results | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false
        Write-Verbose "Saved $($results.Count) row(s) to $OutputPath."
    }
}
