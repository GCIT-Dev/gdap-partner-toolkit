#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds licensed shared mailboxes in GDAP customers and optionally removes licences that are safe to remove.
.DESCRIPTION
    For each customer the script:
      1. connects to Exchange Online with the technician's delegated GDAP token and lists shared
         mailboxes (Get-EXOMailbox) with their archive and hold settings,
      2. reads each shared mailbox's Microsoft Entra account through Microsoft Graph to see which
         licences it holds and whether each one is assigned directly or inherited from a group,
      3. reads the mailbox size (Get-EXOMailboxStatistics).

    A shared mailbox still needs a licence when it is over 50 GB, has an archive, or is on litigation
    hold, an in-place or retention hold. Removing the licence from such a mailbox can stop mail
    flow or drop the hold. The script marks those mailboxes SafeToRemove = False and lists the
    reasons. Group-assigned licences can't be removed from the user, so they are reported too.

    Without -Apply the script only reports. With -Apply it removes the directly assigned licences
    from mailboxes marked SafeToRemove = True, through POST /users/{id}/assignLicense, and reads the
    account back. The ReAddSkuIds column lists the SKU IDs removed, so you can add them back with
    assignLicense if needed.
.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input.
.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer -IncludeGdapStatus. Customers without an
    active GDAP relationship are reported as Skipped.
.PARAMETER MaxUnlicensedSizeGB
    Size limit for an unlicensed shared mailbox. Microsoft's limit is 50 GB. Mailboxes at or above
    this size are never marked safe to remove.
.PARAMETER Apply
    Remove directly assigned licences from shared mailboxes marked SafeToRemove. Without this switch
    the script only reports. Combine with -WhatIf to preview the changes.
.PARAMETER OutputPath
    Optional path of a CSV file for the results.
.EXAMPLE
    ./find-licensed-shared-mailboxes.ps1 -AllCustomers -OutputPath ./licensed-shared-mailboxes.csv

    Reports every licensed shared mailbox across all GDAP customers, with the reasons any licence must stay.
.EXAMPLE
    ./find-licensed-shared-mailboxes.ps1 -TenantId 'contoso.onmicrosoft.com' -Apply -WhatIf

    Shows which licences would be removed in one customer, without changing anything.
.NOTES
    Replaces the original 2017 method: MSOnline and DAP (Connect-MsolService, Get-MsolPartnerContract,
    Get-MsolUser, Set-MsolUserLicense) with basic authentication remote PowerShell (Invoke-Command to
    outlook.office365.com/powershell-liveid?DelegatedOrg=), which did not support MFA.
    Required GDAP roles: Global Reader for report-only runs (Exchange Administrator alone can't
    read users and licences through Microsoft Graph, so pair it with Directory Readers). License
    Administrator or User Administrator as well to remove licences.
    Required partner app permissions: Office 365 Exchange Online Exchange.Manage (delegated).
    Microsoft Graph User.ReadWrite.All (delegated, for reading users and assignLicense) and
    Directory.ReadWrite.All or Organization.Read.All (delegated, for GET /subscribedSkus). The least
    privileged Graph permissions are User.Read.All, LicenseAssignment.Read.All and
    LicenseAssignment.ReadWrite.All, which the partner app manifests don't list but the permissions
    above cover.
.LINK
    https://gcit.com.au/knowledge-base/find-remove-unnecessary-licenses-shared-mailboxes-office-365-customer-tenants/
.LINK
    ../../docs/05-exchange-access.md
.LINK
    ../../docs/07-migrating-from-dap-msonline.md
.LINK
    https://learn.microsoft.com/en-us/exchange/collaboration-exo/shared-mailboxes
.LINK
    https://learn.microsoft.com/en-us/graph/api/user-assignlicense?view=graph-rest-1.0
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant', SupportsShouldProcess, ConfirmImpact = 'High')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [ValidateRange(1, 100)]
    [int]$MaxUnlicensedSizeGB = 50,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $ErrorActionPreference = 'Stop'
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $resultColumns = @(
        'Status', 'DisplayName', 'PrimarySmtpAddress', 'UserPrincipalName', 'Licences', 'GroupAssignedLicences',
        'MailboxSizeGB', 'ArchiveActive', 'LitigationHoldEnabled', 'InPlaceHoldCount', 'RetentionHoldEnabled',
        'SafeToRemove', 'BlockingReasons', 'Action', 'ReAddSkuIds', 'Detail'
    )

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

    function ConvertTo-ByteCount {
        # Exchange returns sizes as text such as '1.2 GB (1,288,490,189 bytes)'.
        param([object]$Size)
        if ($null -eq $Size) { return $null }
        $match = [regex]::Match([string]$Size, '\(([\d,]+) bytes\)')
        if ($match.Success) { return [int64]($match.Groups[1].Value -replace ',', '') }
        $null
    }
}

process {
    foreach ($id in $TenantId) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $limitBytes = [int64]$MaxUnlicensedSizeGB * 1GB

    foreach ($customer in @(Get-TargetCustomer -All:$AllCustomers -Requested $requested.ToArray())) {
        if ($customer.PSObject.Properties['GdapStatus'] -and $customer.GdapStatus -ne 'active') {
            $row = ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Skipped'; Detail = "No active GDAP relationship (status: $($customer.GdapStatus))." }
            $results.Add($row)
            $row
            continue
        }

        $customerRows = [System.Collections.Generic.List[object]]::new()
        try {
            Write-Verbose "Checking shared mailbox licences for $($customer.DisplayName)"
            $skuNames = @{}
            foreach ($sku in @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber')) {
                $skuNames[[string]$sku.skuId] = [string]$sku.skuPartNumber
            }

            $null = Connect-MspExchangeOnline -TenantId $customer.TenantId
            $properties = 'ExternalDirectoryObjectId', 'ArchiveStatus', 'LitigationHoldEnabled', 'InPlaceHolds', 'RetentionHoldEnabled'
            $sharedMailboxes = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties $properties)

            foreach ($mailbox in $sharedMailboxes) {
                $values = @{ DisplayName = [string]$mailbox.DisplayName; PrimarySmtpAddress = [string]$mailbox.PrimarySmtpAddress; Status = 'Succeeded' }
                try {
                    $userId = [string]$mailbox.ExternalDirectoryObjectId
                    if (-not $userId) { continue }
                    $user = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri "v1.0/users/$($userId)?`$select=id,userPrincipalName,assignedLicenses,licenseAssignmentStates"
                    $assigned = @($user.assignedLicenses)
                    if ($assigned.Count -eq 0) { continue }

                    $states = @($user.licenseAssignmentStates)
                    $groupSkus = @($states | Where-Object { $_.assignedByGroup } | ForEach-Object { [string]$_.skuId } | Sort-Object -Unique)
                    $directSkus = @($states | Where-Object { -not $_.assignedByGroup } | ForEach-Object { [string]$_.skuId } | Sort-Object -Unique)
                    if ($states.Count -eq 0) { $directSkus = @($assigned | ForEach-Object { [string]$_.skuId }) }

                    $statistics = Get-EXOMailboxStatistics -Identity $userId
                    $sizeBytes = ConvertTo-ByteCount -Size $statistics.TotalItemSize
                    $holds = @($mailbox.InPlaceHolds | Where-Object { $_ })

                    $reasons = [System.Collections.Generic.List[string]]::new()
                    if ($null -eq $sizeBytes) { $reasons.Add('Mailbox size unknown') }
                    elseif ($sizeBytes -ge $limitBytes) { $reasons.Add("Mailbox is $MaxUnlicensedSizeGB GB or larger") }
                    if ([string]$mailbox.ArchiveStatus -eq 'Active') { $reasons.Add('Archive mailbox is on') }
                    if ($mailbox.LitigationHoldEnabled) { $reasons.Add('Litigation hold is on') }
                    if ($holds.Count -gt 0) { $reasons.Add('In-place or retention policy hold applies') }
                    if ($mailbox.RetentionHoldEnabled) { $reasons.Add('Retention hold is on') }
                    if ($directSkus.Count -eq 0) { $reasons.Add('Licences are only assigned through groups') }

                    $values.UserPrincipalName = [string]$user.userPrincipalName
                    $values.Licences = ($assigned | ForEach-Object { $skuId = [string]$_.skuId; if ($skuNames.ContainsKey($skuId)) { $skuNames[$skuId] } else { $skuId } }) -join ', '
                    $values.GroupAssignedLicences = ($groupSkus | ForEach-Object { if ($skuNames.ContainsKey($_)) { $skuNames[$_] } else { $_ } }) -join ', '
                    $values.MailboxSizeGB = if ($null -ne $sizeBytes) { [math]::Round($sizeBytes / 1GB, 2) } else { $null }
                    $values.ArchiveActive = [string]$mailbox.ArchiveStatus -eq 'Active'
                    $values.LitigationHoldEnabled = [bool]$mailbox.LitigationHoldEnabled
                    $values.InPlaceHoldCount = $holds.Count
                    $values.RetentionHoldEnabled = [bool]$mailbox.RetentionHoldEnabled
                    $values.SafeToRemove = $reasons.Count -eq 0
                    $values.BlockingReasons = $reasons -join ', '
                    $values.Action = if ($reasons.Count -eq 0) { 'ReportOnly' } else { 'KeepLicence' }

                    if ($Apply -and $values.SafeToRemove) {
                        if ($PSCmdlet.ShouldProcess("$($customer.DisplayName) shared mailbox $($values.UserPrincipalName)", "Remove licences $($values.Licences)")) {
                            $body = @{ addLicenses = @(); removeLicenses = @($directSkus) }
                            $null = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri "v1.0/users/$userId/assignLicense" -Body $body -Confirm:$false
                            # Read back the direct assignments only. A SKU that is also assigned through a
                            # group stays in assignedLicenses after its direct assignment is removed.
                            $check = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri "v1.0/users/$($userId)?`$select=id,assignedLicenses,licenseAssignmentStates"
                            $remaining = @($check.licenseAssignmentStates | Where-Object { -not $_.assignedByGroup } | ForEach-Object { [string]$_.skuId } | Where-Object { $directSkus -contains $_ })
                            $values.Action = if ($remaining.Count -eq 0) { 'LicenceRemoved' } else { 'ChangeNotConfirmed' }
                            $values.ReAddSkuIds = $directSkus -join ', '
                            if ($remaining.Count -gt 0) { $values.Status = 'Failed' }
                        }
                        else {
                            $values.Action = 'WhatIf'
                        }
                    }
                }
                catch {
                    $values.Status = 'Failed'
                    $values.Detail = $_.Exception.Message
                }
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values $values))
            }

            if ($customerRows.Count -eq 0) {
                $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Succeeded'; Action = 'None'; Detail = "No licensed shared mailboxes found in $($sharedMailboxes.Count) shared mailboxes." }))
            }
        }
        catch {
            $customerRows.Add((ConvertTo-ResultRow -Customer $customer -Values @{ Status = 'Failed'; Detail = $_.Exception.Message }))
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
