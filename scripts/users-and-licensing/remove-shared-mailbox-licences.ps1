#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Finds shared mailboxes that have licences assigned, flags the ones that still need a licence, and optionally removes the rest.

.DESCRIPTION
    Shared mailboxes usually don't need a licence, so a licence on one is often
    wasted money. Some shared mailboxes do need one, and removing it can break
    them. Microsoft requires a licence on a shared mailbox when it:
    - is larger than 50 GB,
    - has an online archive,
    - is on litigation hold or another hold.

    For each customer the script reads the shared mailboxes through Exchange
    Online (Get-EXOMailbox), checks each one's licences through Microsoft Graph and
    its size with Get-EXOMailboxStatistics, and returns one row per licensed shared
    mailbox. KeepLicenceReason explains why a licence must stay. RestoreCommand is a
    ready-made command that puts the licences back if you need to.

    By default the script only reports. With -Apply it removes the directly
    assigned licences from licensed shared mailboxes that have no KeepLicenceReason,
    one ShouldProcess confirmation per mailbox (supports -WhatIf), and reads the
    user back to confirm. Licences inherited from a group are never removed here.
    Remove the mailbox from the licensing group instead.

    The Exchange Online session is closed after each customer.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as
    contoso.onmicrosoft.com. Accepts pipeline input, including objects from
    Get-MspCustomer.

.PARAMETER AllCustomers
    Runs against every customer with an active GDAP relationship, as returned by
    Get-MspCustomer -IncludeGdapStatus.

.PARAMETER SizeWarningGB
    Mailboxes at or above this size (default 45 GB) are kept licensed, to leave
    headroom under the 50 GB limit for unlicensed shared mailboxes.

.PARAMETER Apply
    Removes licences from the shared mailboxes that are safe to unlicense. Without
    it the script only reports.

.PARAMETER OutputPath
    Optional path of a CSV file for the results. The file is overwritten.

.EXAMPLE
    ./remove-shared-mailbox-licences.ps1 -TenantId 'contoso.onmicrosoft.com' -OutputPath ./contoso-shared-mailbox-licences.csv

    Reports licensed shared mailboxes and saves the restore commands.

.EXAMPLE
    ./remove-shared-mailbox-licences.ps1 -AllCustomers -Apply -WhatIf

    Shows which licences would be removed in every active GDAP customer, without removing any.

.NOTES
    Replaces the original 2018 method: Connect-MsolService -Credential (no MFA), Get-MsolUser, Set-MsolUserLicense -RemoveLicenses (MSOnline module, retired 30 May 2025) and Exchange Online remote PowerShell with basic authentication (Invoke-Command against powershell-liveid?DelegatedOrg=).
    Required GDAP roles: Exchange Recipient Administrator or Global Reader to report, plus License Administrator or User Administrator for -Apply.
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage, Microsoft Graph delegated User.Read.All and LicenseAssignment.Read.All to report, LicenseAssignment.ReadWrite.All for -Apply (covered in manifests/partner-app.full.json by Exchange.Manage, User.ReadWrite.All and Directory.ReadWrite.All).

.LINK
    https://gcit.com.au/knowledge-base/find-remove-unnecessary-licenses-shared-mailboxes-office-365-tenant/

.LINK
    https://learn.microsoft.com/en-us/microsoft-365/admin/email/about-shared-mailboxes

.LINK
    docs/05-exchange-access.md
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'AllCustomers')]
    [switch]$AllCustomers,

    [ValidateRange(1, 100)]
    [int]$SizeWarningGB = 45,

    [switch]$Apply,

    [ValidateNotNullOrEmpty()]
    [string]$OutputPath
)

begin {
    $columns = @('CustomerTenantId', 'CustomerName', 'DisplayName', 'PrimarySmtpAddress', 'UserPrincipalName', 'UserId', 'Licences', 'GroupAssignedLicences', 'SizeGB', 'ArchiveStatus', 'LitigationHoldEnabled', 'InPlaceHoldCount', 'KeepLicenceReason', 'RestoreCommand', 'Action', 'Status', 'Error')
    $targets = [System.Collections.Generic.List[string]]::new()
    $knownNames = @{}
    $results = [System.Collections.Generic.List[object]]::new()

    function Get-ResultRow {
        param(
            [Parameter(Mandatory)][string[]]$Column,
            [Parameter(Mandatory)][System.Collections.IDictionary]$Value
        )
        $row = [ordered]@{}
        foreach ($name in $Column) {
            $row[$name] = if ($Value.Contains($name)) { $Value[$name] } else { $null }
        }
        [pscustomobject]$row
    }

    function ConvertTo-SizeGB {
        param($Size)
        $text = [string]$Size
        if ($text -match '\(([\d,]+) bytes\)') {
            return [math]::Round(([double]($Matches[1] -replace ',', '')) / 1GB, 2)
        }
        return $null
    }
}

process {
    if ($PSCmdlet.ParameterSetName -eq 'Tenant') {
        foreach ($item in $TenantId) { $targets.Add($item) }
    }
}

end {
    if ($AllCustomers) {
        foreach ($customer in @(Get-MspCustomer -IncludeGdapStatus | Where-Object { $_.GdapStatus -eq 'active' })) {
            $targets.Add($customer.TenantId)
            $knownNames[$customer.TenantId] = $customer.DisplayName
        }
    }

    foreach ($target in $targets) {
        $tenant = $target
        $customerName = $null
        try {
            $tenant = Resolve-MspTenantId -Tenant $target
            $customerName = $knownNames[$tenant]
            if (-not $customerName) {
                $customerName = @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/organization?$select=displayName')[0].displayName
            }

            $skuNames = @{}
            foreach ($sku in @(Invoke-MspGraphRequest -TenantId $tenant -Uri 'v1.0/subscribedSkus?$select=skuId,skuPartNumber')) {
                $skuNames[[string]$sku.skuId] = $sku.skuPartNumber
            }

            Connect-MspExchangeOnline -TenantId $tenant | Out-Null
            try {
                $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -Properties ArchiveStatus, LitigationHoldEnabled, InPlaceHolds, ExternalDirectoryObjectId -ErrorAction Stop)
                foreach ($mailbox in $mailboxes) {
                    $userId = [string]$mailbox.ExternalDirectoryObjectId
                    if (-not $userId) { continue }
                    try {

                        $user = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/users/{0}?$select=id,userPrincipalName,assignedLicenses,licenseAssignmentStates' -f $userId)
                        $assigned = @($user.assignedLicenses | Where-Object { $_ })
                        if ($assigned.Count -eq 0) { continue }

                        $directSkus = [System.Collections.Generic.List[string]]::new()
                        $groupSkus = [System.Collections.Generic.List[string]]::new()
                        foreach ($licence in $assigned) {
                            $skuId = [string]$licence.skuId
                            $states = @($user.licenseAssignmentStates | Where-Object { $_ -and [string]$_.skuId -eq $skuId })
                            if ($states.Count -eq 0 -or @($states | Where-Object { -not $_.assignedByGroup }).Count -gt 0) { $directSkus.Add($skuId) }
                            if (@($states | Where-Object { $_.assignedByGroup }).Count -gt 0) { $groupSkus.Add($skuId) }
                        }

                        $stats = if ($mailbox.ExchangeGuid) {
                            Get-EXOMailboxStatistics -ExchangeGuid ([string]$mailbox.ExchangeGuid) -ErrorAction Stop
                        }
                        else {
                            Get-EXOMailboxStatistics -Identity ([string]$mailbox.UserPrincipalName) -ErrorAction Stop
                        }
                        $sizeGB = ConvertTo-SizeGB -Size $stats.TotalItemSize
                        $holdCount = @($mailbox.InPlaceHolds | Where-Object { $_ }).Count

                        $reasons = [System.Collections.Generic.List[string]]::new()
                        if ($null -eq $sizeGB) { $reasons.Add('Size unknown') }
                        elseif ($sizeGB -ge $SizeWarningGB) { $reasons.Add("Size $sizeGB GB") }
                        if ([string]$mailbox.ArchiveStatus -eq 'Active') { $reasons.Add('Online archive') }
                        if ($mailbox.LitigationHoldEnabled) { $reasons.Add('Litigation hold') }
                        if ($holdCount -gt 0) { $reasons.Add("$holdCount other hold(s)") }

                        $restoreList = ($directSkus | ForEach-Object { "@{ skuId = '$_' }" }) -join ', '
                        $values = [ordered]@{
                            CustomerTenantId      = $tenant
                            CustomerName          = $customerName
                            DisplayName           = $mailbox.DisplayName
                            PrimarySmtpAddress    = $mailbox.PrimarySmtpAddress
                            UserPrincipalName     = $user.userPrincipalName
                            UserId                = $userId
                            Licences              = ($directSkus | ForEach-Object { if ($skuNames.ContainsKey($_)) { $skuNames[$_] } else { $_ } }) -join ', '
                            GroupAssignedLicences = ($groupSkus | ForEach-Object { if ($skuNames.ContainsKey($_)) { $skuNames[$_] } else { $_ } }) -join ', '
                            SizeGB                = $sizeGB
                            ArchiveStatus         = $mailbox.ArchiveStatus
                            LitigationHoldEnabled = $mailbox.LitigationHoldEnabled
                            InPlaceHoldCount      = $holdCount
                            KeepLicenceReason     = $reasons -join ', '
                            RestoreCommand        = if ($directSkus.Count -gt 0) { "Invoke-MspGraphRequest -TenantId '$tenant' -Method POST -Uri 'v1.0/users/$userId/assignLicense' -Body @{ addLicenses = @($restoreList); removeLicenses = @() }" } else { $null }
                            Action                = 'None'
                            Status                = 'OK'
                        }

                        if ($reasons.Count -gt 0) {
                            $values.Action = 'KeepLicence'
                        }
                        elseif ($directSkus.Count -eq 0) {
                            $values.Action = 'GroupAssignedOnly'
                        }
                        elseif (-not $Apply) {
                            $values.Action = 'WouldRemove'
                        }
                        elseif ($PSCmdlet.ShouldProcess("$($mailbox.PrimarySmtpAddress) in $customerName ($tenant)", "Remove licences $($values.Licences)")) {
                            try {
                                $body = @{ addLicenses = @(); removeLicenses = @($directSkus) }
                                Invoke-MspGraphRequest -TenantId $tenant -Method POST -Uri "v1.0/users/$userId/assignLicense" -Body $body -Confirm:$false | Out-Null
                                $check = Invoke-MspGraphRequest -TenantId $tenant -Uri ('v1.0/users/{0}?$select=assignedLicenses' -f $userId)
                                $left = @($check.assignedLicenses | Where-Object { $_ -and $directSkus -contains [string]$_.skuId })
                                $values.Action = if ($left.Count -eq 0) { 'Removed' } else { 'Unconfirmed' }
                                if ($left.Count -gt 0) { $values.Status = 'Warning' }
                            }
                            catch {
                                $values.Action = 'RemoveFailed'
                                $values.Status = 'Failed'
                                $values.Error = $_.Exception.Message
                            }
                        }
                        else {
                            $values.Action = 'WhatIf'
                        }

                        $row = Get-ResultRow -Column $columns -Value $values
                        $results.Add($row)
                        $row
                    }
                    catch {
                        Write-Warning -Message "Mailbox $($mailbox.PrimarySmtpAddress) in $tenant failed: $($_.Exception.Message)"
                        $row = Get-ResultRow -Column $columns -Value @{
                            CustomerTenantId   = $tenant
                            CustomerName       = $customerName
                            DisplayName        = $mailbox.DisplayName
                            PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                            UserId             = $userId
                            Status             = 'Failed'
                            Error              = $_.Exception.Message
                        }
                        $results.Add($row)
                        $row
                    }
                }
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Warning -Message "Customer $target failed: $($_.Exception.Message)"
            $row = Get-ResultRow -Column $columns -Value @{
                CustomerTenantId = $tenant
                CustomerName     = $customerName
                Status           = 'Failed'
                Error            = $_.Exception.Message
            }
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath -and $results.Count -gt 0) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8 -WhatIf:$false -Confirm:$false
    }
}
