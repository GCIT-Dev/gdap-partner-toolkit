#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally creates and assigns, Exchange ActiveSync mobile device mailbox policies.

.DESCRIPTION
    Mobile device mailbox policies apply to devices that sync mail over Exchange ActiveSync. Many
    of their Allow settings (camera, Wi-Fi, tethering, browser) are only honoured by some native
    mail apps and are ignored by Outlook for iOS and Android. For business phones and tablets,
    Microsoft Intune (app protection and compliance policies) with Conditional Access is the
    recommended control. Use this script to audit what is assigned today, or to keep an existing
    EAS policy setup consistent.

    The script connects to Exchange Online in each customer through GDAP with
    Connect-MspExchangeOnline.

    Without -PolicyName it is a report: one row per policy with its settings, and one row per
    mailbox with its assigned policy.

    With -PolicyName it plans to create that policy (if it does not exist, using any Allow settings
    you pass) and, with -AssignTo or -AssignToAllMailboxes, to assign it with
    Set-CASMailbox -ActiveSyncMailboxPolicy. Nothing changes unless you add -Apply. An existing
    policy is never modified. Use -WhatIf with -Apply to preview.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER PolicyName
    Name of the mobile device mailbox policy to create (if missing) and assign.

.PARAMETER AllowCamera
    Setting for a new policy. $false blocks the device camera on clients that honour it.

.PARAMETER AllowWiFi
    Setting for a new policy. $false blocks Wi-Fi on clients that honour it.

.PARAMETER AllowInternetSharing
    Setting for a new policy. $false blocks tethering on clients that honour it.

.PARAMETER AllowBrowser
    Setting for a new policy. $false blocks the built-in browser on clients that honour it.

.PARAMETER AssignTo
    Mailboxes (UPN or primary SMTP address) to assign -PolicyName to.

.PARAMETER AssignToAllMailboxes
    Assign -PolicyName to every mailbox in the tenant.

.PARAMETER Apply
    Make the changes. Without it the script only reports what it would do.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./set-mobile-device-mailbox-policy.ps1 -AllCustomers -OutputPath 'C:\Reports\EasPolicies.csv'

    Reports every mobile device mailbox policy and which mailbox uses which policy, in every customer.

.EXAMPLE
    ./set-mobile-device-mailbox-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -PolicyName 'No Camera Policy' -AllowCamera $false -AssignTo 'jane@contoso.com' -Apply -WhatIf

    Shows the policy creation and assignment without making them. Remove -WhatIf to apply.

.EXAMPLE
    ./set-mobile-device-mailbox-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -PolicyName 'Default' -AssignTo 'jane@contoso.com' -Apply

    Puts a mailbox back on the built-in Default policy, as the original article's revert step did.

.NOTES
    Replaces the original 2016 method: New-MobileDeviceMailboxPolicy and
    Get-Mailbox | Set-CASMailbox -ActiveSyncMailboxPolicy, run after connecting with the retired
    Basic authentication remote PowerShell guide (New-PSSession to outlook.office365.com/powershell-liveid).
    Required GDAP roles: Exchange Administrator (Global Reader is enough for report-only runs).
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online), User.Read
    (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All or
    Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Mobile device mailbox policies in Exchange Online
    https://learn.microsoft.com/en-us/exchange/clients-and-mobile-in-exchange-online/exchange-activesync/mobile-device-mailbox-policies
    Microsoft Learn: App protection policies overview (Intune)
    https://learn.microsoft.com/en-us/intune/intune-service/apps/app-protection-policy

.LINK
    https://gcit.com.au/knowledge-base/how-to-secure-the-devices-that-access-your-office-365-email/

.LINK
    https://gcit.com.au/secure-the-devices-that-access-your-office-365-email/

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

    [ValidateNotNullOrEmpty()]
    [string]$PolicyName,

    [bool]$AllowCamera,

    [bool]$AllowWiFi,

    [bool]$AllowInternetSharing,

    [bool]$AllowBrowser,

    [string[]]$AssignTo,

    [switch]$AssignToAllMailboxes,

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
        foreach ($column in 'RowType', 'PolicyName', 'Mailbox', 'IsDefault', 'AllowCamera', 'AllowWiFi', 'AllowInternetSharing', 'AllowBrowser', 'Action', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    if (($AssignTo -or $AssignToAllMailboxes) -and -not $PolicyName) {
        throw '-AssignTo and -AssignToAllMailboxes need -PolicyName.'
    }
    $settingNames = 'AllowCamera', 'AllowWiFi', 'AllowInternetSharing', 'AllowBrowser'
    $newPolicySettings = @{}
    foreach ($setting in $settingNames) {
        if ($PSBoundParameters.ContainsKey($setting)) { $newPolicySettings[$setting] = $PSBoundParameters[$setting] }
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
            $policies = @(Get-MobileDeviceMailboxPolicy -ErrorAction Stop)

            if (-not $PolicyName) {
                foreach ($policy in $policies) {
                    $data = @{ RowType = 'Policy'; PolicyName = $policy.Name; IsDefault = $policy.IsDefault; Status = 'Reported' }
                    foreach ($setting in $settingNames) { $data[$setting] = $policy.$setting }
                    $row = ConvertTo-ResultRow -Customer $customer -Data $data
                    $results.Add($row)
                    $row
                }
                foreach ($casMailbox in @(Get-EXOCasMailbox -ResultSize Unlimited -Properties ActiveSyncMailboxPolicy, ActiveSyncEnabled -ErrorAction Stop)) {
                    $row = ConvertTo-ResultRow -Customer $customer -Data @{ RowType = 'Mailbox'; PolicyName = [string]$casMailbox.ActiveSyncMailboxPolicy; Mailbox = [string]$casMailbox.PrimarySmtpAddress; Status = 'Reported' }
                    $results.Add($row)
                    $row
                }
                continue
            }

            # 1. Create the policy if it does not exist. An existing policy is left as it is.
            $policyExists = [bool](@($policies | Where-Object { $_.Name -eq $PolicyName }).Count)
            $policyData = @{ RowType = 'Policy'; PolicyName = $PolicyName; Action = 'Create' }
            foreach ($setting in $settingNames) { $policyData[$setting] = $newPolicySettings[$setting] }
            $policyReady = $policyExists
            if ($policyExists) {
                $policyData['Action'] = 'None'
                $policyData['Status'] = 'AlreadyExists'
            }
            elseif (-not $Apply) {
                $policyData['Status'] = 'WouldCreate'
            }
            elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $PolicyName", 'Create mobile device mailbox policy')) {
                try {
                    $null = New-MobileDeviceMailboxPolicy -Name $PolicyName @newPolicySettings -Confirm:$false -ErrorAction Stop
                    $policyData['Status'] = 'Created'
                    $policyReady = $true
                }
                catch {
                    $policyData['Status'] = 'Failed'
                    $policyData['Error'] = $_.Exception.Message
                }
            }
            else {
                $policyData['Status'] = 'WhatIf'
            }
            $row = ConvertTo-ResultRow -Customer $customer -Data $policyData
            $results.Add($row)
            $row

            # 2. Assign it.
            if (-not ($AssignTo -or $AssignToAllMailboxes)) { continue }
            $targets = if ($AssignToAllMailboxes) {
                @(Get-EXOCasMailbox -ResultSize Unlimited -Properties ActiveSyncMailboxPolicy -ErrorAction Stop)
            }
            else {
                @(foreach ($identity in $AssignTo) { Get-EXOCasMailbox -Identity $identity -Properties ActiveSyncMailboxPolicy -ErrorAction Stop })
            }
            foreach ($target in $targets) {
                $address = [string]$target.PrimarySmtpAddress
                $data = @{ RowType = 'Mailbox'; PolicyName = $PolicyName; Mailbox = $address; Action = 'Assign' }
                if ([string]$target.ActiveSyncMailboxPolicy -eq $PolicyName) {
                    $data['Action'] = 'None'
                    $data['Status'] = 'AlreadyAssigned'
                }
                elseif (-not $Apply) {
                    $data['Status'] = 'WouldAssign'
                }
                elseif (-not $policyReady) {
                    $data['Status'] = if ($WhatIfPreference) { 'WhatIf' } else { 'Skipped' }
                    if (-not $WhatIfPreference) { $data['Error'] = 'The policy was not created.' }
                }
                elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $address", "Assign mobile device mailbox policy $PolicyName")) {
                    try {
                        Set-CASMailbox -Identity $address -ActiveSyncMailboxPolicy $PolicyName -Confirm:$false -ErrorAction Stop
                        $data['Status'] = 'Assigned'
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
