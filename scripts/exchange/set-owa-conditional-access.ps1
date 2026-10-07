#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports, and optionally configures, limited Outlook on the web access (no attachment downloads) on unmanaged devices.

.DESCRIPTION
    Outlook on the web can stop people downloading attachments on devices that are not managed or
    compliant, while still letting them read mail in the browser. Two settings work together:
    - a Microsoft Entra Conditional Access policy for Office 365 Exchange Online (or Office 365)
      with the session control "Use app enforced restrictions", and
    - the ConditionalAccessPolicy setting of the OWA mailbox policy (ReadOnly or
      ReadOnlyPlusAttachmentsBlocked).

    The script reads both in each customer: the OWA mailbox policies through Exchange Online
    (Connect-MspExchangeOnline) and the Conditional Access policies through Microsoft Graph
    (Invoke-MspGraphRequest). It is report-only unless you add -Apply.

    With -Apply it sets -OwaRestriction on the OWA mailbox policies named in -OwaMailboxPolicy
    (Set-OwaMailboxPolicy). With -Apply and -CreateConditionalAccessPolicy it also creates a
    Conditional Access policy for browser access to Office 365 Exchange Online with app enforced
    restrictions, in report-only state (enabledForReportingButNotEnforced) so you can check its
    impact in the sign-in logs before you turn it on. Use -WhatIf with -Apply to preview.

    Conditional Access needs Microsoft Entra ID P1 or P2 (included in Microsoft 365 Business
    Premium, E3 and E5).

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains, such as contoso.onmicrosoft.com.
    Accepts pipeline input, including objects from Get-MspCustomer.

.PARAMETER AllCustomers
    Run against every customer returned by Get-MspCustomer that has an active GDAP relationship.
    Customers without one are listed with the status Skipped.

.PARAMETER OwaRestriction
    ReadOnly (default): no downloads, attachments can be viewed in the browser.
    ReadOnlyPlusAttachmentsBlocked: attachments can't be opened at all. Off: no restriction.

.PARAMETER OwaMailboxPolicy
    OWA mailbox policies to change with -Apply. Defaults to OwaMailboxPolicy-Default.

.PARAMETER CreateConditionalAccessPolicy
    With -Apply, create the Conditional Access policy (report-only state) if no policy with
    app enforced restrictions for Exchange Online or Office 365 exists.

.PARAMETER ConditionalAccessPolicyName
    Display name of the new Conditional Access policy.

.PARAMETER ExcludeGroupId
    Object IDs of groups to exclude from the new Conditional Access policy, such as your
    emergency access accounts group.

.PARAMETER Apply
    Make the changes. Without it the script only reports.

.PARAMETER OutputPath
    Optional path of a CSV file to save the results to. The folder is created if needed.

.EXAMPLE
    ./set-owa-conditional-access.ps1 -AllCustomers -OutputPath 'C:\Reports\OwaConditionalAccess.csv'

    Reports the OWA restriction and any matching Conditional Access policies in every customer.

.EXAMPLE
    ./set-owa-conditional-access.ps1 -TenantId 'contoso.onmicrosoft.com' -OwaRestriction ReadOnly -CreateConditionalAccessPolicy -Apply -WhatIf

    Shows the OWA mailbox policy change and the new Conditional Access policy without making them.

.NOTES
    Replaces the original 2018 method: Set-OwaMailboxPolicy -ConditionalAccessPolicy ReadOnly, run
    after connecting with the retired Basic authentication remote PowerShell guide (New-PSSession
    to outlook.office365.com/powershell-liveid), plus Conditional Access steps in the old Azure
    Active Directory blade of the Azure portal.
    Required GDAP roles: Exchange Administrator (OWA mailbox policy) and Conditional Access
    Administrator or Security Administrator (creating the policy). Global Reader or Security
    Reader is enough for report-only runs.
    Required partner app permissions: Exchange.Manage (Office 365 Exchange Online),
    Policy.Read.All (Microsoft Graph, reading Conditional Access policies),
    Policy.ReadWrite.ConditionalAccess (Microsoft Graph, only for -CreateConditionalAccessPolicy),
    User.Read (Microsoft Graph, initial domain lookup), and in the partner tenant Directory.Read.All
    or Directory.ReadWrite.All plus DelegatedAdminRelationship.ReadWrite.All (Microsoft Graph,
    customer list for Get-MspCustomer).

    Microsoft Learn: Create conditionalAccessPolicy (Graph v1.0)
    https://learn.microsoft.com/en-us/graph/api/conditionalaccessroot-post-policies?view=graph-rest-1.0
    Microsoft Learn: Session controls, application enforced restrictions
    https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-conditional-access-session

.LINK
    https://gcit.com.au/outlook-on-the-web-conditional-access/

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

    [ValidateSet('Off', 'ReadOnly', 'ReadOnlyPlusAttachmentsBlocked')]
    [string]$OwaRestriction = 'ReadOnly',

    [ValidateNotNullOrEmpty()]
    [string[]]$OwaMailboxPolicy = @('OwaMailboxPolicy-Default'),

    [switch]$CreateConditionalAccessPolicy,

    [ValidateNotNullOrEmpty()]
    [string]$ConditionalAccessPolicyName = 'Exchange Online: app enforced restrictions in the browser',

    [ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')]
    [string[]]$ExcludeGroupId,

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
        foreach ($column in 'Setting', 'Name', 'CurrentValue', 'DesiredValue', 'Status', 'Error') {
            $row[$column] = $Data[$column]
        }
        [pscustomobject]$row
    }

    # Office 365 Exchange Online, and the Office 365 app group that includes it.
    $exchangeAppId = '00000002-0000-0ff1-ce00-000000000000'
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
            # Conditional Access policies through Microsoft Graph.
            $matching = @()
            $caError = $null
            try {
                $policies = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Method GET -Uri 'v1.0/identity/conditionalAccess/policies' -ErrorAction Stop)
                $matching = @($policies | Where-Object {
                        $_.sessionControls -and $_.sessionControls.applicationEnforcedRestrictions -and
                        $_.sessionControls.applicationEnforcedRestrictions.isEnabled -and
                        (@($_.conditions.applications.includeApplications) | Where-Object { $_ -in $exchangeAppId, 'Office365', 'All' })
                    })
            }
            catch {
                $caError = $_.Exception.Message
            }
            if ($caError) {
                $row = ConvertTo-ResultRow -Customer $customer -Data @{ Setting = 'ConditionalAccessPolicy'; Status = 'Failed'; Error = $caError }
                $results.Add($row)
                $row
            }
            foreach ($policy in $matching) {
                $row = ConvertTo-ResultRow -Customer $customer -Data @{ Setting = 'ConditionalAccessPolicy'; Name = $policy.displayName; CurrentValue = $policy.state; DesiredValue = 'applicationEnforcedRestrictions'; Status = 'Found' }
                $results.Add($row)
                $row
            }
            if (-not $caError -and $matching.Count -eq 0) {
                $caData = @{ Setting = 'ConditionalAccessPolicy'; Name = $ConditionalAccessPolicyName; DesiredValue = 'enabledForReportingButNotEnforced' }
                if (-not $CreateConditionalAccessPolicy) {
                    $caData['Status'] = 'NoneFound'
                }
                elseif (-not $Apply) {
                    $caData['Status'] = 'WouldCreate'
                }
                elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $ConditionalAccessPolicyName", 'Create Conditional Access policy in report-only state')) {
                    $users = @{ includeUsers = @('All') }
                    if ($ExcludeGroupId) { $users['excludeGroups'] = @($ExcludeGroupId) }
                    $body = @{
                        displayName     = $ConditionalAccessPolicyName
                        state           = 'enabledForReportingButNotEnforced'
                        conditions      = @{
                            users          = $users
                            applications   = @{ includeApplications = @($exchangeAppId) }
                            clientAppTypes = @('browser')
                        }
                        sessionControls = @{ applicationEnforcedRestrictions = @{ isEnabled = $true } }
                    }
                    try {
                        $created = Invoke-MspGraphRequest -TenantId $customer.TenantId -Method POST -Uri 'v1.0/identity/conditionalAccess/policies' -Body $body -Confirm:$false -ErrorAction Stop
                        $caData['CurrentValue'] = $created.state
                        $caData['Status'] = 'Created'
                    }
                    catch {
                        $caData['Status'] = 'Failed'
                        $caData['Error'] = $_.Exception.Message
                    }
                }
                else {
                    $caData['Status'] = 'WhatIf'
                }
                $row = ConvertTo-ResultRow -Customer $customer -Data $caData
                $results.Add($row)
                $row
            }

            # OWA mailbox policies through Exchange Online.
            Write-Verbose "Connecting to Exchange Online for $($customer.Name) ($($customer.TenantId))."
            $connection = Connect-MspExchangeOnline -TenantId $customer.TenantId -ErrorAction Stop
            foreach ($owaPolicy in @(Get-OwaMailboxPolicy -ErrorAction Stop)) {
                $current = [string]$owaPolicy.ConditionalAccessPolicy
                $data = @{ Setting = 'OwaMailboxPolicy'; Name = $owaPolicy.Name; CurrentValue = $current; DesiredValue = $OwaRestriction }
                if ($OwaMailboxPolicy -notcontains $owaPolicy.Name) {
                    $data['DesiredValue'] = $null
                    $data['Status'] = 'Reported'
                }
                elseif ($current -eq $OwaRestriction) {
                    $data['Status'] = 'AlreadySet'
                }
                elseif (-not $Apply) {
                    $data['Status'] = 'WouldChange'
                }
                elseif ($PSCmdlet.ShouldProcess("$($customer.Name): $($owaPolicy.Name)", "Set ConditionalAccessPolicy to $OwaRestriction")) {
                    try {
                        Set-OwaMailboxPolicy -Identity $owaPolicy.Name -ConditionalAccessPolicy $OwaRestriction -Confirm:$false -ErrorAction Stop
                        $data['Status'] = 'Changed'
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
