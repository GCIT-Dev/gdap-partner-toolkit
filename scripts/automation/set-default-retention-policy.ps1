#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports Microsoft Purview retention policy coverage in customer tenants, and optionally creates an
    organisation-wide retention policy where a customer has none.

.DESCRIPTION
    Connects to Security and Compliance PowerShell in each customer with Connect-MspSecurityCompliance
    (confirmed in a live test, see docs/05) and lists the existing retention policies.

    By default the script only reports. With -Apply it creates a retention policy covering Exchange
    mailboxes, SharePoint sites, OneDrive accounts, Microsoft 365 Group mailboxes and sites and public
    folders, plus a rule that keeps content for -RetentionDuration days (or Unlimited). As in the
    original, tenants without SharePoint Online (checked through GET /subscribedSkus) get Exchange and
    public folder locations only, and a customer that already has any retention policy is skipped unless
    you add -CreateWhenOtherPoliciesExist.

    Retaining content is a records management, privacy and storage decision. Agree the retention period
    with each customer before you use -Apply. Content under a retain policy can't be permanently removed by
    users or admins until the period ends, and the Recoverable Items folder and the Preservation Hold
    library grow as a result. Teams chats and channel messages need their own policy and are not covered.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER PolicyName
    Name of the retention policy to create. The rule is named "<PolicyName> Rule".

.PARAMETER RetentionDuration
    Retention period in days, or Unlimited (the original default).

.PARAMETER Comment
    Comment stored on the policy, shown to the customer's admins in the Microsoft Purview portal.

.PARAMETER CreateWhenOtherPoliciesExist
    Create the policy when the customer has other retention policies but not this one.

.PARAMETER Apply
    Create the policy and rule where needed. Without it the script only reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./set-default-retention-policy.ps1 -AllCustomers -OutputPath ./retention-coverage.csv

    Reports which customers have no retention policy at all.

.EXAMPLE
    ./set-default-retention-policy.ps1 -TenantId 'contoso.onmicrosoft.com' -RetentionDuration 2555 -Apply -WhatIf

    Shows the seven-year retention policy that would be created in one customer, without creating it.

.NOTES
    Replaces the original 2017 method: MSOnline and DAP (Connect-MsolService, Get-MsolPartnerContract) with
    a basic authentication New-PSSession to https://ps.compliance.protection.outlook.com/powershell-liveid
    ?DelegatedOrg=, run interactively or from an Azure Functions v1 timer function with an AES-encrypted
    stored password.
    Required GDAP roles: Compliance Administrator (create policies), Global Reader or Compliance Administrator
    (report only), plus Global Reader or License Administrator to read licences (GET /subscribedSkus).
    Required partner app permissions: Office 365 Exchange Online delegated Exchange.Manage (used by
    Connect-MspSecurityCompliance), Microsoft Graph delegated User.Read (organisation name lookup) and
    LicenseAssignment.Read.All or Directory.ReadWrite.All (subscribedSkus, to check for SharePoint Online).
    Directory.ReadWrite.All is in the full manifest.
    Connect-MspSecurityCompliance was confirmed working with a delegated GDAP token in a live test on
    7 October 2026. Try one customer before you use -AllCustomers.

.LINK
    https://gcit.com.au/knowledge-base/apply-office-365-retention-policies-customers-via-powershell-delegated-administration/

.LINK
    docs/05-exchange-access.md

.LINK
    docs/07-migrating-from-dap-msonline.md
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateNotNullOrEmpty()]
    [string]$PolicyName = 'Organisation retention policy',

    [ValidatePattern('^(Unlimited|[1-9][0-9]{0,4})$')]
    [string]$RetentionDuration = 'Unlimited',

    [string]$Comment = 'Created by your Microsoft partner. Contact them before you change or remove this policy.',

    [switch]$CreateWhenOtherPoliciesExist,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-ResultRow {
        param($Customer, [object[]]$Policies, [string[]]$Locations, $Action, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId    = $Customer.TenantId
            CustomerName        = $Customer.Name
            ExistingPolicyCount = if ($null -ne $Policies) { @($Policies).Count } else { $null }
            ExistingPolicies    = (@($Policies) | ForEach-Object { $_.Name }) -join '; '
            PolicyName          = $PolicyName
            RetentionDuration   = $RetentionDuration
            Locations           = ($Locations -join ', ')
            Action              = $Action
            Error               = $ErrorMessage
        }
    }
}

process {
    foreach ($id in @($TenantId)) {
        if ($id) { $requested.Add($id) }
    }
}

end {
    $customers = if ($AllCustomers) {
        @(Get-MspCustomer | ForEach-Object { [pscustomobject]@{ TenantId = $_.TenantId; Name = $_.DisplayName } })
    }
    else {
        @($requested | Select-Object -Unique | ForEach-Object { [pscustomobject]@{ TenantId = $_; Name = $null } })
    }

    foreach ($customer in $customers) {
        try {
            if (-not $customer.Name) {
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            # As in the original, only add SharePoint, OneDrive and Microsoft 365 Group locations when the
            # tenant has SharePoint Online. An Exchange-only tenant gets Exchange and public folders only.
            # If the licences can't be read, assume SharePoint Online is present (true for most tenants).
            $hasSharePoint = $true
            try {
                $hasSharePoint = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'subscribedSkus' |
                        Where-Object { $_.capabilityStatus -ne 'Suspended' } |
                        ForEach-Object { $_.servicePlans } |
                        Where-Object { $_.servicePlanName -like 'SHAREPOINT*' -and $_.provisioningStatus -eq 'Success' }).Count -gt 0
            }
            catch {
                Write-Warning "Could not read licences for $($customer.Name), so SharePoint locations are included: $($_.Exception.Message)"
            }
            $locations = if ($hasSharePoint) { @('Exchange', 'SharePoint', 'OneDrive', 'ModernGroup', 'PublicFolder') } else { @('Exchange', 'PublicFolder') }

            $null = Connect-MspSecurityCompliance -TenantId $customer.TenantId
            try {
                $policies = @(Get-RetentionCompliancePolicy)
                $hasNamed = @($policies | Where-Object { $_.Name -eq $PolicyName }).Count -gt 0
                $needed = -not $hasNamed -and ($policies.Count -eq 0 -or $CreateWhenOtherPoliciesExist)

                if (-not $needed) {
                    $action = if ($hasNamed) { 'PolicyExists' } else { 'SkippedOtherPoliciesExist' }
                }
                elseif (-not $Apply) {
                    $action = 'WouldCreate'
                }
                elseif ($PSCmdlet.ShouldProcess($customer.Name, "Create retention policy '$PolicyName' ($RetentionDuration)")) {
                    $policyParams = @{
                        Name                 = $PolicyName
                        Comment              = $Comment
                        ExchangeLocation     = 'All'
                        PublicFolderLocation = 'All'
                        Enabled              = $true
                        ErrorAction          = 'Stop'
                    }
                    if ($hasSharePoint) {
                        $policyParams.SharePointLocation = 'All'
                        $policyParams.OneDriveLocation = 'All'
                        $policyParams.ModernGroupLocation = 'All'
                    }
                    $null = New-RetentionCompliancePolicy @policyParams
                    $ruleParams = @{
                        Name                       = "$PolicyName Rule"
                        Policy                     = $PolicyName
                        RetentionDuration          = $RetentionDuration
                        RetentionComplianceAction  = 'Keep'
                        ErrorAction                = 'Stop'
                    }
                    $null = New-RetentionComplianceRule @ruleParams
                    $created = Get-RetentionCompliancePolicy -Identity $PolicyName -ErrorAction SilentlyContinue
                    $action = if ($created) { 'Created' } else { 'CreatedNotConfirmed' }
                    $policies = @(Get-RetentionCompliancePolicy)
                }
                else {
                    $action = 'WhatIf'
                }
                $row = ConvertTo-ResultRow -Customer $customer -Policies $policies -Locations $locations -Action $action
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
            $results.Add($row)
            $row
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -Action 'Failed' -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
