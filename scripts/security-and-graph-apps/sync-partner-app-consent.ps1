#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Checks, and optionally grants, consent for the MspGdap partner app in customer tenants, so one app can
    call Microsoft Graph and other APIs in every customer.

.DESCRIPTION
    Under DAP, adding an app's service principal to the partner's AdminAgents group gave it access to every
    customer. That no longer works under GDAP. Each multi-tenant app now has to be consented in each
    customer, and the technician's GDAP roles set what it can do there.

    Create the partner app once with New-MspPartnerApp (docs/02), then use this script to keep its consent
    in step with your manifest across customers:

    - Test-MspPartnerAppConsent reports whether each customer has every delegated permission in the
      manifest, and
    - with -Apply, Grant-MspPartnerAppConsent adds the missing consent through the Partner Center
      applicationconsents API and reads it back through Microsoft Graph.

    By default the script only reports. -WhatIf shows what -Apply would do.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER ManifestPath
    Permission manifest to check and grant, for example ./manifests/partner-app.minimal.json. Defaults to
    the manifest MspGdap uses when none is given.

.PARAMETER Apply
    Grant consent where it is missing or incomplete. Without it the script only reports.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./sync-partner-app-consent.ps1 -AllCustomers -ManifestPath ./manifests/partner-app.minimal.json -OutputPath ./consent.csv

    Reports which customers are missing consent for the partner app.

.EXAMPLE
    ./sync-partner-app-consent.ps1 -TenantId 'contoso.onmicrosoft.com' -ManifestPath ./manifests/partner-app.minimal.json -Apply -WhatIf

    Shows the consent that would be granted in one customer.

.NOTES
    Replaces the original 2018 method: the AzureAD and AzureRM modules (New-AzureADApplication
    -AvailableToOtherTenants, New-AzureRmADServicePrincipal) to create a multi-tenant app with a 99-year
    client secret exported to CSV, then Add-AzureADGroupMember to put its service principal in AdminAgents.
    Required GDAP roles: Cloud Application Administrator in each customer, plus AdminAgents membership in
    the partner tenant for the Partner Center API.
    Required partner app permissions: Microsoft Partner Center delegated user_impersonation, Microsoft Graph
    delegated Application.ReadWrite.All and DelegatedPermissionGrant.ReadWrite.All or Directory.ReadWrite.All
    for the readback (see docs/04). Directory.ReadWrite.All and Application.ReadWrite.All are in the full
    manifest.

.LINK
    https://gcit.com.au/knowledge-base/create-an-azure-ad-application-with-access-to-customer-tenants/

.LINK
    docs/02-create-partner-app.md

.LINK
    docs/04-preconsent-customers.md
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

    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$ManifestPath,

    [switch]$Apply,

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()
    $manifestParams = @{}
    if ($ManifestPath) { $manifestParams.ManifestPath = (Resolve-Path -LiteralPath $ManifestPath).ProviderPath }

    function ConvertTo-ResultRow {
        param($Customer, $Consented, $Outcome, $Action, [object[]]$Steps, $ErrorMessage)
        [pscustomobject][ordered]@{
            CustomerTenantId = $Customer.TenantId
            CustomerName     = $Customer.Name
            Consented        = $Consented
            Outcome          = $Outcome
            Action           = $Action
            Detail           = (@($Steps | Where-Object { $_.Status -notin @('Passed') } | ForEach-Object { '{0} {1}: {2}' -f $_.Status, $_.Step, $_.Detail })) -join '; '
            Error            = $ErrorMessage
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
                # Look the customer up in the partner tenant. A customer without consent can't answer a Graph
                # call in its own tenant yet, which is exactly the case this script is for.
                $known = $null
                try { $known = @(Get-MspCustomer -TenantId $customer.TenantId -ErrorAction Stop) | Select-Object -First 1 }
                catch { Write-Verbose "Customer lookup for $($customer.TenantId) failed: $($_.Exception.Message)" }
                if ($known) {
                    $customer.Name = $known.DisplayName
                    $customer.TenantId = $known.TenantId
                }
            }

            $test = Test-MspPartnerAppConsent -TenantId $customer.TenantId @manifestParams
            $consented = [bool]$test.Success
            $outcome = $test.Outcome
            $steps = @($test.Steps)
            if ($consented) {
                $action = 'AlreadyConsented'
            }
            elseif (-not $Apply) {
                $action = 'WouldGrant'
            }
            elseif ($PSCmdlet.ShouldProcess($customer.Name, 'Grant partner app consent')) {
                $grant = Grant-MspPartnerAppConsent -TenantId $customer.TenantId @manifestParams -Confirm:$false -ErrorAction SilentlyContinue |
                    Select-Object -Last 1
                $consented = [bool]$grant.Success
                $outcome = $grant.Outcome
                $steps = @($grant.Steps)
                $action = if ($consented) { 'Granted' } else { 'GrantFailed' }
            }
            else {
                $action = 'WhatIf'
            }
            $row = ConvertTo-ResultRow -Customer $customer -Consented $consented -Outcome $outcome -Action $action -Steps $steps
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
