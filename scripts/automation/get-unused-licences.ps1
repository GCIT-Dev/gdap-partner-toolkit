#Requires -Version 7.4
#Requires -Modules MspGdap

<#
.SYNOPSIS
    Reports paid licences that are bought but not assigned in customer tenants.

.DESCRIPTION
    Reads the subscribed SKUs in each customer tenant through Microsoft Graph (GET /subscribedSkus) and
    returns one row per SKU with the enabled, warning, suspended, consumed and unused unit counts. By
    default only SKUs with at least one unused enabled unit are returned, and well-known free or viral
    SKUs (which report very large unit counts) are left out.

    Run it interactively, on a schedule in Azure Automation, or from the UnusedLicenceAlert Azure Function
    in the functions folder, which turns each row into a queue message for Power Automate to email.

.PARAMETER TenantId
    One or more customer tenant IDs (GUID) or verified domains. Accepts pipeline input, including the
    output of Get-MspCustomer.

.PARAMETER AllCustomers
    Process every customer returned by Get-MspCustomer.

.PARAMETER MinimumUnused
    Only return SKUs with at least this many unused enabled units. Use 0 to return every SKU.

.PARAMETER ExcludeSkuPartNumber
    SKU part numbers to leave out, for example free and trial SKUs. Pass an empty array to include all.

.PARAMETER OutputPath
    Optional path for a CSV export of the results.

.EXAMPLE
    ./get-unused-licences.ps1 -AllCustomers -OutputPath ./unused-licences.csv

    Lists every customer SKU with at least one unused licence and saves a CSV.

.EXAMPLE
    'contoso.onmicrosoft.com' | ./get-unused-licences.ps1 -MinimumUnused 0

    Lists every SKU in one customer, including fully used ones.

.NOTES
    Replaces the original 2017 method: an Azure Functions v1 timer function with MSOnline
    (Connect-MsolService -Credential, Get-MsolPartnerContract, Get-MsolAccountSku) using an AES-encrypted
    stored password, writing to Azure Table storage with the storage account key.
    Required GDAP roles: Global Reader (or License Administrator).
    Required partner app permissions: Microsoft Graph delegated LicenseAssignment.Read.All or a higher
    privileged permission Learn lists for subscribedSkus (Directory.ReadWrite.All is in the full manifest).

.LINK
    https://gcit.com.au/knowledge-base/get-email-alerts-unused-office-365-licenses-azure-functions-azure-storage-microsoft-flow/

.LINK
    docs/07-migrating-from-dap-msonline.md

.LINK
    docs/08-unattended-automation.md
#>
[CmdletBinding(DefaultParameterSetName = 'Tenant')]
[OutputType([pscustomobject])]
param(
    [Parameter(Mandatory, ParameterSetName = 'Tenant', ValueFromPipeline, ValueFromPipelineByPropertyName)]
    [Alias('CustomerTenantId')]
    [ValidateNotNullOrEmpty()]
    [string[]]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'All')]
    [switch]$AllCustomers,

    [ValidateRange(0, [int]::MaxValue)]
    [int]$MinimumUnused = 1,

    [AllowEmptyCollection()]
    [string[]]$ExcludeSkuPartNumber = @(
        'FLOW_FREE', 'POWER_BI_STANDARD', 'POWERAPPS_VIRAL', 'POWERAPPS_DEV', 'TEAMS_EXPLORATORY',
        'WINDOWS_STORE', 'STREAM', 'CCIBOTS_PRIVPREV_VIRAL', 'RIGHTSMANAGEMENT_ADHOC', 'MICROSOFT_BUSINESS_CENTER'
    ),

    [string]$OutputPath
)

begin {
    $requested = [System.Collections.Generic.List[string]]::new()
    $results = [System.Collections.Generic.List[object]]::new()

    function ConvertTo-ResultRow {
        param($Customer, $Sku, $ErrorMessage)
        $enabled = if ($Sku) { [int]$Sku.prepaidUnits.enabled } else { $null }
        $consumed = if ($Sku) { [int]$Sku.consumedUnits } else { $null }
        [pscustomobject][ordered]@{
            CustomerTenantId = $Customer.TenantId
            CustomerName     = $Customer.Name
            SkuPartNumber    = $Sku.skuPartNumber
            SkuId            = $Sku.skuId
            CapabilityStatus = $Sku.capabilityStatus
            Enabled          = $enabled
            Warning          = if ($Sku) { [int]$Sku.prepaidUnits.warning } else { $null }
            Suspended        = if ($Sku) { [int]$Sku.prepaidUnits.suspended } else { $null }
            Consumed         = $consumed
            Unused           = if ($Sku) { [math]::Max(0, $enabled - $consumed) } else { $null }
            Message          = if ($Sku) { '{0} has {1} unused {2} licence(s)' -f $Customer.Name, [math]::Max(0, $enabled - $consumed), $Sku.skuPartNumber } else { $null }
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
                $organisation = Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'organization?$select=id,displayName' | Select-Object -First 1
                $customer.Name = $organisation.displayName
                if ($organisation.id) { $customer.TenantId = $organisation.id }
            }

            $skus = @(Invoke-MspGraphRequest -TenantId $customer.TenantId -Uri 'subscribedSkus')
            foreach ($sku in $skus) {
                if ($ExcludeSkuPartNumber -contains $sku.skuPartNumber) { continue }
                $row = ConvertTo-ResultRow -Customer $customer -Sku $sku
                if ($row.Unused -lt $MinimumUnused) { continue }
                $results.Add($row)
                $row
            }
        }
        catch {
            Write-Warning "Customer $($customer.TenantId): $($_.Exception.Message)"
            $row = ConvertTo-ResultRow -Customer $customer -ErrorMessage $_.Exception.Message
            $results.Add($row)
            $row
        }
    }

    if ($OutputPath) {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding utf8
    }
}
