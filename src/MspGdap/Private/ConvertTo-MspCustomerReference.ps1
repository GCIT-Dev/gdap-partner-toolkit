function ConvertTo-MspCustomerReference {
    <#
    .SYNOPSIS
        Normalises a customer object (Get-MspCustomer output, a Partner Center customer, a GDAP relationship
        customer or a plain tenant ID string) into TenantId, DisplayName and PartnerCenterId.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory, ValueFromPipeline)][AllowNull()][object]$InputObject)
    process {
        if ($null -eq $InputObject) { return }
        if ($InputObject -is [string]) {
            return [pscustomobject]@{ TenantId = $InputObject; DisplayName = $null; PartnerCenterId = $null }
        }
        $read = {
            param($Object, [string[]]$Names)
            foreach ($name in $Names) {
                $current = $Object
                foreach ($part in $name.Split('.')) {
                    if ($null -eq $current) { break }
                    $prop = $current.PSObject.Properties[$part]
                    $current = if ($prop) { $prop.Value } else { $null }
                }
                if ($null -ne $current -and [string]$current -ne '') { return [string]$current }
            }
            return $null
        }
        $tenantId = & $read $InputObject @('TenantId', 'tenantId', 'CustomerTenantId', 'companyProfile.tenantId', 'customer.tenantId')
        $name = & $read $InputObject @('DisplayName', 'displayName', 'Name', 'CompanyName', 'companyProfile.companyName', 'customer.displayName')
        $pcId = & $read $InputObject @('PartnerCenterId', 'CustomerId', 'id')
        if (-not $tenantId -and $pcId) { $tenantId = $pcId }
        [pscustomobject]@{ TenantId = $tenantId; DisplayName = $name; PartnerCenterId = $pcId }
    }
}
