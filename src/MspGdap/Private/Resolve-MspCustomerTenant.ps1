function Resolve-MspCustomerTenant {
    <#
    .SYNOPSIS
        Resolves a customer tenant ID or domain and refuses the configured partner tenant.
    .DESCRIPTION
        Used by every public command that acts on a customer only (consent, Exchange, GDAP, Connect-*).
        A tenant that resolves to the configured partner tenant is refused with
        MspGdap.Tenant.PartnerTenantNotAllowed, so a typo or copy-paste can never point a customer
        command at the partner tenant. When no partner tenant is configured yet the check is skipped.
    .PARAMETER TenantId
        Customer tenant GUID or verified domain.
    .EXAMPLE
        $tenant = Resolve-MspCustomerTenant -TenantId 'fabrikam.onmicrosoft.com'
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$TenantId
    )
    $resolved = ([string](Resolve-MspTenantId -TenantId $TenantId)).ToLowerInvariant()
    $partner = $null
    try {
        $config = Get-MspConfiguration -ErrorAction Stop
        if ($config -and $config.PSObject.Properties['PartnerTenantId'] -and $config.PartnerTenantId) {
            $partner = ([string]$config.PartnerTenantId).ToLowerInvariant()
        }
    }
    catch {
        $partner = $null
    }
    if ($partner -and $resolved -eq $partner) {
        throw (New-MspErrorRecord -Message "TenantId $TenantId is your partner tenant. This command only acts on customer tenants." -ErrorId 'MspGdap.Tenant.PartnerTenantNotAllowed' -Category InvalidArgument -TargetObject $TenantId)
    }
    $resolved
}
