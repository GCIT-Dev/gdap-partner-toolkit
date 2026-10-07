function Resolve-MspTargetTenant {
    <#
    .SYNOPSIS
        Resolves the tenant a tenant-scoped call acts on, with partner-tenant safety.
    .DESCRIPTION
        Every tenant-scoped MspGdap function takes a mandatory -TenantId or an
        explicit -PartnerTenant switch. This helper enforces the rule that a
        call never lands in the partner tenant by accident: if -TenantId
        resolves to the configured partner tenant and -PartnerTenant was not
        given, it throws. An empty tenant never falls back to anything.
    .PARAMETER TenantId
        Customer tenant GUID or verified domain.
    .PARAMETER PartnerTenant
        Target the configured partner tenant deliberately.
    .PARAMETER Configuration
        The active configuration object.
    .EXAMPLE
        $tid = Resolve-MspTargetTenant -TenantId 'contoso.onmicrosoft.com' -Configuration $config
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [AllowEmptyString()]
        [AllowNull()]
        [string]$TenantId,

        [switch]$PartnerTenant,

        [Parameter(Mandatory)]
        [pscustomobject]$Configuration
    )

    if (-not $Configuration.PartnerTenantId) {
        throw (New-MspErrorRecord -Message 'PartnerTenantId is not configured. Run Set-MspConfiguration -PartnerTenantId.' -ErrorId 'MspGdap.Configuration.Incomplete' -Category NotSpecified)
    }
    $partner = ([string]$Configuration.PartnerTenantId).ToLowerInvariant()

    if ($PartnerTenant) {
        if ($TenantId) {
            $resolved = Resolve-MspTenantId -Tenant $TenantId
            if ($resolved -ne $partner) {
                throw (New-MspErrorRecord -Message "-PartnerTenant was given with -TenantId $TenantId, which is not the configured partner tenant. Use one or the other." -ErrorId 'MspGdap.Tenant.Conflict' -Category InvalidArgument -TargetObject $TenantId)
            }
        }
        return $partner
    }

    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        throw (New-MspErrorRecord -Message 'A customer -TenantId is required. MspGdap never defaults to a tenant. Use -PartnerTenant to act on your own partner tenant.' -ErrorId 'MspGdap.Tenant.Missing' -Category InvalidArgument)
    }

    $resolvedTenant = Resolve-MspTenantId -Tenant $TenantId
    if ($resolvedTenant -eq $partner) {
        throw (New-MspErrorRecord -Message "TenantId $TenantId is your partner tenant. Pass -PartnerTenant instead of -TenantId to act on it deliberately." -ErrorId 'MspGdap.Tenant.PartnerTenantNotAllowed' -Category InvalidArgument -TargetObject $TenantId)
    }
    $resolvedTenant
}
