function Get-MspAuthHeader {
    <#
    .SYNOPSIS
        Returns a headers hashtable with a bearer token for Invoke-RestMethod.
    .DESCRIPTION
        Wraps Get-MspAccessToken, so the token comes from the cache when it is
        still valid. The hashtable holds the token in plain text because HTTP
        needs it. Do not log or export it.
    .PARAMETER TenantId
        Customer tenant GUID or verified domain. Mandatory unless -PartnerTenant is used.
    .PARAMETER PartnerTenant
        Use your own partner tenant deliberately.
    .PARAMETER Resource
        Resource alias, URI or application ID. Defaults to Microsoft Graph.
    .PARAMETER Scope
        Specific scopes for one resource instead of .default.
    .PARAMETER UserPrincipalName
        Technician whose refresh token is used.
    .PARAMETER ForceRefresh
        Skip the cache.
    .PARAMETER AdditionalHeaders
        Extra headers to merge in (for example ConsistencyLevel).
    .EXAMPLE
        $headers = Get-MspAuthHeader -TenantId 'contoso.onmicrosoft.com'
        Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/organization' -Headers $headers
    .EXAMPLE
        $headers = Get-MspAuthHeader -TenantId $tid -Resource ManagementApi
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant', Position = 0, ValueFromPipelineByPropertyName)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = 'Partner')]
        [switch]$PartnerTenant,

        [string]$Resource,

        [string[]]$Scope,

        [string]$UserPrincipalName,

        [switch]$ForceRefresh,

        [hashtable]$AdditionalHeaders
    )
    process {
        $params = @{
            Resource     = $Resource
            Scope        = $Scope
            ForceRefresh = $ForceRefresh
            AsPlainText  = $true
        }
        if ($UserPrincipalName) { $params.UserPrincipalName = $UserPrincipalName }
        if ($PartnerTenant) { $params.PartnerTenant = $true } else { $params.TenantId = $TenantId }

        try {
            $token = Get-MspAccessToken @params
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
        $headers = @{ Authorization = "Bearer $token" }
        if ($AdditionalHeaders) {
            foreach ($key in $AdditionalHeaders.Keys) {
                if ($key -ne 'Authorization') { $headers[$key] = $AdditionalHeaders[$key] }
            }
        }
        $headers
    }
}
