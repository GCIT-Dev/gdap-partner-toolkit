function Resolve-MspTenantId {
    <#
    .SYNOPSIS
        Resolves a verified domain to its Entra tenant ID (GUID).
    .DESCRIPTION
        A GUID is returned unchanged (lower case). A domain is looked up through
        the public OpenID Connect discovery document
        https://login.microsoftonline.com/<domain>/v2.0/.well-known/openid-configuration,
        which needs no token. Results are cached for the session.
    .PARAMETER Tenant
        Tenant GUID or a verified domain such as contoso.onmicrosoft.com.
    .EXAMPLE
        Resolve-MspTenantId -Tenant 'contoso.onmicrosoft.com'
    .EXAMPLE
        'contoso.com', 'fabrikam.com' | Resolve-MspTenantId
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [Alias('TenantId', 'DomainName')]
        [ValidateNotNullOrEmpty()]
        [string]$Tenant
    )
    process {
        $value = $Tenant.Trim().ToLowerInvariant()
        if ($value -match $script:MspConstants.GuidPattern) {
            return $value
        }
        if ($value -in 'common', 'organizations', 'consumers') {
            $PSCmdlet.ThrowTerminatingError((New-MspErrorRecord -Message "'$Tenant' is not a specific tenant. Use a tenant ID or a verified domain." -ErrorId 'MspGdap.Tenant.InvalidIdentifier' -Category InvalidArgument -TargetObject $Tenant))
        }
        if ($value -notmatch '^(?=.{3,253}$)[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?(\.[a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?)+$') {
            $PSCmdlet.ThrowTerminatingError((New-MspErrorRecord -Message "'$Tenant' is not a tenant GUID or a valid domain name." -ErrorId 'MspGdap.Tenant.InvalidIdentifier' -Category InvalidArgument -TargetObject $Tenant))
        }
        if ($script:MspState.TenantIdCache.ContainsKey($value)) {
            return $script:MspState.TenantIdCache[$value]
        }

        $uri = '{0}/{1}/v2.0/.well-known/openid-configuration' -f $script:MspConstants.Authority, $value
        try {
            $document = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction Stop
        }
        catch {
            $status = Get-MspHttpStatusCode -ErrorRecord $_
            if ($status -ge 400 -and $status -lt 500) {
                $PSCmdlet.ThrowTerminatingError((New-MspErrorRecord -Message "No Entra tenant was found for domain '$Tenant'." -ErrorId 'MspGdap.Tenant.NotFound' -Category ObjectNotFound -TargetObject $Tenant))
            }
            $statusText = if ($status) { "HTTP $status" } else { 'no response' }
            $PSCmdlet.ThrowTerminatingError((New-MspErrorRecord -Message "Could not look up domain '$Tenant' at login.microsoftonline.com ($statusText). Check the network connection and try again. $($_.Exception.Message)" -ErrorId 'MspGdap.Tenant.LookupFailed' -Category ConnectionError -TargetObject $Tenant))
        }
        $issuer = if ($document -and $document.PSObject.Properties['issuer']) { [string]$document.issuer } else { '' }
        if ($issuer -notmatch '^https://login\.microsoftonline\.com/([0-9a-fA-F-]{36})/v2\.0/?$') {
            $PSCmdlet.ThrowTerminatingError((New-MspErrorRecord -Message "The discovery document for '$Tenant' did not contain a tenant ID." -ErrorId 'MspGdap.Tenant.NotFound' -Category ObjectNotFound -TargetObject $Tenant))
        }
        $tenantId = $Matches[1].ToLowerInvariant()
        $script:MspState.TenantIdCache[$value] = $tenantId
        Write-Verbose "Resolved $value to tenant $tenantId."
        $tenantId
    }
}
