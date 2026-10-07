function Connect-MspTeams {
    <#
    .SYNOPSIS
        Connects Microsoft Teams PowerShell to one customer with the technician's MspGdap tokens.
    .DESCRIPTION
        Optional. Needs the MicrosoftTeams module. Gets two delegated access tokens in the customer, one for
        Microsoft Graph and one for the Skype and Teams Tenant Admin API (48ac35b8-9aa8-4d74-927d-1f4a14a0b239),
        and passes both to Connect-MicrosoftTeams -AccessTokens. The partner app must be consented for the
        Teams admin API (user_impersonation) in the customer, which partner-app.full.json includes.
        The returned tenant is checked against -TenantId. Tokens are not refreshed by the Teams module.
        Your partner tenant is refused as -TenantId. The session is recorded so Disconnect-Msp closes it.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .EXAMPLE
        Connect-MspTeams -TenantId 'fabrikam.onmicrosoft.com'
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Microsoft Teams is a product name.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId)

    $moduleVersion = Assert-MspModuleAvailable -Name 'MicrosoftTeams' -MinimumVersion '5.0.0'
    $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
    $graph = Get-MspAccessToken -TenantId $tenant -Resource 'https://graph.microsoft.com'
    $teams = Get-MspAccessToken -TenantId $tenant -Resource '48ac35b8-9aa8-4d74-927d-1f4a14a0b239'
    $tokens = @((ConvertTo-MspPlainToken -InputObject $graph), (ConvertTo-MspPlainToken -InputObject $teams))
    try { $result = Connect-MicrosoftTeams -AccessTokens $tokens -ErrorAction Stop }
    finally { $tokens = $null }
    if ($script:MspState -and $script:MspState.Connections) { $null = $script:MspState.Connections.Add('Teams') }

    $connectedTenant = $null
    if ($result) {
        if ($result.PSObject.Properties['TenantId'] -and $result.TenantId) { $connectedTenant = ([string]$result.TenantId).ToLowerInvariant() }
        elseif ($result.PSObject.Properties['Tenant'] -and $result.Tenant) {
            $tenantValue = if ($result.Tenant.PSObject.Properties['Id']) { $result.Tenant.Id } else { $result.Tenant }
            $connectedTenant = ([string]$tenantValue).ToLowerInvariant()
        }
    }
    if ($connectedTenant -and $connectedTenant -ne $tenant) {
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null
        throw ("Microsoft Teams connected to tenant {0}, not {1}. The session was closed." -f $connectedTenant, $tenant)
    }
    if (-not $connectedTenant) { Write-Warning 'Connect-MicrosoftTeams did not report a tenant ID, so the tenant could not be double-checked.' }
    [pscustomobject]@{
        PSTypeName    = 'MspGdap.TeamsConnection'
        TenantId      = $tenant
        Account       = if ($result -and $result.PSObject.Properties['Account']) { [string]$result.Account } else { $null }
        Verified      = [bool]$connectedTenant
        ModuleVersion = $moduleVersion
    }
}
