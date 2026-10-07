function Connect-MspGraph {
    <#
    .SYNOPSIS
        Connects the Microsoft Graph PowerShell SDK to one customer with the technician's MspGdap token.
    .DESCRIPTION
        Optional. Needs Microsoft.Graph.Authentication. Passes the cached Graph access token for the customer to
        Connect-MgGraph -AccessToken (a SecureString), then checks Get-MgContext reports the same tenant.
        The SDK does not refresh a token it was given. Run Connect-MspGraph again when ExpiresOn passes.
        For plain REST calls you do not need this: use Invoke-MspGraphRequest.
        Your partner tenant is refused as -TenantId. The session is recorded so Disconnect-Msp closes it.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .EXAMPLE
        Connect-MspGraph -TenantId 'fabrikam.onmicrosoft.com'
        Get-MgUser -Top 5
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId)

    $moduleVersion = Assert-MspModuleAvailable -Name 'Microsoft.Graph.Authentication' -MinimumVersion '2.0.0'
    $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
    $tokenResult = Get-MspAccessToken -TenantId $tenant -Resource 'https://graph.microsoft.com'
    $secure = ConvertTo-MspSecureToken -InputObject $tokenResult

    $connect = @{ AccessToken = $secure; ErrorAction = 'Stop' }
    if ((Get-Command -Name 'Connect-MgGraph').Parameters.ContainsKey('NoWelcome')) { $connect['NoWelcome'] = $true }
    Connect-MgGraph @connect | Out-Null
    if ($script:MspState -and $script:MspState.Connections) { $null = $script:MspState.Connections.Add('MgGraph') }

    $context = Get-MgContext
    if (-not $context -or ([string]$context.TenantId).ToLowerInvariant() -ne $tenant) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        throw ("Microsoft Graph SDK connected to tenant {0}, not {1}. The session was closed." -f $(if ($context) { $context.TenantId } else { 'none' }), $tenant)
    }
    $expires = $null
    if ($tokenResult -and $tokenResult -isnot [string] -and $tokenResult -isnot [System.Security.SecureString]) {
        foreach ($name in @('ExpiresOn', 'ExpiresOnUtc', 'ExpiresAt')) {
            if ($tokenResult.PSObject.Properties[$name]) { $expires = $tokenResult.$name; break }
        }
    }
    Write-Warning 'The Microsoft Graph SDK will not refresh this token. Run Connect-MspGraph again after it expires.'
    [pscustomobject]@{
        PSTypeName    = 'MspGdap.GraphConnection'
        TenantId      = $tenant
        Account       = $context.Account
        Scopes        = @($context.Scopes)
        AuthType      = [string]$context.AuthType
        ExpiresOn     = $expires
        ModuleVersion = $moduleVersion
    }
}
