function Connect-MspSecurityCompliance {
    <#
    .SYNOPSIS
        Connects Security and Compliance PowerShell to one customer with a delegated GDAP token.
    .DESCRIPTION
        Connect-IPPSSession -AccessToken exists from ExchangeOnlineManagement 3.8.0, but Microsoft Learn does not
        document which token audience it expects for a GDAP partner. Security and Compliance PowerShell is a
        different resource from Exchange Online, so the default is an access token for
        https://ps.compliance.protection.outlook.com in the customer, the audience the module itself requests.
        -Organization is set to the customer's initial .onmicrosoft.com domain, although Learn only documents
        -Organization for certificate connections. This default combination was confirmed working in a live
        test on 7 October 2026 (Get-RetentionCompliancePolicy in a GDAP customer, no extra consent needed).
        Because Microsoft documents it only briefly, change -TokenResource if Microsoft documents another
        audience, and try one customer first after a module upgrade.
        -UseDelegatedOrganization switches to -DelegatedOrganization, which Microsoft says must be used with
        -AzureADAuthorizationEndpointUri. Only https://login.microsoftonline.com and
        https://login.microsoftonline.us authorisation endpoints are accepted.
        The session's tenant is checked after connecting where Get-ConnectionInformation reports it, your
        partner tenant is refused as -TenantId, and the session is recorded so Disconnect-Msp closes it.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .PARAMETER Organization
        Initial .onmicrosoft.com domain. Read from Graph when omitted.
    .PARAMETER TokenResource
        Resource the delegated access token is requested for. Default https://ps.compliance.protection.outlook.com.
    .PARAMETER UseDelegatedOrganization
        Pass -DelegatedOrganization instead of -Organization.
    .PARAMETER AzureADAuthorizationEndpointUri
        Authorisation endpoint for -UseDelegatedOrganization, for example
        https://login.microsoftonline.com/<customer tenant ID>/oauth2/authorize.
    .PARAMETER Prefix
        Noun prefix for the imported cmdlets.
    .PARAMETER KeepExistingConnections
        Do not close other Exchange Online and Security and Compliance sessions first.
    .EXAMPLE
        Connect-MspSecurityCompliance -TenantId 'fabrikam.onmicrosoft.com'
    #>
    [CmdletBinding(DefaultParameterSetName = 'Organization')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [ValidatePattern('^[A-Za-z0-9-]+\.onmicrosoft\.(com|us|de)$')][string]$Organization,
        [string]$TokenResource = 'https://ps.compliance.protection.outlook.com',
        [Parameter(Mandatory, ParameterSetName = 'Delegated')][switch]$UseDelegatedOrganization,
        [Parameter(Mandatory, ParameterSetName = 'Delegated')][ValidatePattern('^https://login\.microsoftonline\.(com|us)/')][string]$AzureADAuthorizationEndpointUri,
        [string]$Prefix,
        [switch]$KeepExistingConnections
    )
    Write-Verbose 'Connect-IPPSSession -AccessToken with a delegated GDAP token is sparsely documented by Microsoft. It was confirmed working in a live test on 7 October 2026.'
    $moduleVersion = Assert-MspModuleAvailable -Name 'ExchangeOnlineManagement' -MinimumVersion '3.8.0'
    $problem = Test-MspExoCompatibility -ModuleVersion $moduleVersion -MinimumModuleVersion '3.8.0'
    if ($problem) { throw $problem }

    $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
    if (-not $Organization) { $Organization = (Get-MspInitialDomain -TenantId $tenant).InitialDomain }

    if (-not $KeepExistingConnections) {
        $open = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
        if ($open.Count -gt 0) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue }
    }
    $before = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object { $_.ConnectionId })

    $connect = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($Prefix) { $connect['Prefix'] = $Prefix }
    if ($UseDelegatedOrganization) {
        $connect['DelegatedOrganization'] = $Organization
        $connect['AzureADAuthorizationEndpointUri'] = $AzureADAuthorizationEndpointUri
    }
    else {
        $connect['Organization'] = $Organization
    }
    $tokenResult = Get-MspAccessToken -TenantId $tenant -Resource $TokenResource
    $connect['AccessToken'] = ConvertTo-MspPlainToken -InputObject $tokenResult
    try { Connect-IPPSSession @connect | Out-Null }
    finally { $connect['AccessToken'] = $null; $connect.Remove('AccessToken') }

    $sessions = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' -and $before -notcontains $_.ConnectionId })
    $session = @($sessions | Where-Object { $_.IsEopSession -or ([string]$_.ConnectionUri -match 'compliance') }) | Select-Object -First 1
    if (-not $session) { $session = $sessions | Select-Object -First 1 }
    if (-not $session) { throw "Connect-IPPSSession returned, but no new connected session for $tenant was found." }
    if ($session.TenantID -and ([string]$session.TenantID).ToLowerInvariant() -ne $tenant) {
        Disconnect-ExchangeOnline -ConnectionId $session.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue
        throw ("Security and Compliance connected to tenant {0}, not {1}. The session was closed." -f $session.TenantID, $tenant)
    }
    if ($script:MspState -and $script:MspState.Connections) {
        $null = $script:MspState.Connections.Add('SecurityCompliance')
        if ($session.ConnectionId -and $script:MspState.ExchangeConnections) { $null = $script:MspState.ExchangeConnections.Add([string]$session.ConnectionId) }
    }

    [pscustomobject]@{
        PSTypeName         = 'MspGdap.ExchangeConnection'
        TenantId           = $tenant
        Organization       = $Organization
        Mode               = 'SecurityComplianceDelegated'
        ConnectionId       = $session.ConnectionId
        State              = $session.State
        UserPrincipalName  = $session.UserPrincipalName
        AppId              = $session.AppId
        TokenExpiryTimeUTC = $session.TokenExpiryTimeUTC
        ModuleVersion      = $moduleVersion
        Experimental       = $false
    }
}
