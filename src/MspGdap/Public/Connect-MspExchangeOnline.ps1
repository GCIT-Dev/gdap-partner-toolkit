function Connect-MspExchangeOnline {
    <#
    .SYNOPSIS
        Connects Exchange Online PowerShell to one customer, delegated through GDAP (default) or app-only.
    .DESCRIPTION
        Delegated (default): gets an access token for https://outlook.office365.com in the customer tenant
        from the technician's MspGdap token, then runs
          Connect-ExchangeOnline -AccessToken <token> -DelegatedOrganization <initial onmicrosoft.com domain> -ShowBanner:$false
        Microsoft Learn documents -AccessToken (module 3.1.0 or later) with -DelegatedOrganization, which accepts
        the primary .onmicrosoft.com domain or the tenant ID. The token-based GDAP combination is only shown in a
        retired Microsoft sample, so treat it as supported but sparsely documented.

        App-only (-AppOnly): for the separate automation app prepared with Enable-MspExchangeAppAccess:
          Connect-ExchangeOnline -AppId <id> -CertificateThumbprint <thumb> -Organization <initial domain>
        CNG certificates are not supported by Exchange app-only authentication.

        Existing Exchange Online sessions are closed first (unless -KeepExistingConnections) so commands cannot
        run against the previous customer. After connecting, the session's tenant is checked against -TenantId,
        and a mismatched session is disconnected with an error. Your partner tenant is refused as -TenantId.
        The session is recorded so Disconnect-Msp closes it.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .PARAMETER Organization
        Initial .onmicrosoft.com domain, if you already know it. Otherwise read from Graph.
    .PARAMETER AppOnly
        Connect as the separate automation app (certificate) instead of the technician.
    .PARAMETER AppId
        Application ID of the automation app. App-only only.
    .PARAMETER CertificateThumbprint
        Thumbprint of the automation app certificate in the local store (Windows). App-only only.
    .PARAMETER Certificate
        The automation app certificate as an X509Certificate2 object. App-only only.
    .PARAMETER Prefix
        Noun prefix for the imported Exchange cmdlets, passed to Connect-ExchangeOnline -Prefix.
    .PARAMETER CommandName
        Import only these Exchange cmdlets, passed to Connect-ExchangeOnline -CommandName.
    .PARAMETER KeepExistingConnections
        Do not close other Exchange Online sessions first.
    .EXAMPLE
        Connect-MspExchangeOnline -TenantId 'fabrikam.onmicrosoft.com'
        Get-Mailbox -ResultSize 10
    .OUTPUTS
        MspGdap.ExchangeConnection
    #>
    [CmdletBinding(DefaultParameterSetName = 'Delegated')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [ValidatePattern('^[A-Za-z0-9-]+\.onmicrosoft\.(com|us|de)$|^[A-Za-z0-9-]+\.partner\.onmschina\.cn$')][string]$Organization,
        [Parameter(Mandatory, ParameterSetName = 'AppOnlyThumbprint')][Parameter(Mandatory, ParameterSetName = 'AppOnlyCertificate')][switch]$AppOnly,
        [Parameter(Mandatory, ParameterSetName = 'AppOnlyThumbprint')][Parameter(Mandatory, ParameterSetName = 'AppOnlyCertificate')][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [Parameter(Mandatory, ParameterSetName = 'AppOnlyThumbprint')][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$CertificateThumbprint,
        [Parameter(Mandatory, ParameterSetName = 'AppOnlyCertificate')][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$Prefix,
        [string[]]$CommandName,
        [switch]$KeepExistingConnections
    )
    $moduleVersion = Assert-MspModuleAvailable -Name 'ExchangeOnlineManagement' -MinimumVersion '3.1.0'
    $problem = Test-MspExoCompatibility -ModuleVersion $moduleVersion
    if ($problem) { throw $problem }

    $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
    $mode = if ($AppOnly) { 'AppOnly' } else { 'Delegated' }

    # Work in $org: $Organization carries [ValidatePattern], so assigning the tenant ID fallback to it would throw.
    $org = $Organization
    if (-not $org) {
        try { $org = (Get-MspInitialDomain -TenantId $tenant).InitialDomain }
        catch {
            if ($AppOnly) { throw "Could not read the initial .onmicrosoft.com domain for $tenant, which app-only connections need. Pass -Organization. $($_.Exception.Message)" }
            Write-Warning "Could not read the initial domain for $tenant ($($_.Exception.Message)). Using the tenant ID for -DelegatedOrganization, which Microsoft also accepts."
            $org = $tenant
        }
    }

    if (-not $KeepExistingConnections) {
        $open = @(Get-ConnectionInformation -ErrorAction SilentlyContinue)
        if ($open.Count -gt 0) {
            Write-Verbose "Closing $($open.Count) existing Exchange Online session(s) so commands cannot reach another customer."
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        }
    }
    $before = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | ForEach-Object { $_.ConnectionId })

    $connect = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ($Prefix) { $connect['Prefix'] = $Prefix }
    if ($CommandName) { $connect['CommandName'] = $CommandName }
    if ($AppOnly) {
        $connect['AppId'] = $AppId
        $connect['Organization'] = $org
        if ($PSCmdlet.ParameterSetName -eq 'AppOnlyThumbprint') { $connect['CertificateThumbprint'] = $CertificateThumbprint } else { $connect['Certificate'] = $Certificate }
        Connect-ExchangeOnline @connect | Out-Null
    }
    else {
        $tokenResult = Get-MspAccessToken -TenantId $tenant -Resource 'https://outlook.office365.com'
        $connect['DelegatedOrganization'] = $org
        $connect['AccessToken'] = ConvertTo-MspPlainToken -InputObject $tokenResult
        try { Connect-ExchangeOnline @connect | Out-Null }
        finally { $connect['AccessToken'] = $null; $connect.Remove('AccessToken') }
    }

    $sessions = @(Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' -and $before -notcontains $_.ConnectionId })
    $session = @($sessions | Where-Object { $_.TenantID -and ([string]$_.TenantID).ToLowerInvariant() -eq $tenant }) | Select-Object -First 1
    if (-not $session) {
        $wrong = @($sessions | Where-Object { $_.TenantID -and ([string]$_.TenantID).ToLowerInvariant() -ne $tenant })
        if ($wrong.Count -gt 0) {
            foreach ($bad in $wrong) { Disconnect-ExchangeOnline -ConnectionId $bad.ConnectionId -Confirm:$false -ErrorAction SilentlyContinue }
            throw ("Exchange Online connected to tenant {0}, not {1}. The session was closed." -f $wrong[0].TenantID, $tenant)
        }
        $session = @($sessions | Where-Object { ($_.DelegatedOrganization -eq $org) -or ($_.Organization -eq $org) }) | Select-Object -First 1
        if (-not $session) { throw "Connect-ExchangeOnline returned, but no new connected session for $tenant was found (Get-ConnectionInformation)." }
        Write-Warning 'The Exchange session does not report a tenant ID. It was matched by organisation name instead.'
    }
    if ($script:MspState -and $script:MspState.Connections) {
        $null = $script:MspState.Connections.Add('ExchangeOnline')
        if ($session.ConnectionId -and $script:MspState.ExchangeConnections) { $null = $script:MspState.ExchangeConnections.Add([string]$session.ConnectionId) }
    }

    [pscustomobject]@{
        PSTypeName         = 'MspGdap.ExchangeConnection'
        TenantId           = $tenant
        Organization       = $org
        Mode               = $mode
        ConnectionId       = $session.ConnectionId
        State              = $session.State
        UserPrincipalName  = $session.UserPrincipalName
        AppId              = if ($AppOnly) { $AppId } else { $session.AppId }
        TokenExpiryTimeUTC = $session.TokenExpiryTimeUTC
        ModuleVersion      = $moduleVersion
        Experimental       = $false
    }
}
