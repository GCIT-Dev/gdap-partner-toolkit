function Test-MspExchangeAppAccess {
    <#
    .SYNOPSIS
        Checks, read-only, that an automation app has app-only Exchange Online access in a customer.
    .DESCRIPTION
        Same checks as Enable-MspExchangeAppAccess without changing anything: Exchange Online service principal,
        automation app service principal, Exchange.ManageAsApp grant, and each directory role assignment.
        With -TestConnection it also connects app-only with your certificate under the command prefix
        MspGdapTest (so the check can never run against another open Exchange session), runs
        Get-MspGdapTestOrganizationConfig, and disconnects that session.
        A failed test returns Success = false. It is not written as an error.
    .PARAMETER TenantId
        Customer tenant ID or verified domain.
    .PARAMETER AppId
        Application ID of the automation app.
    .PARAMETER Role
        Roles expected on the app. Default 'Exchange Administrator'.
    .PARAMETER AllowGlobalAdministrator
        Accept Global Administrator in -Role.
    .PARAMETER AllowPrivilegedRole
        Accept Compliance Administrator, Security Administrator or Helpdesk Administrator in -Role.
    .PARAMETER TestConnection
        Also connect app-only and run one read-only Exchange cmdlet.
    .PARAMETER CertificateThumbprint
        Thumbprint of the automation app certificate in the local certificate store (Windows).
    .PARAMETER Certificate
        The automation app certificate as an X509Certificate2 object.
    .EXAMPLE
        Test-MspExchangeAppAccess -TenantId 'fabrikam.onmicrosoft.com' -AppId '<AutomationAppId>' -TestConnection -CertificateThumbprint '<Thumbprint>'
    #>
    [CmdletBinding(DefaultParameterSetName = 'NoConnection')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)][Alias('CustomerTenantId')][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$')][string]$AppId,
        [ValidateNotNullOrEmpty()][string[]]$Role = @('Exchange Administrator'),
        [switch]$AllowGlobalAdministrator,
        [switch]$AllowPrivilegedRole,
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')][Parameter(Mandatory, ParameterSetName = 'Certificate')][switch]$TestConnection,
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$CertificateThumbprint,
        [Parameter(Mandatory, ParameterSetName = 'Certificate')][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    process {
        $AppId = $AppId.ToLowerInvariant()
        $steps = New-Object System.Collections.Generic.List[object]
        $info = [ordered]@{ AppId = $AppId; ServicePrincipalId = $null; InitialDomain = $null }
        $tenant = $TenantId
        $finish = { New-MspOperationResult -Operation 'Test-MspExchangeAppAccess' -TenantId $tenant -Target $AppId -Steps $steps.ToArray() -Property $info }

        try {
            $roles = @(Resolve-MspExchangeAppRole -Role $Role -AllowGlobalAdministrator:$AllowGlobalAdministrator -AllowPrivilegedRole:$AllowPrivilegedRole)
            $tenant = Resolve-MspCustomerTenant -TenantId $TenantId
            $state = Get-MspExchangeAppAccessState -TenantId $tenant -AppId $AppId -Role $roles
        }
        catch {
            $steps.Add((New-MspStepResult -Step 'Read current state' -Status Unknown -Detail $_.Exception.Message))
            return (& $finish)
        }

        if ($state.ExchangeServicePrincipal -and $state.ManageAsAppRoleId) { $steps.Add((New-MspStepResult -Step 'Exchange Online service principal' -Status Passed -Detail 'Present and publishes Exchange.ManageAsApp.')) }
        else { $steps.Add((New-MspStepResult -Step 'Exchange Online service principal' -Status Failed -Detail 'Missing, or does not publish Exchange.ManageAsApp.')) }

        if ($state.AppServicePrincipal) {
            $info.ServicePrincipalId = [string]$state.AppServicePrincipal.id
            $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status Passed -Detail "Present ($($state.AppServicePrincipal.id))."))
        }
        else { $steps.Add((New-MspStepResult -Step 'Automation app service principal' -Status Failed -Detail 'Not present. Run Enable-MspExchangeAppAccess.')) }

        if ($state.HasManageAsApp) { $steps.Add((New-MspStepResult -Step 'Exchange.ManageAsApp' -Status Passed -Detail 'Granted.')) }
        else { $steps.Add((New-MspStepResult -Step 'Exchange.ManageAsApp' -Status Failed -Detail 'Not granted.')) }

        foreach ($roleState in $state.Roles) {
            if ($roleState.Assigned) { $steps.Add((New-MspStepResult -Step "Role: $($roleState.DisplayName)" -Status Passed -Detail 'Assigned at directory scope.')) }
            else { $steps.Add((New-MspStepResult -Step "Role: $($roleState.DisplayName)" -Status Failed -Detail 'Not assigned.')) }
        }

        try { $info.InitialDomain = (Get-MspInitialDomain -TenantId $tenant).InitialDomain }
        catch { $steps.Add((New-MspStepResult -Step 'Initial domain' -Status Warning -Detail $_.Exception.Message)) }

        if ($TestConnection) {
            $connection = $null
            try {
                $prefix = 'MspGdapTest'
                $connectParams = @{ TenantId = $tenant; AppOnly = $true; AppId = $AppId; KeepExistingConnections = $true; Prefix = $prefix; CommandName = @('Get-OrganizationConfig') }
                if ($info.InitialDomain) { $connectParams['Organization'] = $info.InitialDomain }
                if ($PSCmdlet.ParameterSetName -eq 'Thumbprint') { $connectParams['CertificateThumbprint'] = $CertificateThumbprint } else { $connectParams['Certificate'] = $Certificate }
                $connection = Connect-MspExchangeOnline @connectParams
                # The prefixed cmdlet exists only in the session opened above, never in another customer's session.
                $org = & "Get-$($prefix)OrganizationConfig" -ErrorAction Stop
                $steps.Add((New-MspStepResult -Step 'App-only connection' -Status Passed -Detail "Connected and read the organisation config ($($org.Name))."))
            }
            catch {
                $steps.Add((New-MspStepResult -Step 'App-only connection' -Status Failed -Detail $_.Exception.Message))
            }
            finally {
                if ($connection -and $connection.ConnectionId) {
                    try { Disconnect-ExchangeOnline -ConnectionId $connection.ConnectionId -Confirm:$false -ErrorAction Stop } catch { Write-Verbose "Disconnect failed: $($_.Exception.Message)" }
                    if ($script:MspState -and $script:MspState.ExchangeConnections) { $null = $script:MspState.ExchangeConnections.Remove([string]$connection.ConnectionId) }
                }
            }
        }
        & $finish
    }
}
