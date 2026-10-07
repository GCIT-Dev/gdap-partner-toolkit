function Test-MspSetupContext {
    <#
    .SYNOPSIS
        Confirms the setup commands are about to write to the partner tenant the caller named, and nowhere else.
    .DESCRIPTION
        MgGraph transport: Get-MgContext must exist and its TenantId must equal -PartnerTenantId.
        MspGdap transport: Get-MspConfiguration (when available) must name the same partner tenant.
        Returns a step result. A Failed step means the caller must stop before any write.
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][ValidateSet('MgGraph', 'MspGdap')][string]$Transport,
        [Parameter(Mandatory)][string]$PartnerTenantId,
        [string[]]$RequiredScope = @('Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
    )
    $step = 'Check partner tenant connection'
    if ($Transport -eq 'MgGraph') {
        try { $null = Assert-MspModuleAvailable -Name 'Microsoft.Graph.Authentication' -MinimumVersion '2.0.0' }
        catch { return (New-MspStepResult -Step $step -Status Failed -Detail $_.Exception.Message) }
        $context = Get-MgContext
        $connectHint = "Run: Connect-MgGraph -TenantId '$PartnerTenantId' -Scopes $($RequiredScope -join ',')"
        if (-not $context) { return (New-MspStepResult -Step $step -Status Failed -Detail "Not connected to Microsoft Graph. $connectHint") }
        if ([string]$context.TenantId -ne $PartnerTenantId) {
            return (New-MspStepResult -Step $step -Status Failed -Detail ("Connected to tenant {0}, not the partner tenant {1}. Nothing was changed. {2}" -f $context.TenantId, $PartnerTenantId, $connectHint))
        }
        $granted = @($context.Scopes)
        $missing = @($RequiredScope | Where-Object { $granted -notcontains $_ -and $granted -notcontains 'Directory.ReadWrite.All' })
        if ($missing.Count -gt 0 -and $granted.Count -gt 0) {
            return (New-MspStepResult -Step $step -Status Warning -Detail ("Connected to {0} as {1}, but these scopes are not in the session: {2}. Writes may fail. {3}" -f $PartnerTenantId, $context.Account, ($missing -join ', '), $connectHint))
        }
        return (New-MspStepResult -Step $step -Status Passed -Detail ("Microsoft Graph session is in partner tenant {0} as {1}." -f $PartnerTenantId, $context.Account))
    }

    if (Get-Command -Name 'Get-MspConfiguration' -ErrorAction SilentlyContinue) {
        $config = $null
        try { $config = Get-MspConfiguration -ErrorAction Stop } catch { $config = $null }
        if ($config -and $config.PSObject.Properties['PartnerTenantId'] -and $config.PartnerTenantId -and ([string]$config.PartnerTenantId -ne $PartnerTenantId)) {
            return (New-MspStepResult -Step $step -Status Failed -Detail ("MspGdap is configured for partner tenant {0}, not {1}. Nothing was changed." -f $config.PartnerTenantId, $PartnerTenantId))
        }
    }
    return (New-MspStepResult -Step $step -Status Passed -Detail "Using the MspGdap partner tenant token for $PartnerTenantId.")
}
