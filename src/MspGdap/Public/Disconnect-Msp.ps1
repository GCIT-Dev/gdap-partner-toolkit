function Disconnect-Msp {
    <#
    .SYNOPSIS
        Ends the MspGdap session and clears everything held in memory.
    .DESCRIPTION
        Clears the access token cache, the loaded certificate, any session-only
        PFX password, the selected technician, the tenant ID lookup cache and
        the in-memory configuration (it is re-read from disk on next use).
        Exchange Online, Security and Compliance, Microsoft Graph SDK and
        Microsoft Teams sessions opened through MspGdap (Connect-Msp*) are
        disconnected too. Exchange and Security and Compliance sessions are
        closed by connection ID, so Exchange sessions you opened yourself
        stay open. The Graph SDK and Teams modules hold one session each, so
        those are disconnected whenever MspGdap connected them.

        Refresh tokens stay in the vault. To offboard a technician, remove the
        secret from the vault and revoke their sessions in Entra ID, because a
        password change alone does not revoke a confidential client refresh token.
    .EXAMPLE
        Disconnect-Msp
    #>
    [CmdletBinding()]
    [OutputType([void])]
    param()

    $connections = @($script:MspState.Connections)
    $exchangeIds = @($script:MspState.ExchangeConnections)
    if ($connections -contains 'ExchangeOnline' -or $connections -contains 'SecurityCompliance') {
        if (Get-Command -Name 'Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue) {
            # Close only the sessions MspGdap opened, by connection ID, so other sessions stay open.
            # If a session did not report an ID, fall back to closing every Exchange session.
            $targets = if ($exchangeIds.Count -gt 0) { $exchangeIds } else { @($null) }
            foreach ($id in $targets) {
                try {
                    if ($id) { Disconnect-ExchangeOnline -ConnectionId $id -Confirm:$false -ErrorAction Stop }
                    else { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop }
                }
                catch {
                    Write-Warning "Disconnect-ExchangeOnline failed: $($_.Exception.Message)"
                }
            }
        }
    }
    if ($connections -contains 'MgGraph' -and (Get-Command -Name 'Disconnect-MgGraph' -ErrorAction SilentlyContinue)) {
        try {
            $null = Disconnect-MgGraph -ErrorAction Stop
        }
        catch {
            Write-Warning "Disconnect-MgGraph failed: $($_.Exception.Message)"
        }
    }
    if ($connections -contains 'Teams' -and (Get-Command -Name 'Disconnect-MicrosoftTeams' -ErrorAction SilentlyContinue)) {
        try {
            $null = Disconnect-MicrosoftTeams -ErrorAction Stop
        }
        catch {
            Write-Warning "Disconnect-MicrosoftTeams failed: $($_.Exception.Message)"
        }
    }
    $script:MspState.Connections.Clear()
    $script:MspState.ExchangeConnections.Clear()

    $script:MspTokenCache.Clear()
    if ($script:MspState.Certificate) {
        try { $script:MspState.Certificate.Dispose() } catch { Write-Verbose 'Certificate handle already released.' }
    }
    $script:MspState.Certificate = $null
    $script:MspState.SessionCertificate = $null
    $script:MspState.CertificateSource = $null
    $script:MspState.CertificatePassword = $null
    $script:MspState.CurrentUpn = $null
    $script:MspState.TenantIdCache.Clear()
    $script:MspState.Config = $null
    Write-Verbose 'MspGdap session cleared.'
}
