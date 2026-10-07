function Invoke-MspGraphCall {
    <#
    .SYNOPSIS
        Thin adapter over the core Invoke-MspGraphRequest for the setup, consent, GDAP and Exchange commands.
    .DESCRIPTION
        Keeps every call these commands make to the core request function in one place:
        - builds an absolute https://graph.microsoft.com/<version>/<path> URI and refuses any other host
        - targets either one customer tenant (-TenantId) or the partner tenant (-PartnerTenant), never a default
        - passes -Confirm:$false on writes when the core function supports ShouldProcess, because the
          calling command has already asked ShouldProcess for the whole change
        - unwraps a raw collection response ({ "@odata.context": ..., "value": [...] }) into its items
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant')][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory, ParameterSetName = 'Partner')][switch]$PartnerTenant,
        [ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method = 'GET',
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Path,
        [object]$Body,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0'
    )

    # Resolve-MspGraphUri refuses any host other than https://graph.microsoft.com.
    $uri = [string](Resolve-MspGraphUri -Uri $Path -ApiVersion $ApiVersion)
    $request = @{ Method = $Method; Uri = $uri; ErrorAction = 'Stop' }
    if ($PSCmdlet.ParameterSetName -eq 'Partner') { $request['PartnerTenant'] = $true } else { $request['TenantId'] = $TenantId }
    if ($PSBoundParameters.ContainsKey('Body')) { $request['Body'] = $Body }

    $command = Get-Command -Name 'Invoke-MspGraphRequest' -ErrorAction Stop
    if ($command.Parameters.ContainsKey('ApiVersion')) { $request['ApiVersion'] = $ApiVersion }
    if ($Method -ne 'GET' -and $command.Parameters.ContainsKey('Confirm')) { $request['Confirm'] = $false }

    $response = Invoke-MspGraphRequest @request
    foreach ($item in @($response)) {
        if ($null -eq $item) { continue }
        $valueProp = $item.PSObject.Properties['value']
        $contextProp = $item.PSObject.Properties['@odata.context']
        if ($valueProp -and $contextProp -and ($valueProp.Value -is [System.Array] -or $null -eq $valueProp.Value)) {
            foreach ($inner in @($valueProp.Value)) { if ($null -ne $inner) { $inner } }
        }
        else {
            $item
        }
    }
}
