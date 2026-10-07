function Invoke-MspPartnerGraph {
    <#
    .SYNOPSIS
        Graph request in the PARTNER tenant for the partner app setup commands, over one of two transports.
    .DESCRIPTION
        MgGraph:  Invoke-MgGraphRequest from Microsoft.Graph.Authentication. Used before the partner app exists,
                  after Connect-MgGraph -TenantId <PartnerTenantId> -Scopes ... (see New-MspPartnerApp help).
        MspGdap:  Invoke-MspGraphRequest -PartnerTenant with the technician's MspGdap token (the partner app
                  must already exist and hold Application.ReadWrite.All).
        GET responses are unwrapped and paged. Errors throw.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('MgGraph', 'MspGdap')][string]$Transport,
        [ValidateSet('GET', 'POST', 'PATCH', 'DELETE')][string]$Method = 'GET',
        [Parameter(Mandatory)][string]$Path,
        [object]$Body,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0'
    )
    if ($Transport -eq 'MspGdap') {
        $call = @{ PartnerTenant = $true; Method = $Method; Path = $Path; ApiVersion = $ApiVersion }
        if ($PSBoundParameters.ContainsKey('Body')) { $call['Body'] = $Body }
        return (Invoke-MspGraphCall @call)
    }

    $uri = [string](Resolve-MspGraphUri -Uri $Path -ApiVersion $ApiVersion)
    $request = @{ Method = $Method; Uri = $uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $request['Body'] = if ($Body -is [string]) { $Body } else { ConvertTo-Json -InputObject $Body -Depth 20 -Compress }
        $request['ContentType'] = 'application/json'
    }
    $response = Invoke-MgGraphRequest @request
    if ($Method -ne 'GET') { return $response }

    while ($null -ne $response) {
        $valueProp = $response.PSObject.Properties['value']
        if ($valueProp -and $response.PSObject.Properties['@odata.context']) {
            foreach ($item in @($valueProp.Value)) { if ($null -ne $item) { $item } }
            $next = $response.PSObject.Properties['@odata.nextLink']
            if ($next -and $next.Value) {
                $nextUri = [string](Resolve-MspGraphUri -Uri ([string]$next.Value))
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
                continue
            }
        }
        else {
            $response
        }
        break
    }
}
