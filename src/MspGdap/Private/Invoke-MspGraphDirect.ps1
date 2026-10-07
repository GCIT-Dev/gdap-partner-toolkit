function Invoke-MspGraphDirect {
    <#
    .SYNOPSIS
        Sends one Graph request that needs extra headers (for example If-Match) and returns status, headers and body.
    .DESCRIPTION
        Used only where the core request function cannot pass custom headers, such as GDAP access assignment
        PATCH and DELETE, which Microsoft documents as requiring If-Match with the object's ETag.
        The bearer header comes from Get-MspAuthHeader. Nothing from the header is logged.
        The URI goes through Resolve-MspGraphUri, so the token is only ever sent to https://graph.microsoft.com,
        and -Headers can never replace the Authorization header.
        HTTP errors do not throw: the result carries StatusCode and Error so the caller can decide.
    .PARAMETER TenantId
        Customer tenant ID.
    .PARAMETER PartnerTenant
        Send the request to the partner tenant.
    .PARAMETER Method
        HTTP method.
    .PARAMETER Path
        Relative Graph path, or an absolute https://graph.microsoft.com URI.
    .PARAMETER Body
        Request body. Objects are sent as JSON.
    .PARAMETER Headers
        Extra headers such as If-Match. Authorization is ignored.
    .PARAMETER ApiVersion
        v1.0 or beta.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tenant')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Tenant')][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory, ParameterSetName = 'Partner')][switch]$PartnerTenant,
        [Parameter(Mandatory)][ValidateSet('GET', 'POST', 'PATCH', 'PUT', 'DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Path,
        [object]$Body,
        [hashtable]$Headers,
        [ValidateSet('v1.0', 'beta')][string]$ApiVersion = 'v1.0'
    )
    $uri = Resolve-MspGraphUri -Uri $Path -ApiVersion $ApiVersion

    $authParams = @{ Resource = 'https://graph.microsoft.com' }
    if ($PSCmdlet.ParameterSetName -eq 'Partner') { $authParams['PartnerTenant'] = $true } else { $authParams['TenantId'] = $TenantId }
    $requestHeaders = @{}
    if ($Headers) {
        foreach ($key in @($Headers.Keys)) {
            if ([string]$key -ieq 'Authorization') { continue }
            $requestHeaders[$key] = $Headers[$key]
        }
    }
    $auth = Get-MspAuthHeader @authParams
    foreach ($key in @($auth.Keys)) { $requestHeaders[$key] = $auth[$key] }

    $request = @{ Uri = $uri; Method = $Method; Headers = $requestHeaders; UseBasicParsing = $true; ErrorAction = 'Stop' }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        $request['Body'] = if ($Body -is [string]) { $Body } else { ConvertTo-Json -InputObject $Body -Depth 20 -Compress }
        $request['ContentType'] = 'application/json'
    }

    try {
        $response = Invoke-WebRequest @request
        $parsed = $null
        if ($response.Content) { try { $parsed = $response.Content | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null } }
        [pscustomobject]@{ StatusCode = [int]$response.StatusCode; Headers = $response.Headers; Body = $parsed; Error = $null }
    }
    catch {
        $status = Get-MspHttpStatusCode -ErrorRecord $_
        if ($null -eq $status) { throw }
        $message = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        [pscustomobject]@{ StatusCode = $status; Headers = $null; Body = $null; Error = $message }
    }
    finally {
        $requestHeaders = $null
        $request = $null
    }
}
